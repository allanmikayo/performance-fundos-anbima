"""Índice de retorno (valorização) de fundos, ponderado por PL, por grupo
(infra x ex-infra) -- frequência DIÁRIA.

ENTRADA:
  1) serie_historica.csv  -- saída do extrai_fundos_anbima.py (cota, PL por data)
  2) classificacao_fundos.csv  -- de/para: chave do fundo -> grupo (infra/ex-infra)
     Se esse arquivo NÃO existir, o script gera uma base FICTÍCIA automaticamente
     (metade infra, metade ex-infra, de forma determinística) só para você ver o
     pipeline rodando ponta a ponta. No computador corporativo, basta substituir
     esse arquivo pela classificação real e rodar de novo.

METODOLOGIA (ver também a aba "metodologia" no Excel de saída):
  Passo 1  retorno de cada fundo i no dia t, a partir da PRÓPRIA cota:
              r_i,t = cota_i,t / cota_i,(t-1) - 1
           (a cota já é líquida de taxas e limpa de captação/resgate)
  Passo 2  peso de cada fundo pelo PL do dia ANTERIOR (t-1), dentro do grupo G:
              w_i,t = PL_i,(t-1) / Σ_{j in G} PL_j,(t-1)
           (peso defasado = carteira montada no início do dia; evita olhar o futuro)
  Passo 3  retorno do grupo no dia t:
              R_G,t = Σ_{i in G} w_i,t * r_i,t
  Passo 4  índice encadeado, base 100 na primeira data:
              I_G,t = I_G,(t-1) * (1 + R_G,t) ;  I_G,0 = 100

  Tratamentos:
   - Alinhamento diário: painel data x fundo; a cota é repetida (forward-fill)
     só DENTRO da janela ativa de cada fundo (entre a 1ª e a última cota dele).
     Em dia sem cota nova o retorno é 0; nada é inventado antes do início nem
     depois do encerramento (evita viés de sobrevivência artificial).
   - Fundo novo entra quando tem 2 cotas (precisa de t-1 e t).
   - Outliers: retorno diário com |r| acima de LIMITE_OUTLIER é tratado como
     erro de dado (cota parada/salto) e EXCLUÍDO do índice naquele dia; os pesos
     do grupo são renormalizados entre os fundos válidos. Tudo fica marcado.

  ATENÇÃO (risco metodológico): este índice assume que a cota é "total return".
  Fundos incentivados de infraestrutura costumam DISTRIBUIR rendimentos; na data
  da distribuição a cota cai sem perda real, o que SUBESTIMA o retorno desses
  fundos. Se for o caso, a comparação infra x ex-infra precisa de uma fonte de
  distribuições para reconstruir o retorno total. Ver conversa.

SAÍDA: indice_retorno_fundos.xlsx, com as abas:
  - indice_diario        curvas dos índices (infra, ex-infra) + retorno diário de cada grupo
  - retornos_por_fundo   trilha completa: cota, cota_ant, retorno, PL, PL_ant, peso,
                         contribuição -- para você conferir CADA número à mão
  - resumo               retorno acumulado, anualizado e volatilidade por grupo
  - metodologia          as fórmulas acima, por extenso
"""
from __future__ import annotations

# ╔══════════════════════════════════════════════════════════════════╗
# ║                         CONFIGURAÇÃO                              ║
# ╚══════════════════════════════════════════════════════════════════╝
from pathlib import Path

PASTA_SCRIPT = Path(__file__).resolve().parent

# >>> ALTERAR (1) -- SÉRIE HISTÓRICA (entrada)
# Caminho do serie_historica.csv. Se deixar None, o script procura o mais
# recente em output/<data_hora>/serie_historica.csv.
ARQUIVO_SERIE = None

# >>> ALTERAR (2) -- CLASSIFICAÇÃO infra / ex-infra
# CSV com duas colunas: a chave do fundo e o grupo. Se o arquivo não existir,
# o script cria uma base FICTÍCIA com este mesmo nome para o teste rodar.
ARQUIVO_CLASSIFICACAO = PASTA_SCRIPT / "classificacao_fundos.csv"
COL_CHAVE_CLASSIF = "chave"      # nome da coluna de chave no arquivo de classificação
COL_GRUPO = "grupo"              # nome da coluna de grupo (valores: infra / ex-infra)

# >>> ALTERAR (3) -- CHAVE de ligação série <-> classificação
# Qual coluna da série casa com a chave da classificação.
#   "codigo_classe_consultado" (padrão) -- código ANBIMA da classe
#   "cnpj_input"                         -- CNPJ (se a série foi extraída por CNPJ)
CHAVE_SERIE = "codigo_classe_consultado"

# Rótulos dos dois grupos (têm que bater com os valores da coluna grupo)
GRUPO_INFRA = "infra"
GRUPO_EX = "ex-infra"

# Base do índice e limite de outlier de retorno diário (0.20 = 20%)
INDICE_BASE = 100.0
LIMITE_OUTLIER = 0.20

PASTA_SAIDA = PASTA_SCRIPT / "analise_output"
SEED_FICTICIA = 42
# ════════════════════ FIM DA CONFIGURAÇÃO ═══════════════════════════


import re
import sys
from datetime import datetime

import numpy as np
import pandas as pd

LOG: list[str] = []


def log(msg: str) -> None:
    linha = f"[{datetime.now():%H:%M:%S}] {msg}"
    print(linha, flush=True)
    LOG.append(linha)


# --------------------------------------------------------------------------
# Passo 0 -- localizar e ler a série histórica
# --------------------------------------------------------------------------
def achar_serie() -> Path:
    if ARQUIVO_SERIE:
        p = Path(ARQUIVO_SERIE)
        if not p.exists():
            sys.exit(f"ARQUIVO_SERIE não encontrado: {p}")
        return p
    candidatos = sorted(PASTA_SCRIPT.glob("output/*/serie_historica.csv"))
    if not candidatos:
        sys.exit("Não achei nenhum output/*/serie_historica.csv. Rode a extração "
                 "antes, ou preencha ARQUIVO_SERIE na CONFIGURAÇÃO.")
    return candidatos[-1]


def ler_serie(p: Path) -> pd.DataFrame:
    log(f"Lendo série: {p}")
    # dtype=str preserva CNPJ/códigos; converto só cota e PL para número.
    df = pd.read_csv(p, sep=";", dtype=str, encoding="utf-8-sig", keep_default_na=False)
    obrig = [CHAVE_SERIE, "data_competencia", "valor_cota", "valor_patrimonio_liquido"]
    faltando = [c for c in obrig if c not in df.columns]
    if faltando:
        sys.exit(f"Colunas ausentes na série: {faltando}. Colunas achadas: {list(df.columns)}")

    def num(s: pd.Series) -> pd.Series:
        # CSV veio com separador de milhar "." e decimal "," (padrão BR do extrator)
        return pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                             errors="coerce")

    out = pd.DataFrame({
        "chave": df[CHAVE_SERIE].str.strip(),
        "data": pd.to_datetime(df["data_competencia"], errors="coerce"),
        "cota": num(df["valor_cota"]),
        "pl": num(df["valor_patrimonio_liquido"]),
    })
    antes = len(out)
    out = out.dropna(subset=["chave", "data", "cota", "pl"])
    out = out[(out["cota"] > 0) & (out["pl"] > 0)]
    out = out[out["chave"] != ""]
    out = out.drop_duplicates(["chave", "data"]).sort_values(["chave", "data"])
    log(f"  linhas válidas: {len(out)} (de {antes}); fundos: {out['chave'].nunique()}; "
        f"datas: {out['data'].min():%Y-%m-%d} a {out['data'].max():%Y-%m-%d}")
    return out


# --------------------------------------------------------------------------
# Passo 0b -- classificação infra / ex-infra (real ou fictícia)
# --------------------------------------------------------------------------
def gerar_classificacao_ficticia(chaves: list[str]) -> pd.DataFrame:
    # determinística: metade infra, metade ex-infra, pelo hash estável da chave
    rng = np.random.default_rng(SEED_FICTICIA)
    ordem = sorted(chaves)
    sorteio = rng.integers(0, 2, size=len(ordem))
    df = pd.DataFrame({COL_CHAVE_CLASSIF: ordem,
                       COL_GRUPO: np.where(sorteio == 0, GRUPO_INFRA, GRUPO_EX)})
    df.to_csv(ARQUIVO_CLASSIFICACAO, index=False, sep=";", encoding="utf-8-sig")
    log(f"  !!! BASE FICTÍCIA gerada em {ARQUIVO_CLASSIFICACAO.name} "
        f"({(df[COL_GRUPO]==GRUPO_INFRA).sum()} infra / {(df[COL_GRUPO]==GRUPO_EX).sum()} ex-infra).")
    log("  !!! SUBSTITUA esse arquivo pela classificação REAL e rode de novo.")
    return df


def ler_classificacao(chaves_serie: list[str]) -> pd.DataFrame:
    if not Path(ARQUIVO_CLASSIFICACAO).exists():
        log(f"Classificação não encontrada -- gerando base fictícia para teste.")
        cls = gerar_classificacao_ficticia(chaves_serie)
    else:
        log(f"Lendo classificação: {ARQUIVO_CLASSIFICACAO}")
        cab = Path(ARQUIVO_CLASSIFICACAO).read_text(encoding="utf-8-sig").splitlines()[0]
        sep = ";" if ";" in cab else ","
        cls = pd.read_csv(ARQUIVO_CLASSIFICACAO, dtype=str, sep=sep, encoding="utf-8-sig")
        if COL_CHAVE_CLASSIF not in cls or COL_GRUPO not in cls:
            sys.exit(f"Classificação precisa das colunas '{COL_CHAVE_CLASSIF}' e '{COL_GRUPO}'. "
                     f"Achei: {list(cls.columns)}")
    cls = cls.rename(columns={COL_CHAVE_CLASSIF: "chave", COL_GRUPO: "grupo"})
    cls["chave"] = cls["chave"].astype(str).str.strip()
    cls["grupo"] = cls["grupo"].astype(str).str.strip()
    cls = cls[cls["chave"] != ""].drop_duplicates("chave")
    # só os dois grupos esperados
    validos = cls["grupo"].isin([GRUPO_INFRA, GRUPO_EX])
    if (~validos).any():
        log(f"  {(~validos).sum()} linha(s) com grupo fora de "
            f"('{GRUPO_INFRA}','{GRUPO_EX}') ignoradas: {cls.loc[~validos,'grupo'].unique()[:5]}")
    return cls[validos]


# --------------------------------------------------------------------------
# Passo 1 -- painel diário (data x fundo) de cota e PL, com ffill na janela ativa
# --------------------------------------------------------------------------
def painel_diario(serie: pd.DataFrame, chaves_grupo: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    s = serie[serie["chave"].isin(chaves_grupo)]
    cal = pd.Index(sorted(s["data"].unique()), name="data")        # calendário = dias com dado
    cota = s.pivot(index="data", columns="chave", values="cota").reindex(cal)
    pl = s.pivot(index="data", columns="chave", values="pl").reindex(cal)

    # forward-fill só DENTRO da janela ativa de cada fundo (1ª a última cota).
    def ffill_janela(m: pd.DataFrame) -> pd.DataFrame:
        cheio = m.ffill()
        for c in m.columns:
            col = m[c]
            fv, lv = col.first_valid_index(), col.last_valid_index()
            if fv is None:
                continue
            mascara = (cheio.index < fv) | (cheio.index > lv)       # fora da janela -> NaN
            cheio.loc[mascara, c] = np.nan
        return cheio

    return ffill_janela(cota), ffill_janela(pl)


# --------------------------------------------------------------------------
# Passos 2 a 4 em Python -- SÓ para conferência no console/log.
# O Excel NÃO usa estes números: lá tudo é fórmula viva (ver exportar_excel).
# --------------------------------------------------------------------------
def indice_do_grupo(cota: pd.DataFrame, pl: pd.DataFrame) -> pd.DataFrame:
    ret = cota / cota.shift(1) - 1.0                    # Passo 1
    pl_ant = pl.shift(1)                                # insumo do peso
    ret_val = ret.mask(ret.abs() > LIMITE_OUTLIER)      # tira outliers
    valido = ret_val.notna() & pl_ant.notna() & (pl_ant > 0)
    pl_peso = pl_ant.where(valido)
    soma_pl = pl_peso.sum(axis=1)
    peso = pl_peso.div(soma_pl, axis=0)                 # Passo 2
    ret_grupo = (peso * ret_val.where(valido)).sum(axis=1).where(soma_pl > 0)  # Passo 3
    idx = INDICE_BASE * (1.0 + ret_grupo.fillna(0.0)).cumprod()               # Passo 4
    return pd.DataFrame({"retorno_grupo": ret_grupo, "indice": idx,
                         "n_fundos": valido.sum(axis=1).astype(int)})


def estatisticas(ind: pd.DataFrame, nome: str) -> dict:
    s = ind["indice"].dropna()
    r = ind["retorno_grupo"].dropna()
    if len(s) < 2:
        return {"grupo": nome, "obs": len(s)}
    dias = (s.index[-1] - s.index[0]).days or 1
    acum = s.iloc[-1] / s.iloc[0] - 1
    return {"grupo": nome, "inicio": s.index[0].date(), "fim": s.index[-1].date(),
            "obs_dias": len(s), "fundos_medio": round(ind["n_fundos"].mean(), 1),
            "retorno_acumulado": acum, "retorno_anualizado": (1 + acum) ** (365.0 / dias) - 1,
            "vol_anualizada": r.std() * np.sqrt(252)}


# --------------------------------------------------------------------------
# Painel diário em formato longo (data, grupo, chave, cota, pl) -- os
# ÚNICOS valores crus do Excel; o resto vira fórmula.
# --------------------------------------------------------------------------
def painel_longo(serie: pd.DataFrame, chaves: list[str], grupo: str) -> pd.DataFrame:
    cota, pl = painel_diario(serie, chaves)
    lon = (cota.stack(future_stack=True).rename("cota").to_frame()
           .join(pl.stack(future_stack=True).rename("pl")).reset_index())
    lon = lon.dropna(subset=["cota"])        # só dias dentro da janela ativa
    lon["grupo"] = grupo
    return lon[["data", "grupo", "chave", "cota", "pl"]]


# --------------------------------------------------------------------------
# SAÍDA: Excel com FÓRMULAS vivas (nada de valor colado nas colunas de cálculo)
# --------------------------------------------------------------------------
def exportar_excel(panel: pd.DataFrame, saida: Path) -> None:
    from openpyxl import Workbook

    panel = panel.sort_values(["chave", "data"]).reset_index(drop=True)
    n = len(panel)
    fim = n + 1                               # última linha de dados (linha 1 = cabeçalho)
    PCT = "0.0000%"; NUM = "0.000000"; DT = "yyyy-mm-dd"

    wb = Workbook()
    # ---- Aba 1: retornos_por_fundo (valores: data/grupo/chave/cota/pl; resto fórmula) ----
    ws = wb.active
    ws.title = "retornos_por_fundo"
    cols = ["data", "grupo", "chave", "cota", "pl", "cota_ant", "pl_ant", "retorno",
            "outlier", "valido", "pl_ant_valido", "soma_pl_dia_grupo", "peso", "contribuicao"]
    ws.append(cols)
    for i, row in enumerate(panel.itertuples(index=False), start=2):
        ws.cell(i, 1, row.data.to_pydatetime()).number_format = DT
        ws.cell(i, 2, row.grupo)
        ws.cell(i, 3, row.chave)
        ws.cell(i, 4, float(row.cota))
        ws.cell(i, 5, float(row.pl))
        # cota_ant / pl_ant: valor do dia anterior DO MESMO fundo (fórmula de lookup pela linha de cima)
        ws.cell(i, 6, f'=IF($C{i}=$C{i-1},$D{i-1},"")')
        ws.cell(i, 7, f'=IF($C{i}=$C{i-1},$E{i-1},"")')
        # retorno = cota/cota_ant - 1
        ws.cell(i, 8, f'=IFERROR($D{i}/$F{i}-1,"")').number_format = PCT
        # outlier = |retorno| > limite
        ws.cell(i, 9, f'=IF($H{i}="","",ABS($H{i})>{LIMITE_OUTLIER})')
        # valido = tem retorno, cota_ant>0, pl_ant>0 e não é outlier
        ws.cell(i, 10, f'=IF($H{i}="",FALSE,AND($F{i}>0,$G{i}>0,NOT($I{i})))')
        # pl_ant_valido = pl_ant se válido, senão 0 (entra no denominador do peso)
        ws.cell(i, 11, f'=IF($J{i},$G{i},0)')
        # soma do pl_ant válido do mesmo dia e grupo (denominador do peso)
        ws.cell(i, 12, f'=SUMIFS($K$2:$K${fim},$A$2:$A${fim},$A{i},$B$2:$B${fim},$B{i})')
        # peso = pl_ant_valido / soma_dia  (Passo 2)
        ws.cell(i, 13, f'=IF($J{i},$K{i}/$L{i},"")').number_format = PCT
        # contribuicao = peso * retorno
        ws.cell(i, 14, f'=IF($J{i},$M{i}*$H{i},"")').number_format = PCT

    # ---- Aba 2: indice_diario (retorno do grupo, índice encadeado, nº de fundos) ----
    datas = sorted(panel["data"].unique())
    wi = wb.create_sheet("indice_diario")
    wi.append(["data", "retorno_infra", "indice_infra", "n_fundos_infra",
               "retorno_ex_infra", "indice_ex_infra", "n_fundos_ex_infra"])
    SH = "retornos_por_fundo"
    for j, d in enumerate(datas, start=2):
        wi.cell(j, 1, pd.Timestamp(d).to_pydatetime()).number_format = DT
        for (gname, cret, cidx, cnf) in [(GRUPO_INFRA, 2, 3, 4), (GRUPO_EX, 5, 6, 7)]:
            # Passo 3: retorno do grupo no dia = soma das contribuições
            wi.cell(j, cret, f'=SUMIFS({SH}!$N$2:$N${fim},{SH}!$A$2:$A${fim},$A{j},'
                             f'{SH}!$B$2:$B${fim},"{gname}")').number_format = PCT
            cl = chr(ord("A") + cret - 1)   # coluna da letra do retorno (B ou E)
            ci = chr(ord("A") + cidx - 1)   # coluna da letra do índice (C ou F)
            # Passo 4: índice base 100 no 1º dia; depois anterior * (1 + retorno)
            wi.cell(j, cidx, "=100" if j == 2 else f"={ci}{j-1}*(1+{cl}{j})").number_format = "0.00"
            # nº de fundos válidos no dia/grupo
            wi.cell(j, cnf, f'=COUNTIFS({SH}!$K$2:$K${fim},">0",{SH}!$A$2:$A${fim},$A{j},'
                            f'{SH}!$B$2:$B${fim},"{gname}")')
    ultcal = len(datas) + 1

    # ---- Aba 3: resumo (também por fórmula, referenciando indice_diario) ----
    wr = wb.create_sheet("resumo")
    wr.append(["grupo", "inicio", "fim", "retorno_acumulado", "retorno_anualizado",
               "vol_anualizada", "fundos_medio"])
    # (nome, linha, col retorno, col indice, col n_fundos na aba indice_diario)
    for nome, linha, cret, cidx, cnf in [(GRUPO_INFRA, 2, "B", "C", "D"),
                                         (GRUPO_EX, 3, "E", "F", "G")]:
        wr.cell(linha, 1, nome)
        wr.cell(linha, 2, "=indice_diario!$A$2").number_format = DT
        wr.cell(linha, 3, f"=indice_diario!$A${ultcal}").number_format = DT
        wr.cell(linha, 4, f"=indice_diario!{cidx}{ultcal}/{INDICE_BASE}-1").number_format = PCT
        wr.cell(linha, 5, f"=(1+D{linha})^(365/(C{linha}-B{linha}))-1").number_format = PCT
        wr.cell(linha, 6, f"=STDEV(indice_diario!{cret}2:{cret}{ultcal})*SQRT(252)").number_format = PCT
        wr.cell(linha, 7, f"=AVERAGE(indice_diario!{cnf}2:{cnf}{ultcal})").number_format = "0.0"

    # ---- Aba 4: metodologia ----
    wm = wb.create_sheet("metodologia")
    for linha in [
        ["Todas as colunas de calculo sao FORMULAS -- abra no Excel para recalcular."],
        ["Passo 1 - retorno por fundo:  retorno = cota / cota_ant - 1"],
        ["Passo 2 - peso pelo PL do dia anterior:  peso = pl_ant_valido / soma_pl_dia_grupo"],
        ["Passo 3 - retorno do grupo:  retorno_grupo = soma das contribuicoes (peso * retorno)"],
        ["Passo 4 - indice base 100:  indice_t = indice_(t-1) * (1 + retorno_grupo_t)"],
        [f"Outlier: |retorno| > {LIMITE_OUTLIER:.0%} marca 'outlier'=VERDADEIRO e zera 'valido' -> sai do dia."],
        ["Cota repetida (ffill) so dentro da janela ativa do fundo; fundo entra quando tem 2 cotas."],
        ["ATENCAO: assume cota = total return. Fundos que distribuem (incentivados de infra) tem a"],
        ["cota reduzida na distribuicao, o que SUBESTIMA o retorno -- validar antes de concluir."],
    ]:
        wm.append(linha)

    wb.save(saida)


# --------------------------------------------------------------------------
def main() -> None:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    serie = ler_serie(achar_serie())
    cls = ler_classificacao(sorted(serie["chave"].unique()))

    serie = serie.merge(cls, on="chave", how="inner")
    cob = serie["chave"].nunique()
    log(f"Fundos com classificação E série: {cob}")
    log(f"  por grupo: {serie.groupby('grupo')['chave'].nunique().to_dict()}")
    if cob == 0:
        sys.exit("Nenhum fundo casou entre série e classificação -- confira CHAVE_SERIE "
                 "e a coluna de chave do arquivo de classificação.")

    paineis, stats = [], []
    for nome in (GRUPO_INFRA, GRUPO_EX):
        chaves = serie.loc[serie["grupo"] == nome, "chave"].unique().tolist()
        if not chaves:
            log(f"Grupo '{nome}' sem fundos -- pulado.")
            continue
        log(f"Grupo '{nome}': {len(chaves)} fundos")
        cota, pl = painel_diario(serie, chaves)
        stats.append(estatisticas(indice_do_grupo(cota, pl), nome))   # só p/ conferência
        paineis.append(painel_longo(serie, chaves, nome))

    panel = pd.concat(paineis, ignore_index=True)
    linhas = len(panel)
    log(f"Painel diário (linhas de fundo x dia): {linhas}")
    if linhas > 1_048_575:
        sys.exit(f"Painel tem {linhas} linhas, acima do limite do Excel (~1,05M). "
                 "Reduza o período ou a lista de fundos.")

    saida = PASTA_SAIDA / "indice_retorno_fundos.xlsx"
    exportar_excel(panel, saida)
    log(f"Excel salvo (fórmulas vivas): {saida}")
    log("  Abra no Excel para recalcular -- visualizadores que não calculam mostram zero/vazio.")

    df_stats = pd.DataFrame(stats)
    if not df_stats.empty:
        log("Conferência (Python, NÃO é o que vai no Excel):\n" + df_stats.to_string(index=False))
    (PASTA_SAIDA / "log_analise.txt").write_text("\n".join(LOG), encoding="utf-8")


if __name__ == "__main__":
    main()
