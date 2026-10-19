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
# Passos 2 a 4 -- retornos, pesos, índice do grupo
# --------------------------------------------------------------------------
def indice_do_grupo(cota: pd.DataFrame, pl: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    # Passo 1: retorno diário por fundo a partir da própria cota
    ret = cota / cota.shift(1) - 1.0
    # Passo 2 (insumo): PL do dia anterior = peso bruto
    pl_ant = pl.shift(1)

    # marca outliers e os remove do índice (viram NaN e não entram no peso)
    outlier = ret.abs() > LIMITE_OUTLIER
    ret_val = ret.mask(outlier)

    # fundo entra no dia t só se tem retorno válido E PL_ant válido
    valido = ret_val.notna() & pl_ant.notna() & (pl_ant > 0)
    pl_peso = pl_ant.where(valido)                      # PL_ant só onde o fundo é válido no dia
    soma_pl = pl_peso.sum(axis=1)                       # Σ PL_ant do grupo, por dia
    peso = pl_peso.div(soma_pl, axis=0)                 # Passo 2: w_i,t
    contrib = peso * ret_val.where(valido)              # w_i,t * r_i,t
    ret_grupo = contrib.sum(axis=1)                     # Passo 3: R_G,t
    ret_grupo = ret_grupo.where(soma_pl > 0)            # dia sem nenhum fundo válido -> NaN

    # Passo 4: índice encadeado base 100 a partir do 1º dia com retorno
    r = ret_grupo.fillna(0.0)
    idx = INDICE_BASE * (1.0 + r).cumprod()
    # antes do 1º retorno válido, índice = base (ou NaN se preferir); deixo base.
    resumo = pd.DataFrame({
        "retorno_grupo": ret_grupo,
        "indice": idx,
        "n_fundos": valido.sum(axis=1).astype(int),
    })

    # trilha de auditoria (long): cada número reconstruível
    def long(m, nome):
        return m.stack(future_stack=True).rename(nome)
    audit = pd.concat([
        long(cota, "cota"), long(cota.shift(1), "cota_ant"),
        long(ret, "retorno_bruto"), long(outlier, "outlier"),
        long(pl, "pl"), long(pl_ant, "pl_ant"),
        long(peso, "peso"), long(contrib, "contribuicao"),
    ], axis=1).reset_index()
    audit = audit[audit[["cota", "pl"]].notna().any(axis=1)]
    return resumo, audit


# --------------------------------------------------------------------------
# Resumo estatístico por grupo
# --------------------------------------------------------------------------
def estatisticas(ind: pd.DataFrame, nome: str) -> dict:
    serie_idx = ind["indice"].dropna()
    r = ind["retorno_grupo"].dropna()
    if len(serie_idx) < 2:
        return {"grupo": nome, "obs": len(serie_idx)}
    dias = (serie_idx.index[-1] - serie_idx.index[0]).days or 1
    acum = serie_idx.iloc[-1] / serie_idx.iloc[0] - 1
    anualizado = (1 + acum) ** (365.0 / dias) - 1
    vol_anual = r.std() * np.sqrt(252)
    return {
        "grupo": nome, "inicio": serie_idx.index[0].date(), "fim": serie_idx.index[-1].date(),
        "obs_dias": len(serie_idx), "fundos_medio": round(ind["n_fundos"].mean(), 1),
        "retorno_acumulado": acum, "retorno_anualizado": anualizado, "vol_anualizada": vol_anual,
    }


# --------------------------------------------------------------------------
def main() -> None:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    serie = ler_serie(achar_serie())
    cls = ler_classificacao(sorted(serie["chave"].unique()))

    serie = serie.merge(cls, on="chave", how="inner")
    cob = serie["chave"].nunique()
    log(f"Fundos com classificação E série: {cob}")
    por_grupo = serie.groupby("grupo")["chave"].nunique().to_dict()
    log(f"  por grupo: {por_grupo}")
    if cob == 0:
        sys.exit("Nenhum fundo casou entre série e classificação -- confira CHAVE_SERIE "
                 "e a coluna de chave do arquivo de classificação.")

    indices, audits, stats = {}, [], []
    for nome in (GRUPO_INFRA, GRUPO_EX):
        chaves = serie.loc[serie["grupo"] == nome, "chave"].unique().tolist()
        if not chaves:
            log(f"Grupo '{nome}' sem fundos -- pulado.")
            continue
        log(f"Grupo '{nome}': {len(chaves)} fundos")
        cota, pl = painel_diario(serie, chaves)
        ind, audit = indice_do_grupo(cota, pl)
        audit.insert(1, "grupo", nome)
        indices[nome] = ind
        audits.append(audit)
        stats.append(estatisticas(ind, nome))

    # monta aba do índice diário (uma coluna por grupo)
    ind_diario = pd.DataFrame(index=sorted(serie["data"].unique()))
    ind_diario.index.name = "data"
    for nome, ind in indices.items():
        suf = nome.replace("-", "_")
        ind_diario[f"indice_{suf}"] = ind["indice"]
        ind_diario[f"retorno_{suf}"] = ind["retorno_grupo"]
        ind_diario[f"n_fundos_{suf}"] = ind["n_fundos"]
    ind_diario = ind_diario.reset_index()

    audit_all = pd.concat(audits, ignore_index=True) if audits else pd.DataFrame()
    df_stats = pd.DataFrame(stats)

    metodologia = pd.DataFrame({"metodologia": [
        "Passo 1 - retorno por fundo: r_i,t = cota_i,t / cota_i,(t-1) - 1",
        "Passo 2 - peso pelo PL do dia anterior: w_i,t = PL_i,(t-1) / soma(PL_(t-1) do grupo)",
        "Passo 3 - retorno do grupo: R_G,t = soma_i( w_i,t * r_i,t )",
        "Passo 4 - indice base 100: I_G,t = I_G,(t-1) * (1 + R_G,t), I_G,0 = 100",
        f"Outlier: |r| > {LIMITE_OUTLIER:.0%} e excluido do dia (coluna 'outlier'); pesos renormalizados.",
        "Cota repetida (ffill) so dentro da janela ativa do fundo; fundo entra com 2 cotas.",
        "ATENCAO: assume cota = total return. Fundos que distribuem (ex.: incentivados de infra) "
        "tem a cota reduzida na distribuicao, o que SUBESTIMA o retorno -- validar antes de concluir.",
        "Confira qualquer numero na aba 'retornos_por_fundo': contribuicao = peso * retorno; "
        "o retorno do grupo no dia e a soma das contribuicoes.",
    ]})

    saida = PASTA_SAIDA / "indice_retorno_fundos.xlsx"
    with pd.ExcelWriter(saida, engine="openpyxl") as w:
        ind_diario.to_excel(w, sheet_name="indice_diario", index=False)
        df_stats.to_excel(w, sheet_name="resumo", index=False)
        metodologia.to_excel(w, sheet_name="metodologia", index=False)
        if not audit_all.empty:
            # limite do Excel: se passar de ~1M linhas, grava a trilha em CSV à parte
            if len(audit_all) > 1_048_575:
                audit_all.to_csv(PASTA_SAIDA / "retornos_por_fundo.csv", index=False,
                                 sep=";", decimal=",", encoding="utf-8-sig")
                log(f"  trilha com {len(audit_all)} linhas -> CSV à parte (acima do limite do Excel)")
            else:
                audit_all.to_excel(w, sheet_name="retornos_por_fundo", index=False)
    log(f"Excel salvo: {saida}")
    if not df_stats.empty:
        log("Resumo:\n" + df_stats.to_string(index=False))
    (PASTA_SAIDA / "log_analise.txt").write_text("\n".join(LOG), encoding="utf-8")


if __name__ == "__main__":
    main()
