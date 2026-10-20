"""Retorno DIÁRIO da cota, por grupo (padrão: infra), ponderado por PL.

Entrega: uma série com DATA e o % de PERFORMANCE do dia (retorno do grupo,
ponderado pelo PL). Usa SÓ as cotas e o PL do serie_historica.csv -- nada de
captação aqui (esse dado você cruza de outra base depois).

Três filtros/tratamentos pedidos:
  1) GRUPO: mantém só os fundos do grupo alvo (padrão 'infra'), pela
     classificacao_fundos.csv.
  2) REPORT DIÁRIO: mantém só fundos cuja mediana do intervalo entre cotas
     é <= MAX_GAP_MEDIANO_DIAS (fundo diário tem mediana ~1; mensal ~30).
  3) FREQUÊNCIA/BURACOS desiguais: o retorno é calculado entre observações
     CONSECUTIVAS da cota. Só vale como retorno "diário" se o intervalo até a
     cota anterior for <= MAX_GAP_RETORNO_DIAS (cobre fim de semana/feriado).
     Gap maior é descartado. Comprimentos diferentes e entradas/saídas se
     resolvem porque o PESO é renormalizado a cada dia entre os fundos presentes.

FÓRMULAS (Passo a passo; no Excel tudo é fórmula viva):
  retorno do fundo i no dia t:   r_i,t = cota_i,t / cota_i,(obs anterior) - 1
  peso pelo PL da obs anterior:  w_i,t = PL_i,ant / soma(PL_ant dos válidos no dia t)
  performance do grupo no dia t: R_t = soma_i( w_i,t * r_i,t )
  índice base 100 (cortesia):    I_t = I_(t-1) * (1 + R_t)

SAÍDA: retorno_diario_cota.xlsx
  - performance_diaria   data | retorno_pct | n_fundos | indice_base100   <-- o que você pediu
  - retornos_por_fundo   trilha por fundo x dia (cota, cota_ant, gap, retorno, peso, contribuição)
  - selecao_fundos       cada fundo do grupo com nº de obs, gap mediano e a regra 'entra' (fórmula)
  - metodologia
"""
from __future__ import annotations

# ╔══════════════════════════════════════════════════════════════════╗
# ║                         CONFIGURAÇÃO                              ║
# ╚══════════════════════════════════════════════════════════════════╝
from pathlib import Path

PASTA_SCRIPT = Path(__file__).resolve().parent

# >>> ALTERAR (1) -- série histórica (None = acha o mais recente em output/)
ARQUIVO_SERIE = None

# >>> ALTERAR (2) -- classificação infra/ex-infra (mesmo arquivo da outra análise)
ARQUIVO_CLASSIFICACAO = PASTA_SCRIPT / "classificacao_fundos.csv"
COL_CHAVE_CLASSIF = "chave"
COL_GRUPO = "grupo"
CHAVE_SERIE = "codigo_classe_consultado"   # ou "cnpj_input"

# >>> ALTERAR (3) -- grupo a analisar
GRUPO_ALVO = "infra"

# Regras de seleção/limpeza
MAX_GAP_MEDIANO_DIAS = 4     # fundo é "diário" se a mediana do intervalo entre cotas <= isto
MIN_OBS = 30                 # ignora séries curtas demais (menos de N cotas no período)
MAX_GAP_RETORNO_DIAS = 5     # retorno só vale se a cota anterior está a <= N dias (fim de semana/feriado)
MIN_FUNDOS_DIA = 1           # nº mínimo de fundos para o dia valer (informativo)

INDICE_BASE = 100.0
PASTA_SAIDA = PASTA_SCRIPT / "analise_output"
# ════════════════════ FIM DA CONFIGURAÇÃO ═══════════════════════════


import sys
from datetime import datetime

import numpy as np
import pandas as pd

LOG: list[str] = []


def log(msg: str) -> None:
    linha = f"[{datetime.now():%H:%M:%S}] {msg}"
    print(linha, flush=True)
    LOG.append(linha)


def achar_serie() -> Path:
    if ARQUIVO_SERIE:
        p = Path(ARQUIVO_SERIE)
        if not p.exists():
            sys.exit(f"ARQUIVO_SERIE não encontrado: {p}")
        return p
    cand = sorted(PASTA_SCRIPT.glob("output/*/serie_historica.csv"))
    if not cand:
        sys.exit("Não achei output/*/serie_historica.csv. Rode a extração ou preencha ARQUIVO_SERIE.")
    return cand[-1]


def ler_serie(p: Path) -> pd.DataFrame:
    log(f"Lendo série: {p}")
    df = pd.read_csv(p, sep=";", dtype=str, encoding="utf-8-sig", keep_default_na=False)
    for c in [CHAVE_SERIE, "data_competencia", "valor_cota", "valor_patrimonio_liquido"]:
        if c not in df.columns:
            sys.exit(f"Coluna ausente na série: {c}. Achei: {list(df.columns)}")

    def num(s):
        return pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                             errors="coerce")

    out = pd.DataFrame({
        "chave": df[CHAVE_SERIE].str.strip(),
        "data": pd.to_datetime(df["data_competencia"], errors="coerce"),
        "cota": num(df["valor_cota"]),
        "pl": num(df["valor_patrimonio_liquido"]),
    }).dropna()
    out = out[(out["cota"] > 0) & (out["pl"] > 0) & (out["chave"] != "")]
    out = out.drop_duplicates(["chave", "data"]).sort_values(["chave", "data"])
    log(f"  linhas válidas: {len(out)}; fundos: {out['chave'].nunique()}; "
        f"datas: {out['data'].min():%Y-%m-%d} a {out['data'].max():%Y-%m-%d}")
    return out


def ler_classificacao() -> pd.DataFrame:
    if not Path(ARQUIVO_CLASSIFICACAO).exists():
        sys.exit(f"Classificação não encontrada: {ARQUIVO_CLASSIFICACAO}. "
                 "Rode a outra análise uma vez para gerar a fictícia, ou crie a real.")
    cab = Path(ARQUIVO_CLASSIFICACAO).read_text(encoding="utf-8-sig").splitlines()[0]
    sep = ";" if ";" in cab else ","
    cls = pd.read_csv(ARQUIVO_CLASSIFICACAO, dtype=str, sep=sep, encoding="utf-8-sig")
    if COL_CHAVE_CLASSIF not in cls or COL_GRUPO not in cls:
        sys.exit(f"Classificação precisa das colunas '{COL_CHAVE_CLASSIF}' e '{COL_GRUPO}'.")
    cls = cls.rename(columns={COL_CHAVE_CLASSIF: "chave", COL_GRUPO: "grupo"})
    cls["chave"] = cls["chave"].astype(str).str.strip()
    cls["grupo"] = cls["grupo"].astype(str).str.strip()
    return cls[cls["chave"] != ""].drop_duplicates("chave")


def diagnostico_frequencia(serie: pd.DataFrame) -> pd.DataFrame:
    """Por fundo: nº de obs e mediana do intervalo (dias) entre cotas."""
    def stats(g):
        d = g["data"].sort_values()
        gap = d.diff().dt.days.median()
        return pd.Series({"n_obs": len(d), "gap_mediano_dias": gap})
    return serie.groupby("chave").apply(stats, include_groups=False).reset_index()


# --------------------------------------------------------------------------
# SAÍDA em Excel com fórmulas vivas
# --------------------------------------------------------------------------
def exportar_excel(obs: pd.DataFrame, selecao: pd.DataFrame, saida: Path) -> None:
    from openpyxl import Workbook

    obs = obs.sort_values(["chave", "data"]).reset_index(drop=True)
    n = len(obs); fim = n + 1
    PCT = "0.0000%"; DT = "yyyy-mm-dd"

    wb = Workbook()
    # ---- Aba 1: retornos_por_fundo ----
    ws = wb.active
    ws.title = "retornos_por_fundo"
    ws.append(["data", "chave", "cota", "pl", "data_ant", "cota_ant", "pl_ant", "gap_dias",
               "retorno", "valido", "pl_ant_valido", "soma_pl_dia", "peso", "contribuicao"])
    for i, r in enumerate(obs.itertuples(index=False), start=2):
        ws.cell(i, 1, r.data.to_pydatetime()).number_format = DT
        ws.cell(i, 2, r.chave)
        ws.cell(i, 3, float(r.cota))
        ws.cell(i, 4, float(r.pl))
        # valores do dia anterior DO MESMO fundo (linha de cima, se mesma chave)
        ws.cell(i, 5, f'=IF($B{i}=$B{i-1},$A{i-1},"")').number_format = DT
        ws.cell(i, 6, f'=IF($B{i}=$B{i-1},$C{i-1},"")')
        ws.cell(i, 7, f'=IF($B{i}=$B{i-1},$D{i-1},"")')
        # gap em dias entre a cota e a cota anterior
        ws.cell(i, 8, f'=IF($E{i}="","",$A{i}-$E{i})')
        # retorno = cota / cota_ant - 1
        ws.cell(i, 9, f'=IFERROR($C{i}/$F{i}-1,"")').number_format = PCT
        # valido: tem retorno, gap pequeno, cota_ant e pl_ant positivos
        ws.cell(i, 10, f'=IF($I{i}="",FALSE,AND($H{i}<={MAX_GAP_RETORNO_DIAS},$F{i}>0,$G{i}>0))')
        # pl_ant_valido entra no denominador do peso
        ws.cell(i, 11, f'=IF($J{i},$G{i},0)')
        # soma do PL anterior válido no mesmo dia (todos os fundos já são do grupo alvo)
        ws.cell(i, 12, f'=SUMIFS($K$2:$K${fim},$A$2:$A${fim},$A{i})')
        # peso = pl_ant_valido / soma_dia
        ws.cell(i, 13, f'=IF($J{i},$K{i}/$L{i},"")').number_format = PCT
        # contribuicao = peso * retorno
        ws.cell(i, 14, f'=IF($J{i},$M{i}*$I{i},"")').number_format = PCT

    # ---- Aba 2: performance_diaria (data | retorno_pct | n_fundos | indice) ----
    datas = sorted(obs["data"].unique())
    wp = wb.create_sheet("performance_diaria")
    wp.append(["data", "retorno_pct", "n_fundos", "indice_base100"])
    SH = "retornos_por_fundo"
    for j, d in enumerate(datas, start=2):
        wp.cell(j, 1, pd.Timestamp(d).to_pydatetime()).number_format = DT
        # performance do dia = soma das contribuições (Passo 3)
        wp.cell(j, 2, f'=SUMIFS({SH}!$N$2:$N${fim},{SH}!$A$2:$A${fim},$A{j})').number_format = PCT
        # nº de fundos válidos no dia
        wp.cell(j, 3, f'=COUNTIFS({SH}!$K$2:$K${fim},">0",{SH}!$A$2:$A${fim},$A{j})')
        # índice base 100 (cortesia)
        wp.cell(j, 4, "=100" if j == 2 else f"=D{j-1}*(1+B{j})").number_format = "0.0000"
    ultp = len(datas) + 1

    # ---- Aba 3: selecao_fundos (regra 'entra' como fórmula) ----
    wsf = wb.create_sheet("selecao_fundos")
    wsf.append(["chave", "grupo", "n_obs", "gap_mediano_dias", "entra"])
    for i, r in enumerate(selecao.itertuples(index=False), start=2):
        wsf.cell(i, 1, r.chave)
        wsf.cell(i, 2, r.grupo)
        wsf.cell(i, 3, int(r.n_obs))
        wsf.cell(i, 4, None if pd.isna(r.gap_mediano_dias) else float(r.gap_mediano_dias))
        # regra visível: entra se é do grupo alvo, diário e com obs suficientes
        wsf.cell(i, 5, f'=IF(AND($B{i}="{GRUPO_ALVO}",$D{i}<={MAX_GAP_MEDIANO_DIAS},'
                       f'$C{i}>={MIN_OBS}),"sim","nao")')

    # ---- Aba 4: metodologia ----
    wm = wb.create_sheet("metodologia")
    for linha in [
        ["Todas as colunas de calculo sao FORMULAS -- abra no Excel para recalcular."],
        [f"Grupo analisado: {GRUPO_ALVO}. Filtros: diario (gap mediano <= {MAX_GAP_MEDIANO_DIAS} dias) "
         f"e no minimo {MIN_OBS} cotas."],
        ["retorno do fundo:  retorno = cota / cota_ant - 1  (cota_ant = cota da observacao anterior do fundo)"],
        [f"so vale como diario se gap_dias <= {MAX_GAP_RETORNO_DIAS} (cobre fim de semana/feriado); senao sai do dia."],
        ["peso:  peso = pl_ant_valido / soma_pl_dia   (PL da observacao anterior, renormalizado por dia)"],
        ["performance do dia:  retorno_pct = soma das contribuicoes (peso * retorno) dos fundos presentes"],
        ["indice base 100:  indice_t = indice_(t-1) * (1 + retorno_pct_t)"],
        ["ATENCAO: assume cota = total return. Fundos de infra que distribuem rendimento tem a cota"],
        ["reduzida na distribuicao, o que subestima o retorno -- validar antes de concluir."],
    ]:
        wm.append(linha)
    wb.save(saida)
    return ultp


# --------------------------------------------------------------------------
def main() -> None:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    serie = ler_serie(achar_serie())
    cls = ler_classificacao()

    serie = serie.merge(cls, on="chave", how="inner")
    serie = serie[serie["grupo"] == GRUPO_ALVO]
    if serie.empty:
        sys.exit(f"Nenhum fundo do grupo '{GRUPO_ALVO}' casou com a série. Confira CHAVE_SERIE/classificação.")
    log(f"Fundos do grupo '{GRUPO_ALVO}' com série: {serie['chave'].nunique()}")

    # seleção de report diário
    freq = diagnostico_frequencia(serie)
    freq["grupo"] = GRUPO_ALVO
    diarios = freq[(freq["gap_mediano_dias"] <= MAX_GAP_MEDIANO_DIAS) & (freq["n_obs"] >= MIN_OBS)]
    chaves = set(diarios["chave"])
    log(f"  diários (gap mediano <= {MAX_GAP_MEDIANO_DIAS}, obs >= {MIN_OBS}): "
        f"{len(chaves)} de {len(freq)}")
    excluidos = freq[~freq["chave"].isin(chaves)]
    if len(excluidos):
        log(f"  excluídos (não-diários/curtos): {len(excluidos)} "
            f"(ex. gaps medianos: {sorted(excluidos['gap_mediano_dias'].dropna().unique())[:6]})")

    obs = serie[serie["chave"].isin(chaves)][["data", "chave", "cota", "pl"]]
    if obs.empty:
        sys.exit("Nenhum fundo diário sobrou após os filtros -- afrouxe MAX_GAP_MEDIANO_DIAS/MIN_OBS.")

    saida = PASTA_SAIDA / "retorno_diario_cota.xlsx"
    exportar_excel(obs, freq.sort_values("chave"), saida)
    log(f"Excel salvo (fórmulas vivas): {saida}")
    log("  Abra no Excel para recalcular (visualizador que não calcula mostra vazio).")

    # conferência em Python (NÃO é o que vai no Excel; só para log)
    o = obs.sort_values(["chave", "data"]).copy()
    o["cota_ant"] = o.groupby("chave")["cota"].shift(1)
    o["pl_ant"] = o.groupby("chave")["pl"].shift(1)
    o["gap"] = o.groupby("chave")["data"].diff().dt.days
    o["ret"] = o["cota"] / o["cota_ant"] - 1
    val = o["ret"].notna() & (o["gap"] <= MAX_GAP_RETORNO_DIAS) & (o["pl_ant"] > 0)
    o = o[val]
    soma = o.groupby("data")["pl_ant"].transform("sum")
    o["contrib"] = (o["pl_ant"] / soma) * o["ret"]
    perf = o.groupby("data")["contrib"].sum()
    idx = INDICE_BASE * (1 + perf).cumprod()
    log(f"  conferência: {len(perf)} dias; retorno acumulado do grupo = {idx.iloc[-1]/INDICE_BASE-1:.4%}; "
        f"fundos/dia médio = {o.groupby('data')['chave'].nunique().mean():.1f}")
    (PASTA_SAIDA / "log_retorno_diario.txt").write_text("\n".join(LOG), encoding="utf-8")


if __name__ == "__main__":
    main()
