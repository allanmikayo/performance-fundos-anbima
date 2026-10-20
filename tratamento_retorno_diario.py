"""Tratamento da série -> retorno DIÁRIO da cota (grupo infra), ponderado por PL.

Lê DOIS arquivos:
  1) serie_historica  -- colunas usadas: cnpj_input, data_competencia,
                         valor_patrimonio_liquido, valor_cota
  2) lista_fundos     -- colunas usadas: cnpj, PL, infra
                         (infra = diz se o fundo é de infra ou não)

A ligação entre os dois é pelo CNPJ (cnpj_input <-> cnpj), limpo para 14
dígitos nos dois lados (tira pontuação; não importa o formato).

Saída ENXUTA: analise_output/retorno_diario_cota.xlsx, aba performance_diaria
com  data | retorno_pct | n_fundos | indice_base100. Cálculo todo em Python.

Metodologia (ver conversa):
  retorno do fundo = valor_cota / valor_cota_anterior - 1     (por observação)
  só vale como diário se o gap até a cota anterior <= MAX_GAP_RETORNO_DIAS
  peso = PL / soma(PL dos válidos no dia)                     (renormalizado por dia)
  retorno_pct (grupo) = soma(peso * retorno)
  indice_t = indice_(t-1) * (1 + retorno_pct_t), base 100
  Só entram fundos marcados como infra E com report diário.
"""
from __future__ import annotations

# ╔══════════════════════════════════════════════════════════════════╗
# ║                         CONFIGURAÇÃO                              ║
# ╚══════════════════════════════════════════════════════════════════╝
from pathlib import Path

PASTA_SCRIPT = Path(__file__).resolve().parent

# >>> ALTERAR (1) -- arquivos de entrada (aceita .csv ou .xlsx)
ARQUIVO_SERIE = PASTA_SCRIPT / "serie_historica.csv"
ARQUIVO_LISTA = PASTA_SCRIPT / "lista_fundos.csv"

# >>> ALTERAR (2) -- nomes das colunas
# na serie_historica:
COL_CNPJ_SERIE = "cnpj_input"
COL_DATA = "data_competencia"
COL_PL_SERIE = "valor_patrimonio_liquido"
COL_COTA = "valor_cota"
# no lista_fundos:
COL_CNPJ_LISTA = "cnpj"
COL_PL_LISTA = "PL"
COL_INFRA = "infra"

# Como interpretar a coluna 'infra': valores que significam "é infra"
VALORES_INFRA = {"infra", "sim", "s", "true", "verdadeiro", "1", "yes", "y", "x"}

# >>> ALTERAR (3) -- ponderação:
#   "serie" = pondera pelo PL do dia anterior, da serie_historica (dinâmico; recomendado)
#   "lista" = pondera pelo PL fixo informado no lista_fundos (mesmo peso todo dia)
PESO_POR = "serie"

# Filtros de report diário / buracos
MAX_GAP_MEDIANO_DIAS = 4     # fundo é "diário" se a mediana do intervalo entre cotas <= isto
MIN_OBS = 30                 # ignora séries curtas (menos de N cotas)
MAX_GAP_RETORNO_DIAS = 5     # retorno só vale se a cota anterior está a <= N dias (fds/feriado)

EXPORTAR_TRILHA = True       # grava também um CSV de auditoria com as colunas auxiliares
INDICE_BASE = 100.0
PASTA_SAIDA = PASTA_SCRIPT / "analise_output"
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


def limpa_cnpj(x) -> str | None:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    d = re.sub(r"\D", "", s)
    return d.zfill(14) if d else None


def num(s: pd.Series) -> pd.Series:
    """Converte texto numérico BR ('1.234.567,89') ou normal para float."""
    t = s.astype(str).str.strip()
    # se tem vírgula, assume formato BR (ponto = milhar, vírgula = decimal)
    tem_virgula = t.str.contains(",", regex=False)
    t = t.where(~tem_virgula, t.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(t, errors="coerce")


def ler_tabela(p: Path, obrig: list[str]) -> pd.DataFrame:
    if not p.exists():
        sys.exit(f"Arquivo não encontrado: {p}")
    if p.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(p, dtype=str)
    else:
        cab = p.read_text(encoding="utf-8-sig", errors="ignore").splitlines()[0]
        sep = ";" if ";" in cab else ("\t" if "\t" in cab else ",")
        df = pd.read_csv(p, dtype=str, sep=sep, encoding="utf-8-sig", keep_default_na=False)
    faltam = [c for c in obrig if c not in df.columns]
    if faltam:
        sys.exit(f"Em {p.name} faltam colunas {faltam}. Achei: {list(df.columns)}")
    return df


def achar_serie() -> Path:
    if Path(ARQUIVO_SERIE).exists():
        return Path(ARQUIVO_SERIE)
    cand = sorted(PASTA_SCRIPT.glob("output/*/serie_historica.csv"))
    if cand:
        return cand[-1]
    sys.exit(f"serie_historica não encontrada em {ARQUIVO_SERIE} nem em output/*/serie_historica.csv")


def ler_serie() -> pd.DataFrame:
    p = achar_serie()
    log(f"Lendo série: {p}")
    df = ler_tabela(p, [COL_CNPJ_SERIE, COL_DATA, COL_PL_SERIE, COL_COTA])
    out = pd.DataFrame({
        "cnpj": df[COL_CNPJ_SERIE].map(limpa_cnpj),
        "data": pd.to_datetime(df[COL_DATA], errors="coerce"),
        "pl": num(df[COL_PL_SERIE]),
        "cota": num(df[COL_COTA]),
    }).dropna(subset=["cnpj", "data", "cota"])
    out = out[(out["cota"] > 0) & (out["cnpj"].str.len() == 14)]
    out = out.drop_duplicates(["cnpj", "data"]).sort_values(["cnpj", "data"])
    log(f"  linhas válidas: {len(out)}; fundos na série: {out['cnpj'].nunique()}; "
        f"datas: {out['data'].min():%Y-%m-%d} a {out['data'].max():%Y-%m-%d}")
    return out


def ler_lista() -> pd.DataFrame:
    log(f"Lendo lista de fundos: {ARQUIVO_LISTA}")
    df = ler_tabela(ARQUIVO_LISTA, [COL_CNPJ_LISTA, COL_INFRA])
    lista = pd.DataFrame({
        "cnpj": df[COL_CNPJ_LISTA].map(limpa_cnpj),
        "infra": df[COL_INFRA].astype(str).str.strip().str.lower().isin(VALORES_INFRA),
    })
    lista["pl_lista"] = num(df[COL_PL_LISTA]) if COL_PL_LISTA in df.columns else pd.NA
    lista = lista.dropna(subset=["cnpj"]).drop_duplicates("cnpj")
    log(f"  fundos na lista: {len(lista)}; marcados como infra: {int(lista['infra'].sum())}")
    return lista


def selecionar_diarios(serie: pd.DataFrame) -> set:
    def stats(g):
        d = g["data"].sort_values()
        return pd.Series({"n_obs": len(d), "gap_mediano": d.diff().dt.days.median()})
    freq = serie.groupby("cnpj").apply(stats, include_groups=False).reset_index()
    ok = freq[(freq["gap_mediano"] <= MAX_GAP_MEDIANO_DIAS) & (freq["n_obs"] >= MIN_OBS)]
    log(f"  diários (gap mediano <= {MAX_GAP_MEDIANO_DIAS}, obs >= {MIN_OBS}): {len(ok)} de {len(freq)}")
    return set(ok["cnpj"])


def calcular(obs: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    o = obs.sort_values(["cnpj", "data"]).copy()
    o["cota_ant"] = o.groupby("cnpj")["cota"].shift(1)
    o["pl_ant"] = o.groupby("cnpj")["pl"].shift(1)
    o["gap_dias"] = o.groupby("cnpj")["data"].diff().dt.days
    o["retorno"] = o["cota"] / o["cota_ant"] - 1
    o["grupo"] = np.where(o["infra"], "infra", "nao_infra")

    # fonte do peso
    o["peso_base"] = o["pl_lista"] if PESO_POR == "lista" else o["pl_ant"]

    o["valido"] = (o["retorno"].notna() & (o["gap_dias"] <= MAX_GAP_RETORNO_DIAS)
                   & (o["cota_ant"] > 0) & (o["peso_base"] > 0))

    v = o[o["valido"]].copy()
    # peso é ponderado DENTRO de cada grupo (infra entre infra; não-infra entre não-infra)
    v["soma_peso_dia"] = v.groupby(["data", "grupo"])["peso_base"].transform("sum")
    v["peso"] = v["peso_base"] / v["soma_peso_dia"]
    v["contribuicao"] = v["peso"] * v["retorno"]

    g = (v.groupby(["data", "grupo"])
           .agg(retorno_pct=("contribuicao", "sum"), n_fundos=("cnpj", "nunique"))
           .reset_index())
    ret = g.pivot(index="data", columns="grupo", values="retorno_pct")
    nf = g.pivot(index="data", columns="grupo", values="n_fundos")
    for col in ("infra", "nao_infra"):           # garante as duas colunas mesmo se faltar um grupo
        if col not in ret.columns:
            ret[col] = np.nan
        if col not in nf.columns:
            nf[col] = np.nan
    idx = INDICE_BASE * (1 + ret[["infra", "nao_infra"]].fillna(0.0)).cumprod()

    perf = pd.DataFrame({
        "data": ret.index,
        "retorno_infra": ret["infra"].values,
        "retorno_nao_infra": ret["nao_infra"].values,
        "indice_infra": idx["infra"].values,
        "indice_nao_infra": idx["nao_infra"].values,
        "n_fundos_infra": nf["infra"].values,
        "n_fundos_nao_infra": nf["nao_infra"].values,
    })
    o = o.merge(v[["data", "cnpj", "soma_peso_dia", "peso", "contribuicao"]],
                on=["data", "cnpj"], how="left")
    return o, perf


def exportar(perf: pd.DataFrame, saida: Path) -> None:
    with pd.ExcelWriter(saida, engine="openpyxl") as w:
        perf.to_excel(w, sheet_name="performance_diaria", index=False)
        ws = w.sheets["performance_diaria"]
        fmt = {1: "yyyy-mm-dd", 2: "0.0000%", 3: "0.0000%", 4: "0.0000", 5: "0.0000"}
        for linha in range(2, len(perf) + 2):
            for col, f in fmt.items():
                ws.cell(linha, col).number_format = f
        met = pd.DataFrame({"metodologia": [
            f"Dois grupos lado a lado: infra e nao_infra. So fundos diarios "
            f"(gap mediano <= {MAX_GAP_MEDIANO_DIAS}, obs >= {MIN_OBS}).",
            f"Ponderacao por PL: fonte = '{PESO_POR}' "
            f"({'PL do dia anterior da serie' if PESO_POR=='serie' else 'PL fixo do lista_fundos'}); "
            "peso renormalizado dentro de cada grupo, por dia.",
            "retorno = valor_cota / valor_cota_anterior - 1 (por observacao do fundo).",
            f"so vale como diario se o gap ate a cota anterior <= {MAX_GAP_RETORNO_DIAS} dias.",
            "retorno_<grupo> = soma(peso * retorno) dos fundos do grupo no dia.",
            "indice_<grupo>_t = indice_(t-1) * (1 + retorno_<grupo>_t), base 100.",
            "ATENCAO: assume cota = total return; fundos de infra que distribuem subestimam o retorno.",
            "Auditoria (cota_ant, peso, contribuicao por fundo/dia): retorno_diario_trilha.csv.",
        ]})
        met.to_excel(w, sheet_name="metodologia", index=False)


def main() -> None:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    serie = ler_serie()
    lista = ler_lista()

    # junta (mantém os dois grupos: infra e não-infra)
    serie = serie.merge(lista, on="cnpj", how="inner")
    log(f"CNPJs casados entre série e lista: {serie['cnpj'].nunique()} "
        f"(infra: {serie.loc[serie['infra'],'cnpj'].nunique()}, "
        f"nao_infra: {serie.loc[~serie['infra'],'cnpj'].nunique()})")
    if serie.empty:
        sys.exit("Nenhum CNPJ casou entre série e lista. Confira os CNPJs nos dois arquivos.")
    if PESO_POR == "lista" and serie["pl_lista"].isna().all():
        sys.exit("PESO_POR='lista' mas a coluna de PL do lista_fundos veio vazia/sem valores.")

    chaves = selecionar_diarios(serie)
    obs = serie[serie["cnpj"].isin(chaves)][["data", "cnpj", "cota", "pl", "pl_lista", "infra"]]
    if obs.empty:
        sys.exit("Nenhum fundo diário após os filtros -- afrouxe MAX_GAP_MEDIANO_DIAS/MIN_OBS.")

    trilha, perf = calcular(obs)

    saida = PASTA_SAIDA / "retorno_diario_cota.xlsx"
    exportar(perf, saida)
    log(f"Excel salvo: {saida}  ({len(perf)} dias)")
    for gr in ("infra", "nao_infra"):
        s = perf[f"indice_{gr}"].dropna()
        if len(s):
            log(f"  {gr}: retorno acumulado = {s.iloc[-1]/INDICE_BASE-1:.4%}; "
                f"fundos/dia médio = {perf[f'n_fundos_{gr}'].mean():.1f}")

    if EXPORTAR_TRILHA:
        csv = PASTA_SAIDA / "retorno_diario_trilha.csv"
        trilha.to_csv(csv, index=False, sep=";", decimal=",", encoding="utf-8-sig")
        log(f"  trilha completa: {csv}")

    (PASTA_SAIDA / "log_tratamento.txt").write_text("\n".join(LOG), encoding="utf-8")


if __name__ == "__main__":
    main()
