"""Extração BRUTA de dados de fundos -- API ANBIMA Feed Fundos v2 (RCVM 175)
a partir de uma LISTA DE CNPJs.

Objetivo: base para um índice de retorno de fundos ponderado (por PL).
Este script só EXTRAI e salva; nenhum cálculo de retorno é feito aqui.

Fluxo:
  0. Lê a lista de CNPJs (arquivo de entrada).
  1. Universo  -- GET /feed/fundos/v2/fundos (paginado, ~30-40 chamadas)
                  Serve para traduzir CNPJ -> codigo_classe (a API só aceita
                  o código ANBIMA nas consultas seguintes). Aceita tanto CNPJ
                  do FUNDO quanto CNPJ da CLASSE. Fica em cache por
                  DIAS_CACHE_UNIVERSO dias para não baixar toda vez.
  2. Cadastro  -- GET /feed/fundos/v2/fundos/{codigo_classe}
  3. Série     -- GET /feed/fundos/v2/fundos/{codigo_classe}/serie-historica
                  (cota, PL, captação, resgates, nº de cotistas)

Como rodar (terminal do VS Code, na pasta do script):
    pip install -r requirements.txt     (só na primeira vez)
    python extrai_fundos_anbima.py

=====================================================================
  PARA RODAR EM OUTRO COMPUTADOR: só mexa no bloco "CONFIGURAÇÃO"
  logo abaixo. Cada item marcado com  >>> ALTERAR  é o que costuma
  mudar de uma máquina para outra.
=====================================================================
"""
from __future__ import annotations

# ╔══════════════════════════════════════════════════════════════════╗
# ║                         CONFIGURAÇÃO                             ║
# ╚══════════════════════════════════════════════════════════════════╝
from pathlib import Path

PASTA_SCRIPT = Path(__file__).resolve().parent

# >>> ALTERAR (1) -- CREDENCIAIS ANBIMA
# Crie um arquivo chamado ".env" na MESMA pasta deste script com duas linhas
# (modelo em "modelo_env.txt"):
#     ANBIMA_CLIENT_ID=seu_client_id
#     ANBIMA_CLIENT_SECRET=seu_client_secret
# Se o .env estiver em outro lugar, coloque o caminho completo aqui:
ARQUIVO_ENV = PASTA_SCRIPT / ".env"

# >>> ALTERAR (2) -- LISTA DE CNPJs (entrada)
# Aceita .csv, .xlsx ou .txt. CNPJ com ou sem pontuação.
#   - .csv / .xlsx: precisa ter uma coluna com o nome em COLUNA_CNPJ
#   - .txt: um CNPJ por linha
ARQUIVO_CNPJS = PASTA_SCRIPT / "cnpjs.csv"
COLUNA_CNPJ = "cnpj"

# >>> ALTERAR (3) -- PASTA DE SAÍDA
PASTA_SAIDA = PASTA_SCRIPT / "output"

# >>> ALTERAR (4) -- REDE CORPORATIVA (proxy / certificado)
# Em casa normalmente fica tudo None. No notebook corporativo, se der erro de
# conexão / SSL (ex.: "SSLError", "certificate verify failed", "ProxyError"):
#   - PROXY: pegue com a TI ou em "Configurações de proxy" do Windows.
#     Ex.: "http://proxy.empresa.com.br:8080"
#     Com usuário/senha: "http://USUARIO:SENHA@proxy.empresa.com.br:8080"
#   - CERTIFICADO_CA: caminho do .pem/.crt da cadeia corporativa
#     (TI fornece). Ex.: r"C:\certs\corporativo.pem"
PROXY = None
CERTIFICADO_CA = None

# PERÍODO da série histórica ("AAAA-MM-DD" ou None).
# None nos dois = API devolve os últimos 5 anos. Períodos > 5 anos são
# quebrados em janelas automaticamente.
DATA_INICIO = "2025-10-01"
DATA_FIM = "2026-10-16"

# >>> ALTERAR (5) -- SAÍDA EM EXCEL (.xlsx) para subir no GitHub
# Além dos CSVs (que ficam em output/<data_hora>/, ignorados pelo git), o
# script grava os 3 resultados principais em Excel numa pasta fixa, que é
# o que você sobe no GitHub. Sobrescreve a cada rodada.
GERAR_EXCEL = True
PASTA_EXCEL = PASTA_SCRIPT / "resultados_excel"

# TESTE: processa só os N primeiros CNPJs da lista (None = todos).
# Recomendado rodar com 10 na primeira vez para conferir o output;
# troque para None quando quiser a lista inteira.
LIMITE_TESTE = 10

# Universo (lista de fundos da ANBIMA) fica em cache por N dias.
DIAS_CACHE_UNIVERSO = 7

# Pausa entre chamadas (s). Aumente se aparecer muito HTTP 429.
PAUSA_ENTRE_CHAMADAS = 0.25
# ════════════════════ FIM DA CONFIGURAÇÃO ═══════════════════════════


import base64
import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta

import pandas as pd
import requests
from dotenv import load_dotenv

BASE_URL = "https://api.anbima.com.br"
MAX_TENTATIVAS = 4
JANELA_MAX_DIAS = 365 * 5 - 5  # limite da API: 5 anos por consulta
CSV_KW = dict(index=False, sep=";", decimal=",", encoding="utf-8-sig")

LOG: list[str] = []


def log(msg: str) -> None:
    linha = f"[{datetime.now():%H:%M:%S}] {msg}"
    print(linha, flush=True)
    LOG.append(linha)


# --------------------------------------------------------------------------
# Entrada
# --------------------------------------------------------------------------
def limpa_cnpj(x) -> str | None:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if s.endswith(".0"):          # Excel às vezes lê CNPJ como número
        s = s[:-2]
    d = re.sub(r"\D", "", s)
    if not d:
        return None
    return d.zfill(14)


def ler_cnpjs() -> list[str]:
    p = Path(ARQUIVO_CNPJS)
    if not p.exists():
        sys.exit(f"Arquivo de CNPJs não encontrado: {p}\n-> ajuste ARQUIVO_CNPJS na CONFIGURAÇÃO.")
    if p.suffix.lower() == ".txt":
        brutos = p.read_text(encoding="utf-8-sig").splitlines()
    else:
        if p.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(p, dtype=str)
        else:
            cab = p.read_text(encoding="utf-8-sig").splitlines()[0] if p.stat().st_size else ""
            sep = ";" if ";" in cab else ("," if "," in cab else ("\t" if "\t" in cab else ";"))
            df = pd.read_csv(p, dtype=str, sep=sep, encoding="utf-8-sig")
        cols = {c.strip().lower(): c for c in df.columns}
        if COLUNA_CNPJ.lower() not in cols:
            sys.exit(f"Coluna '{COLUNA_CNPJ}' não encontrada em {p.name}. Colunas: {list(df.columns)}\n"
                     f"-> ajuste COLUNA_CNPJ na CONFIGURAÇÃO.")
        brutos = df[cols[COLUNA_CNPJ.lower()]].tolist()
    cnpjs, invalidos = [], []
    for b in brutos:
        c = limpa_cnpj(b)
        if c is None:
            continue
        (cnpjs if len(c) == 14 else invalidos).append(c)
    cnpjs = list(dict.fromkeys(cnpjs))  # remove duplicados mantendo a ordem
    if invalidos:
        log(f"  {len(invalidos)} CNPJ(s) com mais de 14 dígitos ignorados: {invalidos[:5]}")
    log(f"CNPJs lidos de {p.name}: {len(cnpjs)} únicos")
    if LIMITE_TESTE:
        cnpjs = cnpjs[:LIMITE_TESTE]
        log(f"  MODO TESTE: usando só os {len(cnpjs)} primeiros (LIMITE_TESTE={LIMITE_TESTE})")
    return cnpjs


# --------------------------------------------------------------------------
# Autenticação / HTTP
# --------------------------------------------------------------------------
def carregar_credenciais() -> tuple[str, str]:
    p = Path(ARQUIVO_ENV)
    if p.exists():
        load_dotenv(p, override=False)
        log(f".env carregado de: {p}")
    cid = os.getenv("ANBIMA_CLIENT_ID", "").strip()
    sec = os.getenv("ANBIMA_CLIENT_SECRET", "").strip()
    if not cid or not sec:
        sys.exit(f"ANBIMA_CLIENT_ID / ANBIMA_CLIENT_SECRET não encontrados (procurei em {p}).\n"
                 "-> crie o .env a partir do modelo_env.txt ou ajuste ARQUIVO_ENV.")
    return cid, sec


class Anbima:
    def __init__(self, client_id: str, client_secret: str):
        self.cid, self.sec = client_id, client_secret
        self.token: str | None = None
        self.expira = 0.0
        self.n_chamadas = 0
        self.s = requests.Session()
        if PROXY:
            self.s.proxies.update({"http": PROXY, "https": PROXY})
        if CERTIFICADO_CA:
            self.s.verify = str(CERTIFICADO_CA)

    def _novo_token(self) -> None:
        basic = base64.b64encode(f"{self.cid}:{self.sec}".encode()).decode()
        try:
            r = self.s.post(
                f"{BASE_URL}/oauth/access-token",
                headers={"Content-Type": "application/json", "Authorization": f"Basic {basic}"},
                json={"grant_type": "client_credentials"},
                timeout=30,
            )
        except requests.exceptions.SSLError as e:
            sys.exit(f"Erro de SSL ao conectar na ANBIMA ({e}).\n"
                     "-> rede corporativa: preencha CERTIFICADO_CA na CONFIGURAÇÃO (item 4).")
        except (requests.exceptions.ProxyError, requests.exceptions.ConnectionError) as e:
            sys.exit(f"Sem conexão com api.anbima.com.br ({e.__class__.__name__}).\n"
                     "-> rede corporativa: preencha PROXY na CONFIGURAÇÃO (item 4).")
        if r.status_code in (401, 403):
            sys.exit(f"OAuth recusado (HTTP {r.status_code}): {r.text[:300]}\n-> confira client_id/secret no .env.")
        r.raise_for_status()
        d = r.json()
        self.token = d["access_token"]
        self.expira = time.time() + float(d.get("expires_in", 3600))
        log(f"token OAuth obtido (expira em {d.get('expires_in')}s)")

    def get(self, path: str, params: dict | None = None) -> tuple[int, object]:
        """Retorna (status, json). 404 -> (404, None) sem exceção."""
        for tentativa in range(1, MAX_TENTATIVAS + 1):
            if not self.token or time.time() > self.expira - 60:
                self._novo_token()
            time.sleep(PAUSA_ENTRE_CHAMADAS)
            self.n_chamadas += 1
            try:
                r = self.s.get(f"{BASE_URL}{path}",
                               headers={"client_id": self.cid, "access_token": self.token},
                               params=params or {}, timeout=90)
            except requests.RequestException as e:
                log(f"  erro de rede ({e.__class__.__name__}) em {path} -- tentativa {tentativa}")
                time.sleep(2 ** tentativa)
                continue
            if r.status_code == 401:
                self.token = None
                continue
            if r.status_code == 403:
                log(f"  HTTP 403 em {path}: seu contrato ANBIMA pode não incluir o pacote de Fundos. {r.text[:200]}")
                return 403, None
            if r.status_code == 404:
                return 404, None
            if r.status_code in (429, 500, 502, 503, 504):
                espera = 2 ** tentativa * 2
                log(f"  HTTP {r.status_code} em {path} -- aguardando {espera}s")
                time.sleep(espera)
                continue
            if r.status_code >= 400:
                log(f"  HTTP {r.status_code} em {path} params={params}: {r.text[:300]}")
                return r.status_code, None
            try:
                return r.status_code, r.json()
            except ValueError:
                log(f"  resposta não-JSON em {path}: {r.text[:300]}")
                return r.status_code, None
        log(f"  desistindo de {path} após {MAX_TENTATIVAS} tentativas")
        return -1, None


# --------------------------------------------------------------------------
# Helpers (formato da resposta não foi testado ao vivo -> tudo defensivo,
# e as respostas cruas são salvas em raw/ para conferência)
# --------------------------------------------------------------------------
def como_lista(x) -> list:
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


def conteudo(resp) -> list[dict]:
    if isinstance(resp, list):
        return resp
    if isinstance(resp, dict):
        return como_lista(resp.get("content", resp))
    return []


def escalares(d: dict, prefixo: str = "") -> dict:
    return {f"{prefixo}{k}": v for k, v in d.items() if not isinstance(v, (list, dict))}


def salvar_json(obj, caminho: Path) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def salvar_excel(df: pd.DataFrame, nome: str) -> None:
    """Grava uma cópia .xlsx na pasta fixa PASTA_EXCEL (a que vai pro GitHub).
    Excel tem limite de ~1.048.576 linhas; se passar, o .xlsx é pulado e fica
    só o CSV (avisado no log)."""
    if not GERAR_EXCEL:
        return
    Path(PASTA_EXCEL).mkdir(parents=True, exist_ok=True)
    destino = Path(PASTA_EXCEL) / f"{nome}.xlsx"
    if len(df) > 1_048_575:
        log(f"  '{nome}' tem {len(df)} linhas (acima do limite do Excel) -- .xlsx pulado, use o CSV")
        return
    try:
        df.to_excel(destino, index=False)
        log(f"  Excel: {destino}")
    except Exception as e:
        log(f"  falha ao gravar {destino.name} ({e.__class__.__name__}: {e}) -- CSV foi mantido")


# --------------------------------------------------------------------------
# Etapa 1 -- universo (com cache) e mapeamento CNPJ -> codigo_classe
# --------------------------------------------------------------------------
def baixar_universo(api: Anbima, out: Path) -> pd.DataFrame:
    log("ETAPA 1 -- listagem de fundos (/feed/fundos/v2/fundos)")
    linhas: list[dict] = []
    pagina = 0
    while True:
        status, resp = api.get("/feed/fundos/v2/fundos", {"page": pagina, "size": 1000})
        if status == 403:
            sys.exit("Sem acesso ao pacote de Fundos da ANBIMA (HTTP 403). Veja log acima.")
        if status == 404 or resp is None:
            break
        if pagina == 0:
            salvar_json(resp, out / "raw" / "listagem_pagina0.json")
        fundos = conteudo(resp)
        if not fundos:
            break
        for f in fundos:
            base_f = escalares(f)
            classes = como_lista(f.get("classes"))
            if not classes:
                linhas.append(base_f)
                continue
            for c in classes:
                base_c = {**base_f, **escalares(c)}
                subs = como_lista(c.get("subclasses") or c.get("subclases"))
                if not subs:
                    linhas.append(base_c)
                for s in subs:
                    linhas.append({**base_c, **escalares(s, "subclasse_")})
        log(f"  página {pagina}: {len(fundos)} fundos (acumulado {len(linhas)} linhas)")
        if len(fundos) < 1000:
            break
        pagina += 1
    return pd.DataFrame(linhas)


def obter_universo(api: Anbima, out: Path) -> pd.DataFrame:
    cache = Path(PASTA_SAIDA) / "_cache_universo.csv"
    if cache.exists():
        idade = (time.time() - cache.stat().st_mtime) / 86400
        if idade <= DIAS_CACHE_UNIVERSO:
            log(f"ETAPA 1 -- universo do cache ({idade:.1f} dias): {cache}")
            return pd.read_csv(cache, sep=";", dtype=str, encoding="utf-8-sig")
    df = baixar_universo(api, out)
    if df.empty:
        sys.exit("Listagem de fundos veio vazia -- veja raw/listagem_pagina0.json e log.txt.")
    df = df.astype("string")
    df.to_csv(cache, index=False, sep=";", encoding="utf-8-sig")
    log(f"  universo salvo em cache: {len(df)} linhas")
    return df


def mapear_cnpjs(universo: pd.DataFrame, cnpjs: list[str]) -> pd.DataFrame:
    """Para cada CNPJ, acha a(s) classe(s). Procura em identificador_classe
    (CNPJ da classe) e identificador_fundo (CNPJ do fundo)."""
    u = universo.copy()
    for col in ("identificador_classe", "identificador_fundo", "identificador_subclasse",
                "subclasse_identificador_subclasse"):
        if col in u:
            u[f"_{col}"] = u[col].map(limpa_cnpj)
    linhas = []
    for c in cnpjs:
        achou = False
        for col, origem in (("_identificador_classe", "classe"), ("_identificador_fundo", "fundo"),
                            ("_identificador_subclasse", "subclasse"),
                            ("_subclasse_identificador_subclasse", "subclasse")):
            if col not in u:
                continue
            m = u[u[col] == c]
            if m.empty:
                continue
            for cod in m["codigo_classe"].dropna().unique():
                linha = m[m["codigo_classe"] == cod].iloc[0]

                def v(col):
                    x = linha.get(col)
                    return None if x is None or pd.isna(x) else x

                linhas.append({
                    "cnpj_input": c, "encontrado_como": origem, "codigo_classe": cod,
                    "codigo_fundo": v("codigo_fundo"),
                    "nome_fundo": v("nome_comercial_fundo"),
                    "nome_classe": v("nome_comercial_classe"),
                    "data_encerramento_classe": v("data_encerramento"),
                })
            achou = True
            break
        if not achou:
            linhas.append({"cnpj_input": c, "encontrado_como": "NAO ENCONTRADO"})
    return pd.DataFrame(linhas)


# --------------------------------------------------------------------------
# Etapa 2 -- cadastro
# --------------------------------------------------------------------------
def extrair_cadastro(api: Anbima, mapa: pd.DataFrame, out: Path) -> pd.DataFrame:
    alvos = mapa.dropna(subset=["codigo_classe"])
    log(f"ETAPA 2 -- detalhes de {len(alvos)} classes")
    linhas = []
    for i, (_, alvo) in enumerate(alvos.iterrows(), 1):
        cod = alvo["codigo_classe"]
        status, resp = api.get(f"/feed/fundos/v2/fundos/{cod}")
        if resp is None:
            log(f"  [{i}/{len(alvos)}] {cod}: sem detalhes (HTTP {status})")
            continue
        salvar_json(resp, out / "raw" / "detalhes" / f"{cod}.json")
        reg = {"cnpj_input": alvo["cnpj_input"], "codigo_classe_consultado": cod}
        topo = resp if isinstance(resp, dict) else {"content": resp}
        for f in conteudo(topo)[:1]:
            reg.update(escalares(f, "fundo_"))
        classes = como_lista(topo.get("classes"))
        if not classes and conteudo(topo):
            classes = como_lista(conteudo(topo)[0].get("classes"))
        cl = next((c for c in classes if c.get("codigo_classe") == cod), classes[0] if classes else {})
        reg.update(escalares(cl, "classe_"))
        for bloco in ("perfil-complementar", "perfil_complementar", "indices-benchmark", "indices_benchmark"):
            itens = como_lista(topo.get(bloco) or cl.get(bloco))
            if itens and isinstance(itens[0], dict):
                reg.update(escalares(itens[0], bloco.replace("-", "_") + "_"))
        for t in como_lista(topo.get("taxas") or cl.get("taxas")):
            if isinstance(t, dict) and t.get("tipo_taxa"):
                reg.setdefault(f"taxa_{t['tipo_taxa']}_pct", t.get("valor_percentual"))
        linhas.append(reg)
        if i % 10 == 0 or i == len(alvos):
            log(f"  [{i}/{len(alvos)}] ok")
    df = pd.DataFrame(linhas)
    df.to_csv(out / "cadastro_classes.csv", **CSV_KW)
    salvar_excel(df, "cadastro_classes")
    log(f"  cadastro: {len(df)} classes, {df.shape[1]} colunas")
    return df


# --------------------------------------------------------------------------
# Etapa 3 -- série histórica
# --------------------------------------------------------------------------
def janelas(inicio: date, fim: date) -> list[tuple[date, date]]:
    out, a = [], inicio
    while a <= fim:
        b = min(a + timedelta(days=JANELA_MAX_DIAS), fim)
        out.append((a, b))
        a = b + timedelta(days=1)
    return out


def extrair_serie(api: Anbima, mapa: pd.DataFrame, out: Path) -> pd.DataFrame:
    alvos = mapa.dropna(subset=["codigo_classe"])
    inicio = date.fromisoformat(DATA_INICIO) if DATA_INICIO else None
    fim = date.fromisoformat(DATA_FIM) if DATA_FIM else date.today()
    jan = janelas(inicio, fim) if inicio else [(None, None)]
    log(f"ETAPA 3 -- série histórica de {len(alvos)} classes, {len(jan)} janela(s) cada")
    frames = []
    for i, (_, alvo) in enumerate(alvos.iterrows(), 1):
        cod = alvo["codigo_classe"]
        regs = []
        for a, b in jan:
            params = {"data-inicio": a.isoformat(), "data-fim": b.isoformat()} if a else {}
            pagina = 0
            while True:
                p = dict(params)
                if pagina:
                    p["page"] = pagina
                status, resp = api.get(f"/feed/fundos/v2/fundos/{cod}/serie-historica", p)
                if resp is None:
                    break
                if i == 1 and pagina == 0:
                    salvar_json(resp, out / "raw" / f"serie_exemplo_{cod}.json")
                regs.extend(conteudo(resp))
                tot = (resp.get("total_pages") or resp.get("totalPages")) if isinstance(resp, dict) else None
                if tot and pagina + 1 < int(tot):
                    pagina += 1
                    continue
                break
        if regs:
            df = pd.DataFrame([escalares(r) for r in regs])
            df.insert(0, "codigo_classe_consultado", cod)
            df.insert(0, "cnpj_input", alvo["cnpj_input"])
            frames.append(df)
        log(f"  [{i}/{len(alvos)}] {alvo['cnpj_input']} / {cod}: {len(regs)} registros")
    if not frames:
        log("  nenhuma série retornada")
        return pd.DataFrame()
    serie = pd.concat(frames, ignore_index=True)
    if "data_competencia" in serie:
        serie = serie.drop_duplicates(["codigo_classe_consultado", "data_competencia"])
        serie = serie.sort_values(["cnpj_input", "codigo_classe_consultado", "data_competencia"])
    serie.to_csv(out / "serie_historica.csv", **CSV_KW)
    salvar_excel(serie, "serie_historica")
    log(f"  série: {len(serie)} linhas")
    return serie


# --------------------------------------------------------------------------
def main() -> None:
    out = Path(PASTA_SAIDA) / datetime.now().strftime("%Y%m%d_%H%M%S")
    (out / "raw").mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    api = None
    try:
        cnpjs = ler_cnpjs()
        api = Anbima(*carregar_credenciais())
        universo = obter_universo(api, out)
        mapa = mapear_cnpjs(universo, cnpjs)
        mapa.to_csv(out / "mapa_cnpj_classe.csv", **CSV_KW)
        salvar_excel(mapa, "mapa_cnpj_classe")
        nao = mapa[mapa["encontrado_como"] == "NAO ENCONTRADO"]["cnpj_input"].tolist()
        log(f"mapeamento: {len(cnpjs) - len(nao)}/{len(cnpjs)} CNPJs encontrados -> "
            f"{mapa['codigo_classe'].nunique() if 'codigo_classe' in mapa else 0} classes")
        if nao:
            log(f"  NÃO encontrados ({len(nao)}): {nao[:10]}{' ...' if len(nao) > 10 else ''}")
        if "codigo_classe" not in mapa or mapa["codigo_classe"].isna().all():
            sys.exit("Nenhum CNPJ encontrado no universo ANBIMA -- veja mapa_cnpj_classe.csv.")
        extrair_cadastro(api, mapa, out)
        extrair_serie(api, mapa, out)
    finally:
        log(f"fim -- {api.n_chamadas if api else 0} chamadas em {time.time() - t0:,.0f}s. Output: {out}")
        (out / "log.txt").write_text("\n".join(LOG), encoding="utf-8")


if __name__ == "__main__":
    main()
