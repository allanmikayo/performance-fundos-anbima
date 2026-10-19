# Performance dos Fundos — extração ANBIMA

Extrai dados brutos de fundos (cota, PL, captação, resgates, nº de cotistas)
da API ANBIMA Feed Fundos v2 (RCVM 175) a partir de uma lista de CNPJs.
Base para construir um índice de retorno de fundos ponderado por PL.

## Como rodar
1. `pip install -r requirements.txt`
2. Copie `modelo_env.txt` para um arquivo `.env` (mesma pasta) e preencha
   `ANBIMA_CLIENT_ID` / `ANBIMA_CLIENT_SECRET`.
3. Preencha `cnpjs.csv` (coluna `cnpj`, com ou sem pontuação).
4. `python extrai_fundos_anbima.py`

Tudo o que muda de um computador para outro está no bloco **CONFIGURAÇÃO**
no topo do script, marcado com `>>> ALTERAR` (credenciais, lista de CNPJs,
pasta de saída e proxy/certificado da rede corporativa).

Comece com `LIMITE_TESTE = 10` para conferir o output; depois troque para
`None` para rodar a lista inteira.

Saídas em `output/<data_hora>/`: `mapa_cnpj_classe.csv`,
`cadastro_classes.csv`, `serie_historica.csv` e as respostas cruas em `raw/`.

> O `.env` com as credenciais fica fora do repositório (`.gitignore`).
