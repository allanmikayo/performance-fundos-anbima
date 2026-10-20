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

## Análise: índice de retorno (infra x ex-infra)

`indice_retorno_fundos.py` lê o `serie_historica.csv` da extração e monta um
índice de valorização diário, ponderado por PL, separado por grupo.

1. `python indice_retorno_fundos.py`
2. Na primeira vez sem `classificacao_fundos.csv`, ele **gera uma base
   fictícia** (metade infra / metade ex-infra) só para o pipeline rodar.
   **Substitua** esse arquivo pela classificação real e rode de novo.

Formato do `classificacao_fundos.csv` (separador `;`):

```
chave;grupo
C0000016713;infra
C0000012345;ex-infra
```

A coluna `chave` casa com `CHAVE_SERIE` do script (padrão
`codigo_classe_consultado`; use `cnpj_input` se preferir casar por CNPJ).

Metodologia (também na aba `metodologia` do Excel):
- retorno por fundo: `r = cota_t / cota_(t-1) - 1`
- peso pelo PL do dia anterior: `w = PL_(t-1) / Σ PL_(t-1) do grupo`
- retorno do grupo: `R = Σ (w · r)`  →  índice base 100: `I_t = I_(t-1)·(1+R)`

Saída em `analise_output/indice_retorno_fundos.xlsx` (abas `indice_diario`,
`retornos_por_fundo` com a trilha de cada número, `resumo`, `metodologia`).

> **Atenção:** o índice assume que a cota é *total return*. Fundos que
> distribuem rendimentos (caso comum nos incentivados de infra) têm a cota
> reduzida na distribuição, o que subestima o retorno — validar antes de
> concluir a comparação infra x ex-infra.

## Retorno diário da cota (um grupo, só fundos diários)

`retorno_diario_cota.py` entrega a série `data | retorno_pct` de um grupo
(padrão `infra`), ponderada por PL, usando **só** fundos com report diário.

- Filtra pelo grupo (classificacao_fundos.csv) e por report diário
  (mediana do intervalo entre cotas ≤ `MAX_GAP_MEDIANO_DIAS`, mín.
  `MIN_OBS` cotas).
- Retorno entre observações consecutivas; só vale como diário se o gap até a
  cota anterior ≤ `MAX_GAP_RETORNO_DIAS` (cobre fim de semana/feriado).
- Peso pelo PL da observação anterior, renormalizado a cada dia entre os
  fundos presentes — resolve comprimentos diferentes e entradas/saídas.

Saída `analise_output/retorno_diario_cota.xlsx`: abas `performance_diaria`
(data, retorno %, nº de fundos, índice base 100), `retornos_por_fundo`
(trilha por fundo × dia), `selecao_fundos` (regra de corte como fórmula) e
`metodologia`. Tudo em fórmulas vivas — abra no Excel para recalcular.
