# Amazon Inventory / Sales Reports (SP-API)

Scripts Python para coletar o inventário e o volume de vendas da sua loja na
Amazon via **SP-API** (Seller Partner API) e gerar planilhas XLSX.

| Script | O que faz |
|--------|-----------|
| `inventario.py` | Baixa inventário/vendas da Amazon e gera o relatório mensal XLSX |
| `evolucao.py`   | Lê todos os `relatorio_mensal_*.xlsx` e gera a evolução de vendas com gráficos |
| `dashboard.py`  | Lê os caches e gera `analytics.html` (dashboard Chart.js, sem servidor) |
| `coleta_mensal.sh` | Roda `--mensal` + `dashboard.py` de forma agendada (systemd timer) |

## Requisitos

- Python 3.10+
- Aplicação registrada no [Developer Console SP-API](https://developer.amazonservices.com/) com permissão de vendedor
  - Papéis necessários mínimos: **Orders** (vendas) e **Reporting** (inventário/listagens)
  - Fluxo de autorização do vendedor para obter o refresh token (OAuth)

## Configuração

1. Clone o repo e copie o modelo de ambiente:

   ```bash
   cp .env.example .env
   ```

2. Preencha o `.env` com as credenciais do seu app SP-API:

   ```dotenv
   LWA_APP_ID=amzn1.application-oa2-client.SUBSTITUA
   LWA_CLIENT_SECRET=amzn1.oa2-cs.v1.SUBSTITUA
   SP_API_REFRESH_TOKEN=Atzr|SUBSTITUA
   SP_API_DEFAULT_MARKETPLACE=BR
   ```

   | Variável | Descrição |
   |----------|-----------|
   | `LWA_APP_ID` | Client ID do app (Developer Console SP-API) |
   | `LWA_CLIENT_SECRET` | Client Secret do app |
   | `SP_API_REFRESH_TOKEN` | Refresh token do vendedor autorizado (fluxo OAuth) |
   | `SP_API_DEFAULT_MARKETPLACE` | Marketplace padrão (default `BR`) |
   | `SP_API_PRIVATE_KEY` | (opcional) Caminho do PEM com chave privada EC |

3. Crie o venv e instale as dependências:

   ```bash
   python3 -m venv .venv
   .venv/bin/pip install -r requirements.txt
   ```

> **Segurança**: o `.env` está no `.gitignore` — nunca versionar as credenciais.
> Relatórios gerados (`.xlsx`) e o cache também ficam fora do git.

## Relatório mensal de vendas (padrão)

Baixa o inventário atual + as vendas do mês e grava `relatorio_mensal_<mês>_<ano>.xlsx`:

```bash
.venv/bin/python inventario.py --mensal
```

- **Janela padrão**: 1º dia do mês atual até hoje
- **Janela personalizada**:

  ```bash
  .venv/bin/python inventario.py --mensal --start 2026-05-01 --end 2026-05-31
  ```

- **Salvar em outro arquivo**:

  ```bash
  .venv/bin/python inventario.py --mensal --out caminho/arquivo.xlsx
  ```

### Abas do relatório mensal

| Aba | Conteúdo |
|-----|----------|
| **Inventario** | Listagens ativas (SKU, ASIN, preço, estoque, status) |
| **Vendas Mensais** | Agregado por SKU (unidades, pedidos, bruto, desconto, líquido) + linha TOTAL |
| **Todas as Vendas** | Detalhe por pedido (id, data, valor, estado, pagamento) + colunas de taxas Amazon |
| **Instruções** | Guia de uso embutido na planilha |

## Taxas Amazon por venda (settlement)

As colunas de taxas na aba **Todas as Vendas** vêm do relatório de settlement
(`GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2`), que a Amazon gera sozinha
(~ a cada 2 semanas). Cada tipo de taxa vira uma coluna (Comissão, FBA por
unidade, Financing, chargeback, etc.) somada por ID do pedido, além de
**Total Taxas Amazon**.

- **Sinal**: negativo = custo pago; positivo = devolução/ajuste (ex.: comissão devolvida).
- **Junção**: por `order-id` — pedidos com 2 SKUs somam as taxas dos itens na linha.
- **Limite**: settlements ficam retidos ~90 dias na SP-API; meses antigos (>~3 meses)
  saem **sem** colunas de taxa. Os dados ficam no cache (`settlements`), então o
  `--update` reconstrói com as taxas salvas.
- **Fora do escopo**: armazenagem, assinatura e publicidade não são taxas por venda.

### Reconstruir sem rede (cache)

Os dados brutos do último `--mensal` ficam em `cache/`. Dá pra regenerar a
planilha sem acessar a Amazon:

```bash
.venv/bin/python inventario.py --update
```

## Histórico de meses antigos

Pedidos além de 30 dias usam a Orders API (pode demorar — pagina pedidos e itens):

```bash
.venv/bin/python inventario.py --historico --start 2024-11-01 --end 2024-11-30
```

O resultado também é cacheado em `cache/cache_historico_<aaaa-mm>.json` e gera
`relatorio_mensal_<mês>_<ano>.xlsx` na mesma pasta.

## Outros relatórios (--report)

| Atalho | Report type | Exige papel |
|--------|-------------|-------------|
| `listagens` | `GET_MERCHANT_LISTINGS_ALL_DATA` | Reporting |
| `pedidos` | `GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL` | Orders |
| `pedidos-update` | `..._BY_LAST_UPDATE_GENERAL` | Orders |
| `arquivado` | `GET_FLAT_FILE_ARCHIVED_ORDERS_DATA_BY_ORDER_DATE` | Orders |
| `myi` | `GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA` | FBA (inventory) |
| `reservado` | `GET_RESERVED_INVENTORY_DATA` | FBA (inventory) |
| `afn` | `GET_AFN_INVENTORY_DATA` | Amazon Fulfillment |
| `saude` | `GET_FBA_INVENTORY_PLANNING_DATA` | Amazon Fulfillment |

Exemplos:

```bash
# Inventário listagens (padrão se nenhum --report)
.venv/bin/python inventario.py --report listagens

# Vendas com janela definida
.venv/bin/python inventario.py --report pedidos --start 2026-05-01 --end 2026-05-31

# Combinar relatórios em uma planilha
.venv/bin/python inventario.py --report listagens pedidos
```

> Relatórios FBA que exigem o papel **Amazon Fulfillment** retornam erro 403
> (`Unauthorized`). O erro inclui a dica de como adicionar o papel no Seller Central.

## Evolução de vendas com gráficos

Lê todos os `relatorio_mensal_*.xlsx` da pasta e gera `evolucao_de_vendas.xlsx`:

```bash
.venv/bin/python evolucao.py
```

A planilha contém:

- **Evolucao de Vendas** — série por mês (pedidos, unidades, total, variação %) + gráficos:
  - Total faturado por mês (colunas)
  - Pedidos (colunas) + unidades (linha) — combo
- **Faturamento Acumulado** — total acumulado mês a mês + gráfico de linha

Para incluir um novo mês: rode `inventario.py --mensal --start ... --end ...`
e depois `evolucao.py` de novo (ele lê o novo arquivo automaticamente).

## Dashboard de analytics

Lê os caches e gera um **dashboard HTML estático** (Chart.js, via CDN — não
precisa de servidor, abre direto no navegador):

```bash
.venv/bin/python dashboard.py --out analytics.html
```

Painéis incluídos:

- **Faturamento/pedidos/unidades/meses cobertos/ticket médio/best mês** (cards)
- **Evolução do faturamento** por mês (barras)
- **Pedidos × unidades** por mês (barras + linha)
- **Top 10 SKUs** por faturamento (barras horizontais)
- **Taxas Amazon por tipo** (soma dos settlements — negativos = custo)
- **Estoque/catálogo corrente** (top SKUs ativos por preço)

> Regra de estoque: o relatório de listagens (`GET_MERCHANT_LISTINGS_ALL_DATA`)
> **não traz a coluna `quantity`** (vem vazia). Quantidade real exige os
> relatórios FBA (`AFN`/`MYI`/`Reservado`), que no seu papel SP-API retornam
> 403 — então o dashboard mostra o **catálogo por preço** quando a quantidade
> não está disponível.

## Observações

- Pedidos **cancelados** são excluídos do resumo de vendas.
- Pedidos com 2 SKUs contam **1 vez** no total de pedidos.
- Valores em **BRL** (marketplace BR).
- O guia embutido também fica na aba **Instruções** dos relatórios mensais.