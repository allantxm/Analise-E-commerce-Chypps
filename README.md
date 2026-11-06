# Chypps Tecnologia | Dashboard de vendas (Streamlit)

Painel de inteligência comercial feito sobre a planilha **fVendas.xlsx** (Mercado Livre, Shopee e
WooCommerce consolidados). A planilha é **somente lida**: nenhuma função grava ou altera o arquivo.

## Como rodar
```
pip install -r requirements.txt
streamlit run dashboard_chypps.py
```
Coloque o `fVendas.xlsx` na mesma pasta do `dashboard_chypps.py` (ou envie outra exportação pela
barra lateral).

## Testes
```
python -m unittest discover -s tests -v
```
Com o `fVendas.xlsx` ao lado (ou `BASE_XLSX=/caminho/fVendas.xlsx`), os testes também conferem os
totais da base real e que o arquivo não mudou (SHA-256).

## Medidas do Power BI no app
| Medida DAX | No app |
|---|---|
| Faturamento Bruto = `SUM(Receita por produtos (BRL))` | cartão "Faturamento bruto" (`dados.faturamento_bruto`) |
| Faturamento Líquido = `SUM(Total (BRL))` | cartão "Faturamento líquido" (`dados.faturamento_liquido`) |
| Mês Atual / Mês Anterior (ALLEXCEPT Conta, Estado) | `dados.metricas_mom` |
| % MoM `DIVIDE(atual - anterior, anterior, 0)` com 📈 / 📉 / ⚪ | chip nos cartões (`dados.rotulo_mom`) |

Como no `ALLEXCEPT`, o MoM **só respeita canal e situação**; período, estado e produto são ignorados.
O mês atual é o mês da última venda; o anterior é o mês imediatamente antes.
Valores da base atual (set–out/2026): bruto R$ 626.056,24, líquido R$ 518.846,98;
outubro líquido R$ 310.628,03 vs setembro R$ 208.218,95 = 📈 +49,18%.

## O que a leitura da planilha trata
- Texto `null` vira vazio; nada vira a palavra "nan" nas tabelas.
- Datas reais do Excel ou texto `dd/mm/aaaa hh:mm:ss` (dia primeiro; `aaaa-mm-dd` é lido como ISO).
- UF por extenso ou sigla (`São Paulo`, `sp`, `SP`) vira sigla.
- **Número do pedido do Mercado Livre**: o Excel arredonda o `N.º de venda` (quase todas as linhas
  ficam com 2000020000000000). Quando um número se repete em mais de 30% das linhas de um canal,
  o pedido é reconstruído por canal + data/hora + comprador. Assim "Pedidos" e "Ticket médio"
  deixam de colapsar. O ideal é exportar o número como texto na origem.
- **Situação**: ~30 textos de status são agrupados em Concluído, Em andamento, Cancelado e
  Devolvido/reembolsado. Por padrão tudo entra (como no Power BI).
- Frete e desconto são do pedido: contam uma vez por pedido, não por linha.
- O indicador visual de margem bruta estimada foi removido do dashboard.

## Pontos de atenção na base (não alterados)
- Pedidos com várias linhas repetem o mesmo `Total (BRL)` em cada linha; o `SUM` (igual ao DAX)
  conta esse total mais de uma vez. O painel avisa na aba "Dados e exportação".
- A planilha contém dados pessoais (comprador, CPF, endereço, telefone). O painel mostra só iniciais
  e não leva essas colunas ao CSV exportado. Evite versionar o `fVendas.xlsx`.
