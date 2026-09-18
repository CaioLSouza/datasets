# Implied target price do Ibovespa (bottom-up)

Consolida os target prices da cobertura XP com o consenso Bloomberg para os papéis
do Ibovespa sem TP nosso e calcula o implied target price do índice.

## Uso

1. `python bottom_up_raw.py` — gera `output\bottom_up_raw.xlsx` (o arquivo precisa estar fechado).
2. Abrir `Implied Ibovespa.xlsx` → **Dados › Atualizar Tudo**.
3. **Bloomberg › Refresh Workbook**, se os BDP não atualizarem sozinhos.

A planilha deste repositório vem **sem dados** (uma linha de placeholder); o primeiro
refresh na rede carrega a composição e a cobertura.

## Arquivos

| Arquivo | O que faz |
|---|---|
| `bottom_up_raw.py` | Lê `COMP SHEET\raw_data.xlsx` e a última composição do Ibov (`economatica-index_composition.parquet`), casa cada papel do índice com um TP da cobertura (direto ou pela outra classe da mesma empresa) e grava as tabelas `ibov`, `coverage` e `meta`. Não aplica filtros. |
| `Implied Ibovespa.xlsx` | Power Query lê o raw; na mesma tabela, colunas `BDP` com `PX_LAST` e `BEST_TARGET_PRICE` e a consolidação. |
| `build_implied_ibov.py` | Remonta a planilha do zero via Excel/COM. Só é necessário para mudar layout ou caminho. |

Caminhos (`\xpdocs\...`) ficam no topo de cada script.

## Regras

- **TP usado** = TP XP válido → consenso Bloomberg → sem TP.
- TP XP é inválido quando a recomendação é Under Review, vazia ou restricted.
- **Filtro de idade**: checkbox *Exclude XP models older than (months)* + célula com X. Ligado, modelos com mais de X meses vão para o consenso.
- Classe não coberta (ex.: BBDC3 com cobertura só da BBDC4): aplica o upside da classe coberta ao preço da outra.
- **Implied = Ibovespa × (1 + Σ peso·upside / Σ peso com TP)**. Papel sem TP nenhum sai da média (pesos renormalizados); o % do índice nessa situação aparece no resumo.

## Observações

- Pesos: composição Economatica na última data disponível; preços e consenso: Bloomberg no momento do refresh.
- `build_implied_ibov.py` grava os formatos de número via marcador no `styles.xml` porque, em Excel com separadores regionais mistos, `Range.NumberFormat` via COM é interpretado com o separador trocado.
