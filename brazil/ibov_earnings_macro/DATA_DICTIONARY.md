# Dicionário de dados

## Convenções gerais

- `month`: período mensal no formato `YYYY-MM`.
- `date`: data efetiva da observação no arquivo de origem.
- Valores ausentes permanecem vazios.
- Percentuais logarítmicos são calculados como `100 × ln(valor_t / valor_t-k)`.
- Variações de Focus, IPCA e swap estão em pontos percentuais.
- Preços, câmbio e índices de commodities entram como variações logarítmicas percentuais.

## `data/earnings_12m_fwd_sector_monthly.csv`

| Coluna | Definição |
|---|---|
| `source_file` | Grid histórico que contém a observação |
| `IBOV` | BEst Net Income agregado do índice |
| Setores GICS | BEst Net Income agregado reportado para cada setor |
| `Not Classified` | Parcela sem classificação setorial no grid, quando disponível |

Unidade dos níveis: BRL, conforme o metadado dos grids. A escala do agregado é a reportada pela fonte.

## `data/earnings_revisions_3m_sector_monthly.csv`

Cada coluna de earnings contém `100 × ln(Earnings_t / Earnings_t-3)`. A revisão fica vazia quando o nível atual ou defasado é ausente ou não positivo.

## `data/earnings_12m_fwd_sector_base100_2020_01.csv`

Índice calculado como `100 × Earnings_t / Earnings_2020-01`. `Not Classified` é excluído.

## `data/macro_and_ibov_eps_monthly.csv`

| Coluna | Unidade | Definição |
|---|---|---|
| `eps` | pontos de EPS | BEst EPS do Ibovespa no último dia disponível do mês |
| `y_rev3` | % log | Revisão de três meses do BEst EPS |
| `ibov_constant_basket` | nível | Earnings da cesta constante usada no diagnóstico |
| `y_constant_basket_rev3` | % log | Revisão de três meses da cesta constante |
| `focus_pib` | % | Mediana Focus de PIB, janela móvel de 12 meses |
| `d_focus_pib3` | p.p. | Variação de três meses do Focus PIB |
| `focus_ipca` | % | Mediana Focus de IPCA, janela móvel de 12 meses |
| `d_focus_ipca3` | p.p. | Variação de três meses do Focus IPCA |
| `ipca12_available` | % | IPCA acumulado em 12 meses conhecido no fechamento do mês |
| `d_ipca12_3_available` | p.p. | Variação de três meses do IPCA 12m disponível |
| `ipca3_flow_available` | % | IPCA acumulado em três meses disponível |
| `bcom_usd` | índice | Bloomberg Commodity Index em USD |
| `d_bcom_usd3` | % log | Retorno de três meses do BCOM em USD |
| `bcom_brl` | índice | BCOM convertido para BRL |
| `d_bcom_brl3` | % log | Retorno de três meses do BCOM em BRL |
| `usdbrl` | BRL/USD | Última cotação disponível do mês |
| `d_usdbrl3` | % log | Retorno de três meses do USD/BRL |
| `swap360` | % a.a. | Swap pré-DI de 360 dias |
| `d_swap360_3` | p.p. | Variação de três meses do swap |
| `brent_usd` | USD/barril | Última observação diária do Brent no mês |
| `d_brent3` | % log | Retorno de três meses do Brent |

## `data/energy_company_earnings_12m_fwd_monthly.csv`

- Petrobras aparece uma única vez. As linhas de classes ON e PN dos grids carregam o mesmo lucro da empresa e foram deduplicadas.
- `Ex-Petrobras dynamic` soma os emissores não Petrobras disponíveis em cada bloco histórico. A composição varia.
- `Ex-Petrobras core (PRIO + Brava)` exige observação simultânea das duas empresas e começa em 2022.

## Resultados econométricos

Campos recorrentes:

| Campo | Definição |
|---|---|
| `beta` | Coeficiente na unidade original |
| `beta_std` | Coeficiente após padronizar alvo e regressores |
| `se_hac` | Erro-padrão HAC/Newey-West |
| `t_hac` | Estatística t usando o erro HAC |
| `p_hac` | p-valor bilateral aproximado pela normal |
| `q_bh` | p-valor ajustado por Benjamini-Hochberg |
| `partial_r2` | Diferença de R² ao retirar a variável do modelo completo |
| `r2` | Coeficiente de determinação da regressão |
| `adj_r2` | R² ajustado |

Os q-values setoriais são calculados separadamente para cada driver e amostra entre os 11 setores GICS. A linha IBOV não entra nessa correção.
