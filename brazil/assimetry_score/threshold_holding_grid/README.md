# Grade de threshold por holding

As duas sensibilidades publicadas no pacote oficial são fatias 1-D: o threshold com o holding fixo em 21 pregões, e o holding com o threshold fixo em +1,50. Cada uma mostra que o parâmetro é ótimo *dado o outro*, o que não justifica a combinação. Esta grade roda as duas dimensões ao mesmo tempo.

Data-base: **12/08/2026**. Cada célula é a especificação oficial completa — universo MLCX + SMLL, score 35/30/35, z-score causal da SMA21, regime alto apenas, long P90 / short P10, quality 80/20 com reposição, 50/50, entrada em D+1 e eventos não sobrepostos — variando apenas o threshold de entrada e o holding.

## O que a grade mostra

Excesso ativo em % a.a., com o número de eventos entre parênteses:

| Threshold | 10d | 15d | **21d** | 30d | 42d | 63d |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| +1,00 (21) | +2,0 | +4,5 | +6,6 | +5,1 | +3,3 | −0,4 |
| +1,25 (17) | −8,3 | −2,5 | −2,1 | −2,0 | +1,8 | −0,1 |
| **+1,50 (13)** | +13,8 | +16,3 | **+16,6** | +11,2 | +11,6 | +6,2 |
| +1,75 (10) | +20,6 | +18,6 | +6,1 | +6,7 | +10,0 | +5,9 |
| +2,00 (9) | +29,9 | +15,0 | +8,9 | +12,3 | +9,5 | +7,4 |

A célula oficial **não é o máximo da grade**. Três linhas são positivas em todos os holdings — +1,50, +1,75 e +2,00 — e entre elas a de +1,50 tem a maior amostra e o comportamento mais regular ao longo do holding, com pico em 21 pregões. Abaixo de +1,50 o resultado desaparece: a linha de +1,25 é negativa em cinco das seis colunas.

O argumento defensável para a escolha é **estabilidade e tamanho de amostra, não máximo**. A vizinhança na descida não é suave, e isso está visível na própria tabela.

O Sharpe do excesso ativo conta a mesma história, com a célula oficial em 1,48 contra −0,14 em +1,25.

## Arquivos

| Arquivo | Conteúdo |
| --- | --- |
| `threshold_holding_grid.csv` | A grade completa, 7 thresholds por 7 holdings: eventos, dias ativos, fração ativa, excesso ativo, Sharpe ativo, excesso de calendário, hit rate, retorno médio por evento e drawdown relativo. |
| `pivot_active_excess_cagr.csv` | Excesso ativo, recorte de exibição do relatório. |
| `pivot_active_excess_cagr_full.csv` | Excesso ativo na grade inteira, incluindo as faixas excluídas. |
| `pivot_active_excess_sharpe.csv` | Sharpe do excesso ativo. |
| `pivot_hit_rate.csv` | Hit rate por evento. |
| `pivot_events.csv` | Número de eventos independentes por célula. |
| `pivot_calendar_excess_cagr.csv` | Excesso sobre o calendário completo. |
| `validation.csv` | Conferência da célula oficial contra o pipeline publicado. |

## Faixas calculadas e não exibidas

Os pivots de exibição cortam thresholds acima de +2,00 e o holding de 5 pregões. Em +2,25 restam 5 eventos e em +2,50 apenas 1, amostras pequenas demais para comparar. O holding de 5 pregões é instável pelo mesmo motivo: gira a carteira toda semana e produz os dois maiores números da grade inteira sobre pouquíssimos dias de retorno. As duas faixas continuam no `threshold_holding_grid.csv` e no pivot `_full`.

## Validação

A célula (+1,50 / 21 pregões) reproduz a especificação oficial exatamente: 13 eventos e excesso ativo de 16,6010%, contra os 16,6010% publicados, e Sharpe de 1,478 contra 1,478. A conferência está em `validation.csv`.

## Ressalvas

Toda a grade é in-sample, sobre a mesma amostra em que os parâmetros foram escolhidos originalmente. O número de eventos cai de 21 para 8 ao longo das linhas, então células vizinhas não são comparáveis em confiança. Custos, aluguel dos shorts, liquidez e neutralização de fatores não estão modelados em nenhuma célula.
