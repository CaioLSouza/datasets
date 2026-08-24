# Bottom Fishing Brazilian Equities — deck

> **Rascunho de trabalho, não é o report publicado.** O corpo de texto está escrito de ponta a ponta e os exhibits estão numerados, mas nada passou por revisão editorial ou por compliance. Não circular como material final da XP Research.

Deck em A4 vertical no template XP Research, data-base **12/08/2026**. Treze páginas.

## Estado por página

| Página | Conteúdo | Estado |
| --- | --- | --- |
| 1 | Capa | Pronta |
| 2 | Índice | **Escrito** |
| 3 | Divisória — *Measuring divergence across Brazilian equities* | Pronta |
| 4–7 | Seção 1 — contexto de mercado, score por ação, score por setor, índice e componentes | **Escrita**, Charts 1–9 e Tables 1–2 |
| 8 | Divisória — *Is it a good strategy to buy the losers and sell the winners?* | Pronta |
| 9–10 | Seção 2 — o basket sem filtro de regime, e a regra | **Escrita**, Charts 10–13 e Tables 3–6 |
| 11 | Divisória — Appendix | Pronta |
| 12 | Disclaimer | Texto de compliance, não editado |
| 13 | Contracapa | Vazia no template de origem |

## Seção 1

Quatro páginas. O argumento é que o mercado brasileiro se dividiu em dois em 2026 e que a divisão aparece em performance, fluxo, múltiplo, nome a nome e por setor.

Commodities +21,3% e bancos +11,0% contra +9,6% do Ibovespa, enquanto cíclicos domésticos caem 5,9%; R$7,2bn de fluxo estrangeiro para commodities contra R$1,9bn saindo de cíclicos; cíclicos a 1,16x o P/L do Ibovespa, o piso da janela 2020–2026, e small caps a 0,99x. O índice fecha a página 7 em 1,42 z-score com os três componentes elevados juntos, o que é o que o torna medida de regime e não spread de valuation.

## Seção 2

Duas páginas. A página 9 mostra que o basket rodando sempre não paga; a página 10 mostra a regra, o resultado e de onde vêm os dois parâmetros.

A Table 4 resume a seção inteira em seis linhas, da esquerda para a direita:

| | Basket | Rule, ex-quality | Rule |
| --- | ---: | ---: | ---: |
| Posições | 138 | 13 | 13 |
| Excesso médio | −0,03% | +1,11% | +1,34% |
| Hit rate | 50,0% | 61,5% | 69,2% |
| Pior posição | −7,70% | −3,13% | −2,94% |
| Excesso a.a. | −0,78% | +13,48% | +16,60% |
| Sharpe | −0,03 | 1,14 | 1,48 |

Todo exhibit usa a especificação oficial, que inclui o filtro de quality. A coluna `Rule, ex-quality` é a única exceção: ela remove o filtro mantendo o resto fixo, para atribuir 3,1 dos 16,60 pontos a ele sem gastar uma página inteira no assunto.

As Tables 5 e 6 são a grade cruzada de threshold por holding, em excesso ativo e em Sharpe. Elas existem para defender +1,50 e 21 pregões em vez de afirmá-los, e o corpo declara de saída que a célula escolhida **não é a mais alta da grade**: o que a sustenta é ter a maior amostra e a linha mais regular. Os dados estão em [`../threshold_holding_grid/`](../threshold_holding_grid/).

## Dados dos exhibits

[`../regime_conditional_spread/`](../regime_conditional_spread/), [`../threshold_holding_grid/`](../threshold_holding_grid/) e [`../official_35_30_35/`](../official_35_30_35/). O `position_metrics.csv` é a fonte única das Tables 3 e 4.

Os exhibits da Seção 1 vêm de planilhas em share de rede da XP que não estão replicadas aqui; os números citados no corpo foram lidos do cache dos próprios gráficos do deck.

## Ressalvas do conteúdo

Os resultados da Seção 2 são in-sample: pesos, threshold, média móvel, filtro de quality e holding foram escolhidos após múltiplos testes sobre a mesma amostra. São 13 eventos independentes, o intervalo bootstrap de 95% da média por posição vai de −0,40% a +3,18% e o sign test unilateral dá p = 0,133. Três eventos respondem por 112% do ganho logarítmico acumulado. Custos, aluguel dos shorts, liquidez, capacidade e neutralização de beta e setor não estão modelados. A leitura defensável é de uma oportunidade episódica de convergência, não de um alpha estrutural validado — e é assim que o deck apresenta o resultado.
