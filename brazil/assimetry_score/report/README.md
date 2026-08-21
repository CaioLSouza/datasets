# Bottom Fishing Brazilian Equities — deck

> **Rascunho de trabalho, não é o report publicado.** A Seção 1 ainda está com placeholders (`Lorem Ipsum` no sumário executivo e `Chart X` nas legendas) e o texto de corpo dela não foi escrito. Só a Seção 2 está completa. Não circular como material final da XP Research.

Deck em A4 vertical no template XP Research, data-base **12/08/2026**.

## Estado por seção

| Slides | Seção | Estado |
| --- | --- | --- |
| 1 | Capa | Pronta |
| 2 | Sumário executivo | **Placeholder** (`Lorem Ipsum`) |
| 3 | Divisória — *Measuring divergence across Brazilian equities* | Pronta |
| 4–7 | Seção 1 — contexto de mercado, score por ação, score por setor, índice e componentes | Gráficos e tabelas posicionados; **corpo de texto vazio e legendas sem numeração** |
| 8 | Divisória — *Is it a good strategy to buy the losers and sell the winners?* | Pronta |
| 9–13 | Seção 2 — regime, convergência, regra, resultado e fragilidade | **Completa** |
| 14 | Divisória — Appendix | Pronta |

## Seção 2

Cinco slides, 12 gráficos nativos do PowerPoint e 3 tabelas, numerados de Chart 10 / Table 3 em diante.

| Slide | Mensagem |
| --- | --- |
| 9 | O trade incondicional não paga; o regime é que decide |
| 10 | A divergência de fato reverte, e em quanto tempo |
| 11 | A regra oficial e o log dos 13 eventos |
| 12 | O que a versão condicional entregou contra o CDI |
| 13 | De onde vem o retorno e quão frágil ele é |

Os dados por trás dos exhibits estão em [`../regime_conditional_spread/`](../regime_conditional_spread/) e em [`../official_35_30_35/`](../official_35_30_35/).

## Ressalvas do conteúdo

Os resultados da Seção 2 são in-sample: pesos, threshold, média móvel, filtro de quality e holding foram escolhidos após múltiplos testes sobre a mesma amostra. São 13 eventos independentes, o intervalo bootstrap de 95% da média por evento inclui zero e o sign test unilateral dá p = 0,133. Custos, aluguel dos shorts, liquidez, capacidade e neutralização de beta e setor não estão modelados. A leitura defensável é de uma oportunidade episódica de convergência, não de um alpha estrutural validado — e é assim que o deck apresenta o resultado.
