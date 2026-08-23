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
| 9–10 | Seção 2 — a estratégia incondicional e a condicional | **Completa, com corpo escrito** |
| 11 | Divisória — Appendix | Pronta |

## Seção 2

Duas páginas, quatro gráficos nativos do PowerPoint e duas tabelas, numerados de Chart 10 / Table 3 em diante. Cada página trata de uma versão da estratégia, com e sem o filtro de quality, e as duas usam a mesma tabela de seis linhas medida por posição de 21 pregões — virar a página compara indicador a indicador.

| Página | Mensagem | Exhibits |
| --- | --- | --- |
| 9 | Comprar losers e vender winners não paga como estratégia permanente | Chart 10 riqueza vs CDI · Chart 11 excesso ano a ano · Table 3 |
| 10 | Condicionar ao regime de divergência muda a resposta | Chart 12 excesso por evento · Table 4 · Chart 13 caminho no holding |

Números de abertura, por posição de 21 pregões:

| | Incondicional | Condicional |
| --- | ---: | ---: |
| Posições | 138 | 13 |
| Excesso médio | −0,03% | +1,34% |
| Hit rate | 50,0% | 69,2% |
| Pior posição | −7,70% | −2,94% |
| Excesso a.a. | −0,78% | +16,60% |
| Sharpe | −0,03 | 1,48 |

Os dados dos exhibits estão em [`../regime_conditional_spread/`](../regime_conditional_spread/) e em [`../official_35_30_35/`](../official_35_30_35/). A escolha dos parâmetros está em [`../threshold_holding_grid/`](../threshold_holding_grid/), mas não aparece no deck.

## O que a Seção 2 não cobre

A convergência do índice, a grade de escolha de threshold e holding, o log evento a evento e a decomposição por perna foram retirados quando a seção foi reduzida a duas páginas. Os dados continuam publicados aqui; se a seção voltar a crescer, é de lá que os exhibits saem.

## Ressalvas do conteúdo

Os resultados da Seção 2 são in-sample: pesos, threshold, média móvel, filtro de quality e holding foram escolhidos após múltiplos testes sobre a mesma amostra. São 13 eventos independentes, o intervalo bootstrap de 95% da média por posição vai de −0,40% a +3,18% e o sign test unilateral dá p = 0,133. Três eventos respondem por 112% do ganho logarítmico acumulado. Custos, aluguel dos shorts, liquidez, capacidade e neutralização de beta e setor não estão modelados. A leitura defensável é de uma oportunidade episódica de convergência, não de um alpha estrutural validado — e é assim que o deck apresenta o resultado.
