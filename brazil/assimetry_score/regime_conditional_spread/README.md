# Spread P90/P10 condicionado ao regime de divergência

Este diretório responde a uma pergunta que o pacote oficial não cobria: comprar as ações mais fora de graça e vender as mais concorridas é uma boa estratégia por si só, ou só quando a divergência da Bolsa está em extremo?

O pacote em [`official_35_30_35/`](../official_35_30_35/) constrói a carteira apenas nas 13 datas de sinal. Aqui a mesma construção é aplicada a **todos os 3.017 pregões** do calendário, o que permite medir o retorno do basket sem nenhum filtro de regime e comparar o resultado por faixa de z-score.

Data-base: **12/08/2026**. Período: 15/05/2014 a 12/08/2026.

## Resultado

Rodando de forma contínua, em blocos não sobrepostos de 21 pregões, o long P90 / short P10 **não bate o CDI**: 9,29% a.a. contra 10,16% do CDI, um excesso de −0,78% a.a., hit rate de 49,3% em 138 blocos e drawdown relativo máximo de −26,3%. O retorno médio do spread por bloco é de −0,06%.

O que muda a resposta é o regime. Classificando cada pregão pelo z-score causal da SMA21 do índice de divergência, o retorno forward de 21 pregões do mesmo basket é:

| Faixa de z-score | Pregões | Spread médio 21d | Hit rate |
| --- | ---: | ---: | ---: |
| < −1,5 | 219 | +1,00% | 52,5% |
| −1,5 a −0,5 | 622 | +0,04% | 46,6% |
| −0,5 a +0,5 | 865 | −0,78% | 43,2% |
| +0,5 a +1,5 | 758 | +0,79% | 53,0% |
| > +1,5 | 281 | +0,97% | 55,9% |
| Cruzamentos de entrada | 13 | +2,71% | 69,2% |

O spread é pior no regime calmo e melhor nos extremos. Entrar no pregão em que o índice cruza +1,5, e não em qualquer dia em que o regime já está elevado, é o que separa +0,97% de +2,71%.

## Arquivos

| Arquivo | Conteúdo |
| --- | --- |
| `section2_slide1_charts.xlsx` | Workbook chart-ready dos quatro exhibits do slide 1 da Seção 2: curva de riqueza do basket sempre ligado contra o CDI, distribuição do spread forward de 21 pregões, tabela resumo e contribuição acumulada de cada perna. Abas `Leia_me` e `Dicionario` primeiro. |
| `daily_portfolio_forward_returns.csv` | Uma linha por pregão: pontas selecionadas, retorno forward de 21 pregões do long, do short e do spread, quality médio de cada ponta, z-score bruto e da SMA21 na data e a faixa de regime. |
| `bucket_forward_returns.csv` | Agregação por faixa de z-score: número de pregões, spread médio e mediano, hit rate, retorno médio de cada ponta e desvio-padrão. |
| `always_on_daily.csv` | Série diária da versão sempre ligada: CDI, overlay, retorno total e as curvas de riqueza, pregão a pregão. |
| `always_on_blocks.csv` | Os 138 blocos não sobrepostos de 21 pregões da versão sempre ligada, com retorno de cada perna, do spread, do overlay e do CDI. |
| `always_on_summary.csv` | Métricas consolidadas da versão sempre ligada. |
| `validation.csv` | Conferência da reconstrução contra os eventos oficiais. |

## Metodologia

A construção de carteira é idêntica à da especificação oficial:

- universo: composição histórica diária MLCX + SMLL;
- score geral: 35% valuation (earnings yield relativo), 30% short interest e 35% low momentum;
- pontas: P90 e P10 do score geral, equal-weight;
- quality: filtro 80/20 com reposição, aplicado às duas pontas e fora do score;
- execução: entrada em D+1 e holding fixo de 21 pregões.

A única diferença é que a carteira é montada em todos os pregões, e não apenas nos cruzamentos. As faixas de z-score usam a SMA21 causal (expanding, `shift(1)`, mínimo de 252 observações), a mesma série que define o regime oficial.

O baseline sempre ligado percorre blocos consecutivos e não sobrepostos de 21 pregões desde o início da amostra, de forma que ele seja diretamente comparável ao overlay condicional, que também usa eventos não sobrepostos.

A série é construída pregão a pregão, e não bloco a bloco. Entre um bloco e o seguinte existe um pregão de intervalo, herdado da convenção oficial de eventos não sobrepostos; nos dias sem posição o capital rende CDI, como na especificação oficial. O CDI acumula em todos os dias corridos entre dois pregões — uma segunda-feira carrega três dias de juros — e o CAGR é anualizado por 252 pregões. Com essa convenção o CDI daqui reproduz o publicado no relatório: 10,16% a.a., diferença de 4e-16 na curva de riqueza contra a série oficial.

Os retornos das ações vêm dos preços ajustados com forward fill no calendário do índice, exatamente como no motor oficial. Nas 13 datas de sinal oficiais a reconstrução reproduz `long_return`, `short_underlying_return`, `pure_spread_return` e `overlay_excess_return` com diferença máxima da ordem de 1e-17 — a conferência está em `validation.csv`.

## Ressalvas

- Os resultados são in-sample. Pesos, threshold, média móvel, filtro de quality e holding foram escolhidos após múltiplos testes sobre a mesma amostra.
- As faixas de z-score usam janelas forward sobrepostas. O número de pregões por faixa não é um número de observações independentes, e os retornos dentro de cada faixa são autocorrelacionados.
- A faixa de cruzamentos tem apenas 13 observações independentes. O intervalo bootstrap de 95% da média por evento inclui zero.
- Custos de corretagem, bid-ask, impacto de mercado, aluguel dos shorts, liquidez e capacidade não estão modelados.
- Beta, setor e tamanho não estão neutralizados.
- O motor mantém um rebalanceamento equal-weight diário implícito durante o holding.
