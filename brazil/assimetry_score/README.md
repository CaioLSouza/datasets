# Assimetry Score — Brasil

Dados processados do projeto de dispersão/assimetria para o universo histórico diário MLCX + SMLL.

## Versão oficial

O pacote oficial, organizado para construção de gráficos e tabelas, está em [`official_35_30_35/`](official_35_30_35/). Todos os arquivos do pacote estão em Excel e incluem abas de metodologia, dicionário de dados e fontes.

Metodologia principal:

- score geral: 35% high value (earnings yield relativo), 30% high short interest e 35% low momentum;
- dispersão: P90 menos P10 do score geral;
- sinal de referência: média móvel de 21 pregões, z-score causal expanding com defasagem de 1 pregão e mínimo de 252 observações;
- evento: cruzamento de alta acima de +1,5 desvio-padrão;
- carteira: long P90 e short P10, capital 50/50, filtro de qualidade 80/20 com reposição;
- execução: D+1, holding fixo de 21 pregões, eventos não sobrepostos e CDI aplicado ao capital.

Data-base dos resultados oficiais: **12/08/2026**.

## Spread condicionado ao regime

Em [`regime_conditional_spread/`](regime_conditional_spread/) está a mesma carteira construída em **todos** os pregões, e não apenas nas datas de sinal. Ela mede o retorno do long P90 / short P10 sem nenhum filtro de regime e por faixa de z-score, o que isola quanto do resultado vem do regime e quanto vem do basket em si: sempre ligado, ele fica 0,77% a.a. abaixo do CDI, contra um spread médio de +2,71% quando a entrada é feita no cruzamento de +1,5.

## Deck

O deck do report está em [`report/`](report/). É um **rascunho**: apenas a Seção 2 está completa, a Seção 1 ainda tem placeholders.

Os CSVs neste diretório são snapshots anteriores do ranking atual e foram preservados para compatibilidade e rastreabilidade.

