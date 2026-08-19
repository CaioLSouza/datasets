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

Os CSVs neste diretório são snapshots anteriores do ranking atual e foram preservados para compatibilidade e rastreabilidade.

