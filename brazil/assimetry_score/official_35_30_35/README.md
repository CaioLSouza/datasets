# Pacote oficial 35/30/35

Pacote chart-ready com os resultados oficiais discutidos no projeto. Cada workbook começa pelas abas `Leia_me` e `Dicionario`; as demais abas são tabelas planas, com filtros, datas e formatos numéricos prontos para uso em gráficos e tabelas.

## Arquivos

| Arquivo | Conteúdo principal |
| --- | --- |
| `01_official_overview.xlsx` | KPIs, estado atual, resultados anuais e por subperíodo, sensibilidades de threshold e holding, convergência, quality, custos, drawdown e checks. |
| `02_official_daily_series.xlsx` | Séries diárias do índice, z-score causal, estratégia, benchmarks de threshold/quality, trajetórias do z-score e holding. |
| `03_official_events_and_constituents.xlsx` | Eventos, constituintes long/short, reposições pelo filtro de quality, frequências, concentração e sinais pulados. |
| `04_current_ranking_35_30_35.xlsx` | Ranking atual completo, scores individuais, score geral, P90/P10, ranking equal-weight e relação histórica/atual com quality. |
| `05_weight_optimization_convergence.xlsx` | Grade de pesos, métricas por threshold, eventos da otimização, resultados P90/P10 e validação walk-forward exploratória. |

## Definição oficial

- Universo: composição histórica diária MLCX + SMLL.
- Componentes: high value por earnings yield relativo, high short interest e low momentum ajustado por volatilidade.
- Pesos: 35% valuation, 30% short interest e 35% momentum.
- Índice de dispersão: P90 menos P10 do score geral.
- Sinal principal: SMA21 do índice; z-score expanding causal, `shift(1)`, mínimo de 252 pregões; entrada no cruzamento acima de +1,5.
- Expressão financeira: long P90 / short P10, 50%/50%, filtro de quality 80/20 com reposição, execução D+1, holding de 21 pregões e eventos não sobrepostos.
- Data-base: 12/08/2026.

## Uso e ressalvas

- Os resultados são majoritariamente in-sample e devem ser tratados como evidência de capacidade de convergência, não como promessa de retorno.
- Custos, liquidez, capacidade, beta-neutralização, alpha residual e implementação operacional ainda precisam de validação adicional.
- A aba `Leia_me` de cada workbook identifica os arquivos-fonte processados e o status de cada tabela.
- Os datasets brutos licenciados ou muito volumosos (`factor_zoo.parquet`, `market_data.csv`, composição histórica e quality composite) não são replicados aqui. O pacote contém os outputs processados necessários para reproduzir as tabelas e os gráficos do report oficial.

