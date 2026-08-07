# Pesquisa de Assessores — versão simples

## Arquitetura

1. `PA Base Historica Long.csv` é a fonte de verdade, em formato long.
2. `PA Report.xlsx` recebe o resultado da query em `fact_agg`.
3. `Dashboard`, `Paineis` e `Tendencias` usam apenas `fact_agg` e não dependem da base antiga.

## Atualização

1. Acrescente as respostas normalizadas do novo Forms ao CSV, mantendo as 18 colunas existentes.
2. No `PA Report.xlsx`, abra a consulta `qFactAgg.m` e carregue o resultado na aba `fact_agg`.
3. Use `Dados > Atualizar Tudo`.
4. Confira a aba `Controles` antes de publicar os gráficos.

O report reserva os intervalos de gráficos e mantém as fórmulas nas abas de saída. O query remove o agregado publicado quando já existem respostas individuais para a mesma onda e pergunta; assim, ondas novas são calculadas do bruto e o legado fica preservado como fallback.

## Limite conhecido do histórico

As respostas individuais disponíveis começam em `202307`. As ondas anteriores entram como `agregado_publicado`, porque a fonte legada só preserva os percentuais publicados, não cada resposta individual. Quando os exports antigos forem recuperados, eles podem ser anexados no mesmo schema sem redesenhar o report.
