// =====================================================================
//  CONSULTAS POWER QUERY — PA Report.xlsx  (a planilha nova)
// =====================================================================
//
//  Cole um bloco por vez em:
//     Dados > Obter Dados > De Outras Fontes > Consulta em Branco
//     > (na barra) Editor Avançado > apagar tudo > colar > Concluído
//     > renomear a consulta com o nome indicado no cabeçalho do bloco
//
//  Comece SEMPRE pela consulta `PastaBases`. As outras dependem dela.
//
//  ---------------------------------------------------------------
//  AS 5 QUE SUSTENTAM O REPORT  (estas são obrigatórias)
//  ---------------------------------------------------------------
//     PastaBases ...... Apenas Criar Conexão
//     paineis ......... Tabela, em uma aba nova `paineis`
//     tendencias ...... Tabela, em uma aba nova `tendencias`
//     q_mes ........... Tabela, em uma aba nova `q_mes`
//     meta ............ Tabela, em uma aba nova `meta`
//     layout .......... Tabela, em uma aba nova `layout`   (referência)
//
//  ---------------------------------------------------------------
//  ATENÇÃO — POR QUE NENHUMA DELAS PROMOVE CABEÇALHO
//  ---------------------------------------------------------------
//  `paineis`, `tendencias` e `q_mes` têm layout de ENDEREÇO FIXO: os
//  gráficos apontam para intervalos absolutos. Promover cabeçalho
//  comeria uma linha e deslocaria tudo em 1.
//
//  Do jeito que está, o Power Query põe um cabeçalho genérico
//  (Column1, Column2, ...) na linha 1 e a linha N da origem cai na
//  linha N+1 da planilha. Esse deslocamento de 1 JÁ ESTÁ APLICADO nos
//  endereços que a aba `layout` mostra — é só copiar de lá.
//
//  Depois de carregar: não ordene, não filtre, não insira e não
//  remova linhas nessas três abas.
// =====================================================================


// ---------------------------------------------------------------------
// [1] PastaBases                                   (Criar Conexão só)
// ---------------------------------------------------------------------
// Se a pasta mudar de lugar, você troca AQUI e só aqui.
"\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\bases"


// ---------------------------------------------------------------------
// [2] paineis                    (um bloco por pergunta — OBRIGATÓRIA)
// ---------------------------------------------------------------------
// É desta aba que sai a maioria dos gráficos do report: um retângulo
// de endereço fixo por pergunta, com pct, pct do mês anterior e delta.
//
// Cada bloco fica no MESMO lugar todo mês, mesmo que a pergunta não
// tenha sido feita naquele mês (aí ele fica vazio). É isso que faz o
// gráfico nunca precisar ser refeito.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="paineis", Kind="Sheet"]}[Data]
in
    Aba


// ---------------------------------------------------------------------
// [3] tendencias                 (séries temporais — OBRIGATÓRIA)
// ---------------------------------------------------------------------
// Uma coluna por série, na ordem de `series_do_report` no config.yaml.
// Alimenta os gráficos de linha, incluindo o da capa.
//
// Coluna A = data. As linhas são sempre `janela_serie`, alinhadas ao
// fim — o mês corrente é sempre a última linha.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="tendencias", Kind="Sheet"]}[Data]
in
    Aba


// ---------------------------------------------------------------------
// [4] q_mes                      (pergunta do mês — OBRIGATÓRIA)
// ---------------------------------------------------------------------
// Quatro slots de endereço fixo. A pergunta que muda todo mês cai
// sozinha no slot 1; o gráfico é montado uma vez e só troca o conteúdo.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="q_mes", Kind="Sheet"]}[Data]
in
    Aba


// ---------------------------------------------------------------------
// [5] meta                       (cabeçalho do mês — OBRIGATÓRIA)
// ---------------------------------------------------------------------
// onda, título em PT e EN, nº de respostas, data de geração.
// Use nos títulos dos slides em vez de digitar o mês na mão.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="meta", Kind="Sheet"]}[Data],
    Renom   = Table.RenameColumns(Aba, {{"Column1", "campo"}, {"Column2", "valor"}})
in
    Renom


// ---------------------------------------------------------------------
// [6] layout                     (os endereços, prontos — REFERÊNCIA)
// ---------------------------------------------------------------------
// A lista de qual intervalo apontar em cada gráfico, já com o
// deslocamento do Power Query aplicado. É daqui que você copia os
// endereços ao montar a planilha — e é aqui que você confere depois,
// se algum dia mexer em `linhas_por_painel` ou `janela_serie`.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="layout", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true])
in
    Cab


// =====================================================================
//  DAQUI PARA BAIXO É OPCIONAL — nada do report depende destas.
//  Carregue só se for fazer análise ad-hoc.
// =====================================================================


// ---------------------------------------------------------------------
// [7] agregado                            (histórico completo, longo)
// ---------------------------------------------------------------------
// Uma linha por onda x pergunta x alternativa, de jul/2023 até hoje.
// Para tabela dinâmica, recorte novo, conferência.
//
// Use SEMPRE a coluna `pct` — é ela que respeita o congelamento:
//   fonte = "publicado"  -> o número que já foi ao ar, intocado
//   fonte = "calculado"  -> onda nova, calculada do bruto
// A coluna `pct_calculado` é só auditoria. Não aponte gráfico para ela.
let
    Arquivo = PastaBases & "\PA Base Historica.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="agregado", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),
    Tipado  = Table.TransformColumnTypes(Cab, {
        {"onda", Int64.Type}, {"q_id", type text}, {"familia", type text},
        {"safra", type text}, {"pergunta_pt", type text},
        {"pergunta_en", type text}, {"opcao_id", type text},
        {"opcao_pt", type text}, {"opcao_en", type text},
        {"chave", type text}, {"ordem_pergunta", Int64.Type},
        {"ordem_opcao", Int64.Type}, {"n", Int64.Type},
        {"base", Int64.Type}, {"pct", type number},
        {"pct_calculado", type number}, {"fonte", type text}})
in
    Tipado


// ---------------------------------------------------------------------
// [8] serie                        (histórico longo, janela do report)
// ---------------------------------------------------------------------
// Mesma janela da aba `tendencias`, mas em formato longo e com TODAS
// as alternativas. Bom para montar um gráfico de linha fora da lista
// de `series_do_report`, ou para tabela dinâmica.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="serie", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),
    Tipado  = Table.TransformColumnTypes(Cab, {
        {"onda", Int64.Type}, {"data", type date}, {"chave", type text},
        {"q_id", type text}, {"opcao_id", type text},
        {"opcao_pt", type text}, {"opcao_en", type text},
        {"pct", type number}})
in
    Tipado


// ---------------------------------------------------------------------
// [9] medias              (médias: sentimento 0-10, Ibovespa, Selic)
// ---------------------------------------------------------------------
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="medias", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),
    Tipado  = Table.TransformColumnTypes(Cab, {
        {"onda", Int64.Type}, {"data", type date}, {"q_id", type text},
        {"safra", type text}, {"n", Int64.Type}, {"media", type number}})
in
    Tipado


// ---------------------------------------------------------------------
// [10] respostas                          (grão respondente, ~124 mil)
// ---------------------------------------------------------------------
// Uma linha por respondente x pergunta x alternativa marcada.
// Para cortes por região, por perfil de alocação, cruzamentos.
let
    Arquivo = PastaBases & "\PA Base Historica.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="respostas", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),
    Tipado  = Table.TransformColumnTypes(Cab, {
        {"onda", Int64.Type}, {"resp_id", type text}, {"q_id", type text},
        {"familia", type text}, {"safra", type text},
        {"opcao_id", type text}, {"opcao_pt", type text},
        {"valor_num", type number}, {"valor_bruto", type text}})
in
    Tipado
