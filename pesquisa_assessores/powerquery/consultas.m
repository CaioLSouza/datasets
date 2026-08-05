// =====================================================================
//  CONSULTAS POWER QUERY — PA Principal.xlsx
// =====================================================================
//
//  Cole um bloco por vez em:
//     Dados > Obter Dados > De Outras Fontes > Consulta em Branco
//     > (na barra) Editor Avançado > apagar tudo > colar > Concluído
//     > renomear a consulta com o nome indicado no cabeçalho do bloco
//
//  Comece SEMPRE pela consulta `PastaBases`. As outras dependem dela.
//
//  Onde cada uma deve ser carregada:
//     PastaBases ................ Apenas Criar Conexão
//     agregado .................. Tabela, em uma aba nova `agregado`
//     matriz .................... Tabela, em uma aba nova `matriz`
//     raw_norm .................. Tabela, em uma aba nova `raw_norm`
//     report / serie / q_mes .... Tabela, em abas novas de mesmo nome
//     meta ...................... Tabela, em uma aba nova `meta`
//     medias .................... Tabela, em uma aba nova `medias`
//
//  IMPORTANTE: nunca carregue nada por cima das abas `Charts`,
//  `Gráfico capa` ou `Summary`. São elas que sustentam os 22 links
//  OLE dos dois PPTs.
// =====================================================================


// ---------------------------------------------------------------------
// [1] PastaBases                                   (Criar Conexão só)
// ---------------------------------------------------------------------
// Se a pasta mudar de lugar, você troca AQUI e só aqui.
"\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\bases"


// ---------------------------------------------------------------------
// [2] agregado                                     (histórico completo)
// ---------------------------------------------------------------------
// Uma linha por onda x pergunta x alternativa. É a tabela que a aba
// Base vai consultar depois da migração (ver INSTRUCOES.md, Etapa 2).
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
// [3] matriz                                  (chave x onda, retângulo)
// ---------------------------------------------------------------------
// Mesmo formato da aba Base: linha = alternativa, coluna = onda.
// Útil se você preferir INDEX/MATCH em vez de SUMIFS.
let
    Arquivo = PastaBases & "\PA Base Historica.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="matriz", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true])
in
    Cab


// ---------------------------------------------------------------------
// [4] raw_norm                          (a Raw Data limpa — ETAPA 1)
// ---------------------------------------------------------------------
// Mesmo formato largo da Raw Data de hoje, porém com cabeçalho canônico,
// rótulos antigos já convertidos nos atuais e SEMPRE com ';' no fim de
// cada célula de múltipla escolha. É o que faz as fórmulas COUNTIFS
// atuais pararem de perder a última opção marcada.
let
    Arquivo = PastaBases & "\PA Base Historica.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="raw_norm", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),
    Tipado  = Table.TransformColumnTypes(Cab, {{"Survey", Int64.Type}})
in
    Tipado


// ---------------------------------------------------------------------
// [5] report                                   (só o mês do report)
// ---------------------------------------------------------------------
// Já traz pct, pct_ant e delta contra o mês anterior — o delta mensal
// sai de graça, sem fórmula na planilha.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="report", Kind="Sheet"]}[Data],
    Cab     = Table.PromoteHeaders(Aba, [PromoteAllScalars=true]),
    Tipado  = Table.TransformColumnTypes(Cab, {
        {"q_id", type text}, {"ordem", Int64.Type}, {"safra", type text},
        {"pergunta_pt", type text}, {"pergunta_en", type text},
        {"opcao_id", type text}, {"opcao_pt", type text},
        {"opcao_en", type text}, {"chave", type text},
        {"n", Int64.Type}, {"base", Int64.Type},
        {"pct", type number}, {"pct_ant", type number},
        {"delta", type number}, {"fonte", type text}})
in
    Tipado


// ---------------------------------------------------------------------
// [6] serie                        (janela móvel p/ gráficos de linha)
// ---------------------------------------------------------------------
// Alimenta a aba `Gráfico capa` e qualquer gráfico de série temporal.
// A janela é definida em config/config.yaml (parametros.janela_serie).
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
// [7] q_mes                       (slots da pergunta do mês — FIXO)
// ---------------------------------------------------------------------
// ATENÇÃO: esta aba tem layout de endereço fixo. Carregue-a numa aba
// dedicada e NÃO ordene, filtre nem insira linhas nela: os gráficos dos
// slots apontam para endereços absolutos. Ver LEIA-ME.md.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="q_mes", Kind="Sheet"]}[Data]
in
    Aba


// ---------------------------------------------------------------------
// [8] meta                                  (cabeçalho do mês)
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
// [10] respostas                     (grão respondente — OPCIONAL)
// ---------------------------------------------------------------------
// ~124 mil linhas, na aba `respostas` da base histórica. Só carregue se
// for fazer recorte ad-hoc (por região, por perfil de alocação, etc).
// Para o report não é necessário.
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
