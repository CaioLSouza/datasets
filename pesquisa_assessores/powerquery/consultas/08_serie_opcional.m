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
