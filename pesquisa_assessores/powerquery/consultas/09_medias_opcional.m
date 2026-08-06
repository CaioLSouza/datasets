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
