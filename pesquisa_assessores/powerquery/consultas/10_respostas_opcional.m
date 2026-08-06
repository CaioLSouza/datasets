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
