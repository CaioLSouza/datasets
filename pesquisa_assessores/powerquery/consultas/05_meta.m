// onda, título em PT e EN, nº de respostas, data de geração.
// Use nos títulos dos slides em vez de digitar o mês na mão.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="meta", Kind="Sheet"]}[Data],
    Renom   = Table.RenameColumns(Aba, {{"Column1", "campo"}, {"Column2", "valor"}})
in
    Renom
