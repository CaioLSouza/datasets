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
