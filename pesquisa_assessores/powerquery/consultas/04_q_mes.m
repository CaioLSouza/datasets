// Quatro slots de endereço fixo. A pergunta que muda todo mês cai
// sozinha no slot 1; o gráfico é montado uma vez e só troca o conteúdo.
let
    Arquivo = PastaBases & "\PA Base Mes Atual.xlsx",
    Fonte   = Excel.Workbook(File.Contents(Arquivo), null, true),
    Aba     = Fonte{[Item="q_mes", Kind="Sheet"]}[Data]
in
    Aba
