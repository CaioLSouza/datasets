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
