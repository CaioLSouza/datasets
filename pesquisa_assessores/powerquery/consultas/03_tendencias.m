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
