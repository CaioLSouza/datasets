# Cria a PA Charts.xlsx -- a planilha onde você monta os gráficos.
#
# Roda UMA vez, na instalação. Depois disso a planilha é sua: você insere os
# gráficos, formata, e todo mês é só abrir e apertar Atualizar Tudo.
#
#   powershell -ExecutionPolicy Bypass -File montar_charts.ps1
#   powershell -ExecutionPolicy Bypass -File montar_charts.ps1 -Dados "C:\...\PA Charts Data.xlsx" -Saida "C:\...\PA Charts.xlsx"
#
# O que ela faz: descobre as abas da PA Charts Data.xlsx e cria uma consulta
# Power Query por aba, cada uma carregada como Tabela na sua própria planilha.
# Como a descoberta é automática, tabela nova que o Python passe a gerar não
# exige mexer neste script -- é só rodar de novo num arquivo novo.
#
# O caminho do arquivo de dados aparece UMA vez, na consulta _fonte. Se os
# arquivos mudarem de lugar, é lá que se corrige.

param(
    [string]$Dados = "\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\bases\charts",
    [string]$Saida = "\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\PA Charts.xlsx"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Dados)) {
    Write-Output "ERRO: nao achei a pasta de dados:"
    Write-Output "  $Dados"
    Write-Output "Rode o atualizar.py primeiro."
    exit 1
}
if (Test-Path $Saida) {
    Write-Output "ERRO: $Saida ja existe."
    Write-Output "Este script cria a planilha do zero e apagaria os seus graficos."
    Write-Output "Se e isso que voce quer, renomeie ou mova a atual primeiro."
    exit 1
}

Write-Output "dados : $Dados"
Write-Output "saida : $Saida"
Write-Output ""

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$criadas = 0

try {
    # ---- descobre as tabelas: um CSV por tabela
    $arquivos = Get-ChildItem -Path $Dados -Filter *.csv | Sort-Object Name
    if ($arquivos.Count -eq 0) {
        Write-Output "ERRO: nenhum CSV em $Dados"
        exit 1
    }
    Write-Output "tabelas encontradas: $($arquivos.Count)"
    Write-Output ""

    $wb = $excel.Workbooks.Add()

    foreach ($arq in $arquivos) {
        $nome = $arq.BaseName
        if ($nome.Length -gt 31) { $nome = $nome.Substring(0, 31) }

        # A tipagem e explicita de proposito. Sem ela o Power Query deixa as
        # colunas como "any", e a coluna de data chega na planilha como numero
        # de serie cru (45139) -- o eixo do grafico nao entende.
        # Colunas de texto sao nomeadas; TODO O RESTO e numero, inclusive as
        # colunas das series (cujo cabecalho e o rotulo da alternativa).
        $texto = '{"rotulo_pt","rotulo_en","alternativa_id","serie_id",' +
                 '"pergunta","pergunta_id","regime","mes_pt","mes_en"}'
        $inteiro = '{"onda","ordem","qtd","respondentes"}'
        # Culture="en-US" casa com o formato que o Python escreve: ponto
        # decimal e data ISO. Assim a configuracao regional da maquina nao
        # muda o resultado.
        $m = 'let bruto = Csv.Document(File.Contents("' + $arq.FullName + '"), ' +
             '[Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), ' +
             'cab = Table.PromoteHeaders(bruto, [PromoteAllScalars=true]), ' +
             'sem_vazio = Table.ReplaceValue(cab, "", null, Replacer.ReplaceValue, Table.ColumnNames(cab)), ' +
             'tipado = Table.TransformColumnTypes(sem_vazio, List.Transform(Table.ColumnNames(sem_vazio), (c) => ' +
             'if c = "data" then {c, type date} ' +
             'else if List.Contains(' + $inteiro + ', c) then {c, Int64.Type} ' +
             'else if List.Contains(' + $texto + ', c) then {c, type text} ' +
             'else {c, type number}), "en-US") ' +
             'in tipado'
        $wb.Queries.Add($nome, $m) | Out-Null

        $ws = $wb.Worksheets.Add()
        $ws.Name = $nome

        $conn = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;' +
                'Location=' + $nome + ';Extended Properties=""'
        $lo = $ws.ListObjects.Add(0, $conn, $null, $null, $ws.Range("A1"))
        $lo.Name = "t_" + $nome
        $lo.QueryTable.CommandType = 2
        $lo.QueryTable.CommandText = "SELECT * FROM [$nome]"
        # sem isto, Atualizar Tudo nao mexe nesta tabela
        $lo.QueryTable.BackgroundQuery = $false
        $lo.QueryTable.RefreshStyle = 1        # insere linhas conforme cresce
        $lo.QueryTable.SaveData = $true
        $lo.QueryTable.Refresh($false) | Out-Null

        # formato de exibicao, por nome de coluna. E so cosmetica -- o grafico
        # le o valor -- mas a tabela fica legivel para conferir numero.
        foreach ($lc in $lo.ListColumns) {
            $h = $lc.Name
            $col = $lc.DataBodyRange
            if ($null -eq $col) { continue }
            switch -Regex ($h) {
                # a coluna de data NAO leva formato aqui de proposito: o tipo
                # "date" do Power Query ja faz o Excel formatar. Definir
                # "mmm-yy" por COM num Excel pt-BR sai literal ("ago-yy"),
                # porque aqui o codigo de ano e "aa".
                '^data$'                          { }
                '^(onda|ordem|qtd|respondentes)$' { $col.NumberFormat = "0" }
                '^sentimento_media$'              { $col.NumberFormat = "0.00" }
                '^(ibovespa|ibovespa_media)$'     { $col.NumberFormat = "#,##0" }
                '^(rotulo_pt|rotulo_en|alternativa_id|serie_id|pergunta|regime|mes_pt|mes_en)$' { }
                default                           { $col.NumberFormat = "0.0%" }
            }
        }
        $ws.Rows.Item(1).Font.Bold = $true
        $ws.Columns.Item(3).ColumnWidth = 44

        $criadas++
        Write-Output ("  {0,-28} {1,5} linhas x {2,2} colunas" -f `
            $nome, $lo.ListRows.Count, $lo.ListColumns.Count)
    }

    # remove as planilhas vazias que o Excel cria junto com a pasta nova
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.ListObjects.Count -eq 0 -and $ws.UsedRange.Count -le 1) {
            $ws.Delete()
        }
    }

    $wb.Worksheets.Item(1).Activate()
    $wb.SaveAs($Saida, 51)
    $wb.Close($false)

    Write-Output ""
    Write-Output "============================================================"
    Write-Output " PRONTO -- $criadas tabelas em $Saida"
    Write-Output ""
    Write-Output " Agora e sua vez: abra a planilha e monte os graficos sobre"
    Write-Output " as tabelas. Use as COLUNAS da tabela como origem (nao um"
    Write-Output " intervalo de celulas) -- assim o grafico acompanha quando a"
    Write-Output " tabela cresce ou encurta."
    Write-Output ""
    Write-Output " Dai em diante, todo mes: abrir e Dados > Atualizar Tudo."
    Write-Output "============================================================"
}
catch {
    Write-Output ""
    Write-Output "ERRO: $($_.Exception.Message)"
    exit 1
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
