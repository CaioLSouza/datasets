# Desenha os graficos numa copia da PA Charts.xlsx.
#
#   powershell -ExecutionPolicy Bypass -File desenhar_graficos.ps1 -Entrada "...\PA Charts.xlsx" -Saida "...\PA Charts com graficos.xlsx"
#
# Recebe a planilha SEM graficos e devolve uma COPIA com um grafico por
# tabela. A original nao e tocada -- as duas versoes saem consistentes porque
# a segunda e feita a partir da primeira.
#
# Tipo, orientacao e cores foram extraidos dos graficos da PA Principal, nao
# escolhidos: serie = coluna empilhada, ranking = barra horizontal, regiao =
# pizza, interesse internacional = coluna agrupada.
#
# A serie e sempre ligada a COLUNA DA TABELA, nunca a um intervalo fixo --
# assim o grafico acompanha a tabela crescendo ou encurtando no Atualizar Tudo.

param(
    [string]$Entrada = "\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\PA Charts.xlsx",
    [string]$Saida   = "\\xpdocs\Research\Equities\Estrategia\Reports\Pesquisa assessores\PA Charts com graficos.xlsx"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Entrada)) { Write-Output "ERRO: nao achei $Entrada"; exit 1 }
if (Test-Path $Saida)         { Write-Output "ERRO: $Saida ja existe. Renomeie ou mova."; exit 1 }

# paleta da PA Principal
$AZUL_ESCURO = 0x442F1F   # 1F2F44 -- BGR no COM
$CINZA       = 0x7C7974   # 74797C
$AMARELO     = 0x00BCFF   # FFBC00
$AZUL_CLARO  = 0xDFB38E   # 8EB3DF
$PRETO       = 0x000000

# xlChartType
$COL_EMPILHADA = 52
$COL_AGRUPADA  = 51
$BARRA         = 57
$PIZZA         = 5
$LINHA_MARCA   = 65

# tabela -> como desenhar. 'series' lista as colunas de valor.
$RECEITAS = @(
  @{ aba='d_regiao';                  tipo=$PIZZA;         cat='rotulo_pt'; series=@('atual');              titulo='Regiao do escritorio' }
  @{ aba='s_alocacao_rv';             tipo=$COL_EMPILHADA; cat='data';      series=@();                     titulo='Alocacao em renda variavel -- ultimos 18 meses' }
  @{ aba='s_proximos_meses';          tipo=$COL_EMPILHADA; cat='data';      series=@();                     titulo='Intencao para os proximos meses -- ultimos 18 meses' }
  @{ aba='d_classes_ativos';          tipo=$BARRA;         cat='rotulo_pt'; series=@('anterior','atual');   titulo='Classes de ativos' }
  @{ aba='d_pct_internacional';       tipo=$BARRA;         cat='rotulo_pt'; series=@('anterior','atual');   titulo='% de clientes no internacional' }
  @{ aba='d_interesse_internacional'; tipo=$COL_AGRUPADA;  cat='rotulo_pt'; series=@('anterior','atual');   titulo='Interesse no internacional' }
  @{ aba='d_riscos_bolsa';            tipo=$BARRA;         cat='rotulo_pt'; series=@('anterior','atual');   titulo='Riscos para a Bolsa' }
  @{ aba='d_setores';                 tipo=$BARRA;         cat='rotulo_pt'; series=@('anterior','atual');   titulo='Setores' }
  @{ aba='d_sentimento';              tipo=$BARRA;         cat='rotulo_pt'; series=@('atual');              titulo='Sentimento (0 a 10)' }
  @{ aba='d_ibovespa_alvo';           tipo=$BARRA;         cat='rotulo_pt'; series=@('anterior','atual');   titulo='Ibovespa esperado' }
  @{ aba='d_apetite_risco';           tipo=$BARRA;         cat='rotulo_pt'; series=@('anterior','atual');   titulo='Apetite a risco' }
  @{ aba='medias';                    tipo=$LINHA_MARCA;   cat='data';      series=@('sentimento_media');   titulo='Sentimento medio' }
  @{ aba='q_mes_1';                   tipo=$BARRA;         cat='rotulo_pt'; series=@('pct');                titulo='Pergunta do mes 1' }
  @{ aba='q_mes_2';                   tipo=$BARRA;         cat='rotulo_pt'; series=@('pct');                titulo='Pergunta do mes 2' }
  @{ aba='q_mes_3';                   tipo=$BARRA;         cat='rotulo_pt'; series=@('pct');                titulo='Pergunta do mes 3' }
  @{ aba='q_mes_4';                   tipo=$BARRA;         cat='rotulo_pt'; series=@('pct');                titulo='Pergunta do mes 4' }
  @{ aba='q_mes_5';                   tipo=$BARRA;         cat='rotulo_pt'; series=@('pct');                titulo='Pergunta do mes 5' }
)

Copy-Item $Entrada $Saida -Force
Write-Output "entrada : $Entrada"
Write-Output "saida   : $Saida"
Write-Output ""

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$feitos = 0; $pulados = 0

try {
  $wb = $excel.Workbooks.Open($Saida)

  # o mes do report, para o subtitulo
  $mes = ""
  try { $mes = $wb.Worksheets.Item("corrente").Cells.Item(2,4).Text } catch { }

  foreach ($r in $RECEITAS) {
    $ws = $null
    try { $ws = $wb.Worksheets.Item($r.aba) } catch { }
    if ($null -eq $ws) { Write-Output ("  {0,-28} aba nao existe -- pulado" -f $r.aba); $pulados++; continue }
    $lo = $ws.ListObjects.Item(1)
    if ($lo.ListRows.Count -eq 0) {
      # Tabela vazia (slot de pergunta do mes nao usado neste mes): NAO cria
      # grafico. Testado: um grafico sem serie nao ganha serie quando a tabela
      # enche, e uma serie apontando para uma celula so nao expande no
      # Atualizar Tudo -- a tabela foi de 0 para 4 linhas e o grafico ficou em
      # 1 ponto. Placeholder vazio seria um grafico quebrado esperando para
      # enganar. No mes em que o slot for usado, copie o grafico do q_mes_1 e
      # troque a tabela de origem.
      Write-Output ("  {0,-28} tabela vazia -- sem grafico (ver comentario)" -f $r.aba)
      $pulados++
      continue
    }

    # nomes das colunas da tabela
    $cols = @(); foreach ($lc in $lo.ListColumns) { $cols += $lc.Name }

    # quais colunas viram serie: as declaradas, ou TODAS as numericas
    # (usado nas series temporais, onde uma coluna por alternativa)
    $sers = $r.series
    if ($sers.Count -eq 0) {
      $sers = @($cols | Where-Object { $_ -notin @('onda','data','ordem','alternativa_id','rotulo_pt','rotulo_en','pergunta','qtd','respondentes','regime','slot') })
    }

    $co = $ws.ChartObjects().Add(430, 8, 480, 300)
    $ch = $co.Chart
    $ch.ChartType = $r.tipo
    while ($ch.SeriesCollection().Count -gt 0) { $ch.SeriesCollection().Item(1).Delete() }

    $i = 0
    $faltando = @($sers | Where-Object { $_ -notin $cols })
    if ($faltando.Count -gt 0) {
      Write-Output ("     AVISO: coluna(s) inexistente(s) na tabela: " + ($faltando -join ', '))
    }
    foreach ($nome in $sers) {
      if ($nome -notin $cols) { continue }
      $lc = $lo.ListColumns($nome)
      if ($null -eq $lc.DataBodyRange) { continue }
      $s = $ch.SeriesCollection().NewSeries()
      $s.Name   = "=" + $lc.Range.Cells.Item(1).Address($true,$true,1,$true)
      $s.Values = $lc.DataBodyRange
      $catcol = $lo.ListColumns($r.cat)
      if ($null -ne $catcol.DataBodyRange) { $s.XValues = $catcol.DataBodyRange }

      # cor: nas series temporais a paleta em ordem; nos d_ o cinza e o
      # mes anterior e o amarelo o atual, como nos graficos da PA Principal
      $cor = switch ($nome) {
        'anterior'         { $CINZA }
        'atual'            { $AMARELO }
        'sentimento_media' { $AZUL_ESCURO }
        default            { @($AZUL_CLARO,$CINZA,$AMARELO,$AZUL_ESCURO,$PRETO)[$i % 5] }
      }
      if ($r.tipo -ne $PIZZA) { $s.Format.Fill.ForeColor.RGB = $cor }
      if ($r.tipo -eq $LINHA_MARCA) { $s.Format.Line.ForeColor.RGB = $cor }
      $i++
    }

    # pizza: uma cor por fatia
    if ($r.tipo -eq $PIZZA -and $ch.SeriesCollection().Count -gt 0) {
      $pts = $ch.SeriesCollection().Item(1).Points()
      $paleta = @($AZUL_CLARO,$CINZA,$AMARELO,$AZUL_ESCURO,$PRETO)
      for ($k = 1; $k -le $pts.Count; $k++) {
        $pts.Item($k).Format.Fill.ForeColor.RGB = $paleta[($k-1) % 5]
      }
    }

    # rotulo de dado em todos -- os graficos da PA Principal usam
    $ch.ApplyDataLabels(2) | Out-Null
    try { $ch.SeriesCollection().Item(1).DataLabels().NumberFormat = "0.0%" } catch { }
    if ($r.tipo -eq $LINHA_MARCA) {
      try { $ch.SeriesCollection().Item(1).DataLabels().NumberFormat = "0.00" } catch { }
    }

    # barra horizontal: maior no topo
    if ($r.tipo -eq $BARRA) { try { $ch.Axes(1).ReversePlotOrder = $true } catch { } }

    $ch.HasTitle = $true
    $ch.ChartTitle.Text = if ($mes) { "$($r.titulo) -- $mes" } else { $r.titulo }
    $ch.ChartTitle.Font.Size = 11
    $ch.ChartTitle.Font.Bold = $true
    $ch.HasLegend = ($ch.SeriesCollection().Count -gt 1)
    if ($ch.HasLegend) { $ch.Legend.Position = -4107 }   # embaixo
    try { $ch.ChartArea.Format.Line.Visible = $false } catch { }
    $co.Name = "g_" + $r.aba

    $feitos++
    Write-Output ("  {0,-28} {1} serie(s), {2} pontos" -f `
      $r.aba, $ch.SeriesCollection().Count, $lo.ListRows.Count)
  }

  $wb.Worksheets.Item(1).Activate()
  $wb.Save(); $wb.Close($false)

  Write-Output ""
  Write-Output "============================================================"
  Write-Output " PRONTO -- $feitos graficos ($pulados pulados)"
  Write-Output ""
  Write-Output " Cada grafico esta na planilha da sua tabela, a direita dela."
  Write-Output " As series estao ligadas as COLUNAS da tabela, entao no"
  Write-Output " Atualizar Tudo elas acompanham a tabela crescendo ou"
  Write-Output " encurtando."
  Write-Output "============================================================"
}
catch { Write-Output ""; Write-Output "ERRO: $($_.Exception.Message)"; exit 1 }
finally {
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
