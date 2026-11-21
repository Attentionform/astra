param([string]$Dir)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  Get-ChildItem -LiteralPath (Join-Path $Dir 'in') -Filter *.docx | ForEach-Object {
    $out = Join-Path (Join-Path $Dir 'out') $_.Name
    $doc = $word.Documents.Open($_.FullName, $false, $true)
    $doc.SaveAs2($out, 16)
    $doc.Close($false)
    Write-Output "saved $($_.Name)"
  }
} finally { $word.Quit() }
