# Update every field in a .docx through hidden Word, save it, export a PDF, and report table page ranges.
# usage: powershell -NoProfile -ExecutionPolicy Bypass -File word_pass.ps1 -In C:\full\report.docx -Out C:\full\report.pdf [-Json C:\full\layout.json]
# Run it on a scratch copy, never on a file the user has open. It closes only the Word process it started.
param([Parameter(Mandatory=$true)][string]$In, [Parameter(Mandatory=$true)][string]$Out, [string]$Json = "")
$ErrorActionPreference = "Stop"
if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }   # a failed run must not leave an old PDF for the checker
$dir = Split-Path $In -Parent; $name = Split-Path $In -Leaf
if (Test-Path -LiteralPath (Join-Path $dir ("~$" + $name.Substring([Math]::Min(2, $name.Length - 1))))) {
  Write-Output "LOCKED: '$name' is open in Word. Copy it to a scratch folder and run on the copy."; exit 3
}
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$w = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
try {
  $w.Visible = $false; $w.DisplayAlerts = 0
  $d = $w.Documents.Open($In, $false, $false, $false)
  $d.Fields.Update() | Out-Null
  foreach ($t in $d.TablesOfContents) { $t.Update() }
  foreach ($t in $d.TablesOfFigures) { $t.Update() }
  $d.Repaginate()
  foreach ($t in $d.TablesOfContents) { $t.UpdatePageNumbers() }
  foreach ($t in $d.TablesOfFigures) { $t.UpdatePageNumbers() }
  $d.Save()
  $d.ExportAsFixedFormat($Out, 17)
  $rows = @(); $i = 0; $splits = 0
  foreach ($t in $d.Tables) {
    $i++
    $s = $t.Range.Characters.First.Information(3); $e = $t.Range.Characters.Last.Information(3)
    $txt = ($t.Range.Text -replace "[\r\n\a]", " "); if ($txt.Length -gt 50) { $txt = $txt.Substring(0, 50) }
    if ($s -ne $e) { $splits++; Write-Output ("SPLIT table {0}: physical pages {1}-{2}: {3}" -f $i, $s, $e, $txt) }
    $rows += [pscustomobject]@{ n = $i; start = $s; end = $e; text = $txt }
  }
  $pages = $d.ComputeStatistics(2)
  Write-Output ("pages {0}, tables {1}, split tables {2}" -f $pages, $i, $splits)
  if ($Json) { [pscustomobject]@{ pages = $pages; splits = $splits; tables = $rows } | ConvertTo-Json -Depth 4 | Out-File -Encoding utf8 $Json }
  $d.Close(0)
} finally {
  try { $w.Quit() } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
  Start-Sleep -Milliseconds 800
  foreach ($id in $mine) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
