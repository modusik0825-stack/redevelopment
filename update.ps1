# Reads every KakaoTalk export (.txt / .zip) in the "raw" folder next to this script
# and writes data.js, which the summary HTML page loads automatically.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root   = $PSScriptRoot
$rawDir = Get-ChildItem -LiteralPath $root -Directory | Where-Object { $_.Name -like '*raw*' } | Select-Object -First 1
if (-not $rawDir) { Write-Host 'raw folder not found next to update.ps1'; exit 1 }

# Only keep text that can reach 2026-03-22 or later (the page analyses from 2026-05, with 40 days of lead-in)
$cutoff = Get-Date -Year 2026 -Month 3 -Day 22 -Hour 0 -Minute 0 -Second 0
$dateRe = [regex]'(?m)^(?:(\d{4})\. (\d{1,2})\. (\d{1,2})\. |-+ (\d{4})\D+(\d{1,2})\D+(\d{1,2})\D)'
function Get-LastDate([string]$t) {
  $tail = if ($t.Length -gt 200000) { $t.Substring($t.Length - 200000) } else { $t }
  $ms = $dateRe.Matches($tail)
  if ($ms.Count -eq 0) { return $null }
  $m = $ms[$ms.Count - 1]
  if ($m.Groups[1].Success) { return Get-Date -Year ([int]$m.Groups[1].Value) -Month ([int]$m.Groups[2].Value) -Day ([int]$m.Groups[3].Value) }
  return Get-Date -Year ([int]$m.Groups[4].Value) -Month ([int]$m.Groups[5].Value) -Day ([int]$m.Groups[6].Value)
}

$items = New-Object System.Collections.Generic.List[object]
function Add-Item([string]$name, [string]$text) {
  $last = Get-LastDate $text
  if ($last -and $last -lt $cutoff) { Write-Host "  skip (old): $name"; return }
  $items.Add(@{ name = $name; text = $text })
  Write-Host "  ok: $name"
}

$utf8 = New-Object System.Text.UTF8Encoding($false)
foreach ($f in Get-ChildItem -LiteralPath $rawDir.FullName -File) {
  if ($f.Extension -ieq '.txt') {
    Add-Item $f.Name ([IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8))
  } elseif ($f.Extension -ieq '.zip') {
    $z = [IO.Compression.ZipFile]::OpenRead($f.FullName)
    try {
      foreach ($e in $z.Entries) {
        if ($e.FullName -notlike '*.txt') { continue }
        $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
        try { Add-Item ($f.Name + '::' + $e.FullName) $sr.ReadToEnd() } finally { $sr.Dispose() }
      }
    } finally { $z.Dispose() }
  }
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('window.RAW_BUILT=' + [Web.HttpUtility]::JavaScriptStringEncode((Get-Date).ToString('yyyy-MM-dd HH:mm'), $true) + ";`nwindow.RAW_FILES=[")
$first = $true
foreach ($it in $items) {
  if (-not $first) { [void]$sb.Append(",`n") }; $first = $false
  [void]$sb.Append('{name:' + [Web.HttpUtility]::JavaScriptStringEncode($it.name, $true) + ',text:' + [Web.HttpUtility]::JavaScriptStringEncode($it.text, $true) + '}')
}
[void]$sb.Append("];`n")
$out = Join-Path $root 'data.js'
[IO.File]::WriteAllText($out, $sb.ToString(), $utf8)
Write-Host ("data.js written: {0} files, {1:N1} MB" -f $items.Count, ((Get-Item -LiteralPath $out).Length / 1MB))

if ($args -notcontains '-noopen') {
  $html = Get-ChildItem -LiteralPath $root -Filter *.html | Select-Object -First 1
  if ($html) { Start-Process $html.FullName }
}
