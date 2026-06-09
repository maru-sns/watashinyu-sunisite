<#
  鹿行5市ダッシュボード データ更新スクリプト
  ------------------------------------------------------------
  実行: PowerShell で  .\update-data.ps1
  動作:
    1) 公式ソース（人口PDF・財政力指数PDF・行財政概要PDF×5・農業産出額Excel）を
       sources\ フォルダへ再ダウンロード（最新版に差し替え）
    2) 農業産出額Excel(e-Stat)を解析し、index.html の agri 値を自動更新
    3) index.html の「最終更新日(LAST_UPDATED)」を本日に更新
  注意:
    - 人口/財政力/決算はPDFのため自動解析せず、ダウンロードのみ（中身はClaudeに依頼して更新）
    - 年度が変わり統計IDが変わった場合は、下の $sources のURLを差し替えてください
#>

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
$html = Join-Path $root 'index.html'
$srcDir = Join-Path $root 'sources'
if (-not (Test-Path $srcDir)) { New-Item -ItemType Directory -Path $srcDir | Out-Null }

# ---- ソース定義（新年度版が出たらURLを差し替える）----
$AGRI_YEAR = '令和5年(2023)'
$agriXlsxUrl = 'https://www.e-stat.go.jp/stat-search/file-download?statInfId=000040260590&fileKind=4'
$sources = @{
  'population_2025-04.pdf'      = 'https://www.kokudo.or.jp/service/data/map/ibaraki.pdf'
  'zaiseiryoku_r05.pdf'         = 'https://www.pref.ibaraki.jp/somu/shichoson/zaisei/kofuzei/zaiseiryokusisu/documents/r05zaiseiryokusisuusaisannteigo.pdf'
  'gaikyo_kashima_r5.pdf'       = 'https://www.pref.ibaraki.jp/somu/shichoson/gyosei/gaikyo/r5/documents/18_r5_kashima.pdf'
  'gaikyo_itako_r5.pdf'         = 'https://www.pref.ibaraki.jp/somu/shichoson/gyosei/gaikyo/r5/documents/19_r5_itako.pdf'
  'gaikyo_kamisu_r5.pdf'        = 'https://www.pref.ibaraki.jp/somu/shichoson/gyosei/gaikyo/r5/documents/28_r5_kamisu.pdf'
  'gaikyo_namegata_r5.pdf'      = 'https://www.pref.ibaraki.jp/somu/shichoson/gyosei/gaikyo/r5/documents/29_r5_namegata.pdf'
  'gaikyo_hokota_r5.pdf'        = 'https://www.pref.ibaraki.jp/somu/shichoson/gyosei/gaikyo/r5/documents/30_r5_hokota.pdf'
}

Write-Host '=== 1) ソースを再取得 ===' -ForegroundColor Cyan
foreach ($name in $sources.Keys) {
  $dest = Join-Path $srcDir $name
  try {
    Invoke-WebRequest -Uri $sources[$name] -OutFile $dest -UseBasicParsing -TimeoutSec 60
    Write-Host ("  OK  {0}" -f $name)
  } catch {
    Write-Host ("  NG  {0}  ({1})" -f $name, $_.Exception.Message) -ForegroundColor Yellow
  }
}

Write-Host '=== 2) 農業産出額Excelを解析 ===' -ForegroundColor Cyan
$agriBin = Join-Path $srcDir 'agri.xlsx'
$agri = @{}
try {
  Invoke-WebRequest -Uri $agriXlsxUrl -OutFile $agriBin -UseBasicParsing -TimeoutSec 60
  $tmp = Join-Path $env:TEMP ('agri_' + [guid]::NewGuid().ToString('N'))
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::ExtractToDirectory($agriBin, $tmp)

  [xml]$ss = Get-Content (Join-Path $tmp 'xl\sharedStrings.xml') -Encoding UTF8
  $strings = @()
  foreach ($si in $ss.sst.si) {
    if ($si.t -is [string]) { $strings += $si.t }
    elseif ($si.t.'#text') { $strings += $si.t.'#text' }
    else { $strings += (($si.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '') }
  }
  [xml]$sh = Get-Content (Join-Path $tmp 'xl\worksheets\sheet1.xml') -Encoding UTF8
  $cityMap = @{ '鹿嶋市'='鹿嶋市'; '潮来市'='潮来市'; '神栖市'='神栖市'; '行方市'='行方市'; '鉾田市'='鉾田市' }

  foreach ($row in $sh.worksheet.sheetData.row) {
    $cells = @{}
    foreach ($c in $row.c) {
      $v = $c.v
      if ($c.t -eq 's' -and $v -ne $null) { $v = $strings[[int]$v] }
      $col = ($c.r -replace '[0-9]', '')   # 列文字 A,B,C...
      $cells[$col] = "$v"
    }
    $nameCell = $cells['A']
    if ($nameCell -and $cityMap.ContainsKey($nameCell)) {
      # C列 = 農業産出額計（単位:千万円）→ 億円 = /10
      $oku = [math]::Round(([double]$cells['C']) / 10, 1)
      $agri[$nameCell] = $oku
      Write-Host ("  {0}: {1} 億円" -f $nameCell, $oku)
    }
  }
  Remove-Item $tmp -Recurse -Force
} catch {
  Write-Host ("  農業データ解析に失敗: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
}

Write-Host '=== 3) index.html を更新 ===' -ForegroundColor Cyan
$content = Get-Content $html -Raw -Encoding UTF8
$changed = $false

# 農業産出額を各市ブロック内で置換
foreach ($city in $agri.Keys) {
  $val = $agri[$city]
  $pattern = "(?s)(\{name:'$city'.*?agri:)[\d.]+"
  $new = [regex]::Replace($content, $pattern, { param($m) $m.Groups[1].Value + $val })
  if ($new -ne $content) { $content = $new; $changed = $true }
}

# 最終更新日と農業年次を更新
$today = (Get-Date).ToString('yyyy-MM-dd')
$content = [regex]::Replace($content, 'const LAST_UPDATED = "[^"]*";', "const LAST_UPDATED = `"$today`";")
$content = [regex]::Replace($content, '(agri:")[^"]*(")', "`${1}$AGRI_YEAR`${2}")

Set-Content -Path $html -Value $content -Encoding UTF8 -NoNewline
Write-Host ("  最終更新日 -> {0}" -f $today)
if ($changed) { Write-Host '  農業産出額を更新しました' -ForegroundColor Green }
else { Write-Host '  農業産出額に変更はありませんでした' }

Write-Host ''
Write-Host '完了。PDF系（人口・財政力・決算）は sources\ に保存済み。' -ForegroundColor Cyan
Write-Host '新年度の決算/財政力に更新したい場合は、Claude に「sources のPDFから最新値で index.html を更新して」と依頼してください。'
