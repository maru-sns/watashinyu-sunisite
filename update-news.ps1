# update-news.ps1 — 鹿嶋・神栖・アントラーズ ニュース取得（Google ニュース RSS 集約）
# Windows タスクスケジューラーから 10 分ごとに自動実行される
# 出力: news-data.json
#
# 仕組み: Google ニュース RSS は全媒体（新聞・スポーツ紙・地域メディア）を横断集約するため、
#         「すべてのソース」を1本で取得でき、同一記事の重複も媒体名付きで統一できる。

# --- PowerShell 5.1 は既定で TLS1.0。これが無いと https 取得が「接続が閉じられました」で失敗する ---
[Net.ServicePointManager]::SecurityProtocol = `
  [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

$ErrorActionPreference = 'SilentlyContinue'
$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
$newsJsonPath = Join-Path $root 'news-data.json'
$UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36'

# --- 取得クエリ（カテゴリ → Google ニュース検索語）---
# 各市/チームを別クエリで引き、カテゴリを確定する
$queries = @(
  @{ category='antlers'; q='"鹿島アントラーズ" OR アントラーズ'; take=20 },
  @{ category='kashima'; q='鹿嶋市 OR 鹿島神宮 OR 鹿島港';      take=14 },
  @{ category='kamisu';  q='神栖市';                            take=14 }
)

function Get-IsoDate([string]$s) {
  try { ([DateTime]::Parse($s)).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
  catch { (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
}

# Google ニュースの title は「記事タイトル - 媒体名」。末尾の媒体名を除去
function Clean-Title([string]$title, [string]$source) {
  $t = $title
  if ($source -and $t.EndsWith(" - $source")) {
    $t = $t.Substring(0, $t.Length - " - $source".Length)
  } elseif ($t -match '^(.*) - [^-]+$') {
    $t = $Matches[1]
  }
  return $t.Trim()
}

$collected = @()

foreach ($qd in $queries) {
  $enc = [uri]::EscapeDataString($qd.q)
  $url = "https://news.google.com/rss/search?q=$enc&hl=ja&gl=JP&ceid=JP:ja"
  Write-Host ("取得: {0} ..." -f $qd.category) -NoNewline
  try {
    $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 25 -UserAgent $UA
    $xml  = [xml]$resp.Content
    $items = @($xml.rss.channel.item | Where-Object { $_ } | Select-Object -First $qd.take)
    foreach ($it in $items) {
      $src   = if ($it.source.'#text') { $it.source.'#text' } else { 'Google ニュース' }
      $title = Clean-Title $it.title $src
      if ([string]::IsNullOrWhiteSpace($title)) { continue }
      $collected += [PSCustomObject]@{
        title       = $title
        description = ''           # Google ニュースは要約を持たないため空（媒体名で補う）
        published   = Get-IsoDate $it.pubDate
        url         = $it.link
        image       = ''           # Google ニュースは記事個別画像を持たない（カテゴリ別アイコン表示）
        source      = $src
        category    = $qd.category
        duplicates  = @()
      }
    }
    Write-Host (" {0}件" -f $items.Count) -ForegroundColor Green
  } catch {
    Write-Host (" 失敗: {0}" -f $_.Exception.Message) -ForegroundColor Red
  }
}

# --- 時系列ソート（新しい順）---
$collected = $collected | Sort-Object { [DateTime]::Parse($_.published) } -Descending

# --- 重複統一: タイトル正規化（記号除去・先頭20字）で同一記事をまとめ、媒体を duplicates に集約 ---
$deduped = @()
$seen    = @{}
foreach ($item in $collected) {
  $norm = ($item.title -replace '[^\p{L}\p{N}]','')
  $key  = $norm.Substring(0, [Math]::Min(20, $norm.Length))
  if ($key -and $seen.ContainsKey($key)) {
    $exist = $seen[$key]
    if ($exist.source -ne $item.source -and -not ($exist.duplicates | Where-Object { $_.source -eq $item.source })) {
      $exist.duplicates += @{ source = $item.source; url = $item.url }
    }
  } else {
    $seen[$key] = $item
    $deduped   += $item
  }
}

# --- 上限（最新40件）---
$deduped = $deduped | Select-Object -First 40

# 注: Google ニュースの link は中継URLのため、記事個別の OG 画像は取得できない
#     （全件同じ Google 側画像になる）。画像はカテゴリ別アイコンで表示する。

# --- JSON 保存（1件以上取得できた時だけ上書き。0件ならダミー/前回データを保持）---
if ($deduped.Count -gt 0) {
  $json = $deduped | ConvertTo-Json -Depth 4
  [System.IO.File]::WriteAllText($newsJsonPath, $json, (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ("保存完了: {0}件 -> news-data.json" -f $deduped.Count) -ForegroundColor Green
} else {
  Write-Host "取得0件のため既存データを保持しました" -ForegroundColor Yellow
}
