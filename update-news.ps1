# update-news.ps1 — ニュース・SNS 専用取得スクリプト（軽量・高頻度向け）
# Windows タスクスケジューラーから 5〜10 分ごとに自動実行される

$ErrorActionPreference = 'SilentlyContinue'
$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
$newsJsonPath = Join-Path $root 'news-data.json'
$antlersNewsUrl = 'https://www.antlers.co.jp/news/feed/'
$asahiRssUrl    = 'https://www.asahi.com/rss/asahi_ibaraki.rss'

# ---- ヘルパー ----
function Get-IsoDate([string]$s) {
  try { ([DateTime]::Parse($s)).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
  catch { (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
}

function Get-Category([string]$title, [string]$desc) {
  $t = "$title $desc"
  if ($t -match 'アントラーズ|鹿島.*サッカー|J1|J2|Jリーグ') { return 'antlers' }
  if ($t -match '鹿嶋|鹿島神宮|鹿島港')                       { return 'kashima' }
  if ($t -match '神栖')                                       { return 'kamisu'  }
  return $null
}

function Get-ImageUrl($item, [string]$desc) {
  if ($item.enclosure -and $item.enclosure.url)              { return $item.enclosure.url }
  if ($item.'media:content' -and $item.'media:content'.url)  { return $item.'media:content'.url }
  if ($desc -match 'src="([^"]+\.(jpg|jpeg|png|webp))"')     { return $Matches[1] }
  return ''
}

$newsItems = @()

# ---- 朝日新聞 RSS ----
try {
  $xml = [xml](Invoke-WebRequest -Uri $asahiRssUrl -UseBasicParsing -TimeoutSec 20).Content
  foreach ($item in $xml.rss.channel.item | Select-Object -First 30) {
    $title = $item.title -replace '<[^>]*>',''
    $desc  = $item.description -replace '<[^>]*>',''
    $cat   = Get-Category $title $desc
    if ($cat) {
      $newsItems += [PSCustomObject]@{
        title       = $title
        description = $desc
        published   = Get-IsoDate $item.pubDate
        url         = $item.link
        image       = Get-ImageUrl $item $item.description
        source      = '朝日新聞'
        category    = $cat
        duplicates  = @()
      }
    }
  }
} catch {}

# ---- アントラーズ公式 RSS ----
try {
  $xml = [xml](Invoke-WebRequest -Uri $antlersNewsUrl -UseBasicParsing -TimeoutSec 20).Content
  foreach ($item in $xml.rss.channel.item | Select-Object -First 20) {
    $title = $item.title -replace '<[^>]*>',''
    $desc  = $item.description -replace '<[^>]*>',''
    $newsItems += [PSCustomObject]@{
      title       = $title
      description = $desc
      published   = Get-IsoDate $item.pubDate
      url         = $item.link
      image       = Get-ImageUrl $item $item.description
      source      = 'アントラーズ公式'
      category    = 'antlers'
      duplicates  = @()
    }
  }
} catch {}

# ---- 重複排除・ソート ----
$newsItems = $newsItems | Sort-Object { [DateTime]::Parse($_.published) } -Descending
$deduped   = @()
$seen      = @{}
foreach ($item in $newsItems) {
  $key = ($item.title -replace '[^\p{L}\p{N}]','').Substring(0,[Math]::Min(25,($item.title -replace '[^\p{L}\p{N}]','').Length))
  if ($seen.ContainsKey($key)) {
    $seen[$key].duplicates += @{ source=$item.source; url=$item.url }
  } else {
    $seen[$key] = $item
    $deduped   += $item
  }
}

# ---- JSON 保存（既存データが空の場合はそのまま保持） ----
if ($deduped.Count -gt 0) {
  $json = $deduped | ConvertTo-Json -Depth 3 -Compress:$false
  [System.IO.File]::WriteAllText($newsJsonPath, $json, [System.Text.Encoding]::UTF8)
}
