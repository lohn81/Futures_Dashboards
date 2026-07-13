param(
  [string]$WorkbookPath = "C:\Work Hunting\Quotes.xlsx",
  [int]$Port = 8765
)

$ErrorActionPreference = "Stop"
$script:LastQuotesJson = $null
$script:LastReadAt = [DateTime]::MinValue
$script:CacheMs = 900
$script:LastTradingViewJson = $null
$script:LastTradingViewReadAt = [DateTime]::MinValue
$script:TradingViewCacheMs = 3600000
$script:TradingViewHistoryPath = Join-Path $PSScriptRoot "tradingview_hourly_snapshots.json"

function Convert-ToNumber($value) {
  if ($null -eq $value) { return $null }
  $text = [string]$value
  $text = $text.Trim()
  if ($text -eq "" -or $text -eq "-") { return $null }
  $text = $text -replace "\s", ""
  if ($text -match "^-?\d{1,3}(,\d{3})+(\.\d+)?$") {
    $text = $text -replace ",", ""
  } elseif ($text -match "^-?\d{1,3}(\.\d{3})+(,\d+)?$") {
    $text = ($text -replace "\.", "") -replace ",", "."
  }
  $number = 0.0
  if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
    return $number
  }
  if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
    return $number
  }
  return $null
}

function Get-ExcelWorkbook {
  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  foreach ($book in $excel.Workbooks) {
    if ($book.FullName -eq $WorkbookPath) {
      return $book
    }
  }
  throw "Workbook is not open in Excel: $WorkbookPath"
}

function Get-QuotesJson {
  $book = Get-ExcelWorkbook
  $sheet = $book.Worksheets.Item(1)
  $range = $sheet.UsedRange
  $rows = $range.Rows.Count
  $cols = $range.Columns.Count
  $headers = @()
  for ($col = 1; $col -le $cols; $col++) {
    $header = [string]$sheet.Cells.Item(1, $col).Text
    $header = $header.Trim()
    if ($header -eq "") { $header = "Column$col" }
    $headers += $header
  }
  $quotes = [ordered]@{}

  for ($row = 2; $row -le $rows; $row++) {
    $asset = [string]$sheet.Cells.Item($row, 1).Text
    $asset = $asset.Trim()
    if ($asset -eq "") { continue }

    $quote = [ordered]@{
      asset = $asset
    }

    for ($col = 2; $col -le $cols; $col++) {
      $header = $headers[$col - 1]
      $text = [string]$sheet.Cells.Item($row, $col).Text
      $number = Convert-ToNumber $text
      $quote[$header] = $text
      if ($null -ne $number) {
        $key = ($header -replace "[^A-Za-z0-9]+", "_").Trim("_").ToLowerInvariant()
        $quote[$key] = $number
      }
    }

    $quote.date = [string]$sheet.Cells.Item($row, 2).Text
    $quote.time = [string]$sheet.Cells.Item($row, 3).Text
    $quote.last = Convert-ToNumber $sheet.Cells.Item($row, 4).Text
    $quote.open = Convert-ToNumber $sheet.Cells.Item($row, 5).Text
    $quote.high = Convert-ToNumber $sheet.Cells.Item($row, 6).Text
    $quote.low = Convert-ToNumber $sheet.Cells.Item($row, 7).Text
    $quote.strike = Convert-ToNumber $sheet.Cells.Item($row, 8).Text
    $quote.trades = Convert-ToNumber $sheet.Cells.Item($row, 9).Text
    $quote.expiration = [string]$sheet.Cells.Item($row, 10).Text

    $quotes[$asset] = $quote
  }

  $payload = [ordered]@{
    updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    workbook = $WorkbookPath
    sheet = $sheet.Name
    columns = $headers
    quotes = $quotes
  }

  return ($payload | ConvertTo-Json -Depth 6)
}

function Get-CachedQuotesJson {
  $now = Get-Date
  $cacheAgeMs = ($now - $script:LastReadAt).TotalMilliseconds
  if ($script:LastQuotesJson -and $cacheAgeMs -lt $script:CacheMs) {
    return $script:LastQuotesJson
  }

  try {
    $fresh = Get-QuotesJson
    $script:LastQuotesJson = $fresh
    $script:LastReadAt = Get-Date
    return $fresh
  } catch {
    if ($script:LastQuotesJson) {
      return $script:LastQuotesJson
    }
    throw
  }
}

function Get-YahooFiveMinuteCandle($symbol, $targetDate) {
  $encoded = [Uri]::EscapeDataString($symbol)
  $url = "https://query1.finance.yahoo.com/v8/finance/chart/$encoded`?interval=1m&range=1d"
  $headers = @{
    "User-Agent" = "Mozilla/5.0"
  }

  try {
    $result = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 8
    $chart = $result.chart.result[0]
    if ($null -eq $chart) {
      return @{ status = "unavailable"; error = "No chart result returned." }
    }

    $startLocal = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day $targetDate.Day -Hour 9 -Minute 0 -Second 0
    $endLocal = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day $targetDate.Day -Hour 9 -Minute 4 -Second 59
    $startUnix = ([DateTimeOffset]$startLocal).ToUnixTimeSeconds()
    $endUnix = ([DateTimeOffset]$endLocal).ToUnixTimeSeconds()
    $timestamps = @($chart.timestamp)
    $quote = $chart.indicators.quote[0]
    $rows = @()

    for ($i = 0; $i -lt $timestamps.Count; $i++) {
      $t = [int64]$timestamps[$i]
      if ($t -lt $startUnix -or $t -gt $endUnix) { continue }
      $open = Convert-ToNumber $quote.open[$i]
      $high = Convert-ToNumber $quote.high[$i]
      $low = Convert-ToNumber $quote.low[$i]
      $close = Convert-ToNumber $quote.close[$i]
      if ($null -eq $open -or $null -eq $high -or $null -eq $low -or $null -eq $close) { continue }
      $rows += [ordered]@{
        time = ([DateTimeOffset]::FromUnixTimeSeconds($t).LocalDateTime).ToString("HH:mm:ss")
        open = $open
        high = $high
        low = $low
        close = $close
      }
    }

    if (!$rows.Count) {
      return @{ status = "no_9am_bar"; error = "No 9:00-9:04:59 bar was available from this source today." }
    }

    return [ordered]@{
      status = "ok"
      sourceSymbol = $symbol
      open = $rows[0].open
      high = ($rows | ForEach-Object { $_.high } | Measure-Object -Maximum).Maximum
      low = ($rows | ForEach-Object { $_.low } | Measure-Object -Minimum).Minimum
      close = $rows[-1].close
      points = $rows.Count
      start = $startLocal.ToString("HH:mm:ss")
      end = $endLocal.ToString("HH:mm:ss")
      rows = $rows
    }
  } catch {
    return @{ status = "error"; error = $_.Exception.Message }
  }
}

function Get-ExternalNineAmCandlesJson {
  $today = Get-Date
  $sources = @(
    @{
      asset = "WINFUT"
      source = "Yahoo Finance"
      sourceSymbol = "^BVSP"
      label = "Ibovespa cash index proxy"
      official = $false
      note = "Reference only. This is not official WINFUT futures data."
    },
    @{
      asset = "WDOFUT"
      source = "Yahoo Finance"
      sourceSymbol = "BRL=X"
      label = "USD/BRL spot proxy"
      official = $false
      note = "Reference only. This is not official WDOFUT futures data."
    }
  )

  $candles = [ordered]@{}
  foreach ($source in $sources) {
    $candle = Get-YahooFiveMinuteCandle $source.sourceSymbol $today
    foreach ($key in $source.Keys) {
      $candle[$key] = $source[$key]
    }
    $candles[$source.asset] = $candle
  }

  return ([ordered]@{
    updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    window = "09:00:00-09:04:59 BRT"
    caveat = "Free public sources normally do not provide official 1-minute WINFUT/WDOFUT futures candles. Values here are external proxies unless marked official."
    candles = $candles
  } | ConvertTo-Json -Depth 8)
}

function Get-TradingViewHistory {
  if (!(Test-Path -LiteralPath $script:TradingViewHistoryPath)) {
    return @()
  }
  try {
    $raw = [IO.File]::ReadAllText($script:TradingViewHistoryPath, [Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    $parsed = $raw | ConvertFrom-Json
    if ($parsed -is [array]) { return @($parsed) }
    return @($parsed)
  } catch {
    return @()
  }
}

function Save-TradingViewHistory($history) {
  $json = @($history) | ConvertTo-Json -Depth 8
  [IO.File]::WriteAllText($script:TradingViewHistoryPath, $json, [Text.Encoding]::UTF8)
}

function Invoke-TradingViewScanner($symbol) {
  $body = @{
    symbols = @{
      tickers = @($symbol)
      query = @{ types = @() }
    }
    columns = @("name", "close", "change", "change_abs", "high", "low", "open", "volume", "Recommend.All", "RSI", "Stoch.K", "Stoch.D")
  } | ConvertTo-Json -Depth 6

  try {
    $result = Invoke-RestMethod -Uri "https://scanner.tradingview.com/brazil/scan" -Method Post -Body $body -ContentType "application/json" -Headers @{ "User-Agent" = "Mozilla/5.0" } -TimeoutSec 10
    if ($result.totalCount -lt 1 -or !$result.data -or !$result.data[0]) {
      return @{ status = "no_scanner_row"; error = "TradingView scanner returned no row for $symbol." }
    }
    $values = @($result.data[0].d)
    return [ordered]@{
      status = "ok"
      name = $values[0]
      close = Convert-ToNumber $values[1]
      change = Convert-ToNumber $values[2]
      changeAbs = Convert-ToNumber $values[3]
      high = Convert-ToNumber $values[4]
      low = Convert-ToNumber $values[5]
      open = Convert-ToNumber $values[6]
      volume = Convert-ToNumber $values[7]
      recommendation = Convert-ToNumber $values[8]
      rsi = Convert-ToNumber $values[9]
      stochasticK = Convert-ToNumber $values[10]
      stochasticD = Convert-ToNumber $values[11]
    }
  } catch {
    return @{ status = "scanner_error"; error = $_.Exception.Message }
  }
}

function Invoke-TradingViewChartSnapshot($asset, $symbol, $url) {
  $scanner = Invoke-TradingViewScanner $symbol
  $page = [ordered]@{
    status = "not_checked"
    httpStatus = $null
    title = $null
    defSymbol = $null
    contentLength = $null
    error = $null
  }

  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $url -Headers @{
      "User-Agent" = "Mozilla/5.0"
      "Accept-Language" = "pt-BR,pt;q=0.9,en;q=0.8"
    } -TimeoutSec 12
    $content = [string]$response.Content
    $title = [regex]::Match($content, "<title>(.*?)</title>", "Singleline").Groups[1].Value
    $defSymbol = [regex]::Match($content, "initData\.defSymbol\s*=\s*`"([^`"]+)`"").Groups[1].Value
    $page.status = "ok"
    $page.httpStatus = [int]$response.StatusCode
    $page.title = $title
    $page.defSymbol = $defSymbol
    $page.contentLength = $content.Length
  } catch {
    $page.status = "page_error"
    $page.error = $_.Exception.Message
  }

  return [ordered]@{
    asset = $asset
    symbol = $symbol
    source = "TradingView"
    sourceUrl = $url
    capturedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    scanner = $scanner
    page = $page
    note = "TradingView chart pages are used as hourly reference snapshots. They are not an official API and may not expose candle data to scripts."
  }
}

function Get-TradingViewSnapshotsJson {
  $now = Get-Date
  $cacheAgeMs = ($now - $script:LastTradingViewReadAt).TotalMilliseconds
  if ($script:LastTradingViewJson -and $cacheAgeMs -lt $script:TradingViewCacheMs) {
    return $script:LastTradingViewJson
  }

  $sources = @(
    @{ asset = "WINFUT"; symbol = "BMFBOVESPA:WIN1!"; url = "https://br.tradingview.com/chart/?symbol=BMFBOVESPA%3AWIN1%21" },
    @{ asset = "WDOFUT"; symbol = "BMFBOVESPA:WDO1!"; url = "https://br.tradingview.com/chart/?symbol=BMFBOVESPA%3AWDO1%21" }
  )
  $history = @(Get-TradingViewHistory)
  $latest = [ordered]@{}

  foreach ($source in $sources) {
    $snapshot = Invoke-TradingViewChartSnapshot $source.asset $source.symbol $source.url
    $history += $snapshot
    $latest[$source.asset] = $snapshot
  }

  $cutoff = (Get-Date).AddDays(-21)
  $history = @($history | Where-Object {
    try { ([datetime]$_.capturedAt) -ge $cutoff } catch { $true }
  })
  Save-TradingViewHistory $history

  $payload = [ordered]@{
    updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    cacheSeconds = 3600
    caveat = "TradingView is not treated as an official API. Scanner values are used only when TradingView returns them; chart pages are stored as hourly reference snapshots."
    latest = $latest
    history = $history
  }
  $script:LastTradingViewJson = ($payload | ConvertTo-Json -Depth 10)
  $script:LastTradingViewReadAt = Get-Date
  return $script:LastTradingViewJson
}

function Write-Response($context, [int]$status, [string]$contentType, [string]$body) {
  $bytes = [Text.Encoding]::UTF8.GetBytes($body)
  $context.Response.StatusCode = $status
  $context.Response.ContentType = $contentType
  $context.Response.ContentEncoding = [Text.Encoding]::UTF8
  $context.Response.Headers.Add("Access-Control-Allow-Origin", "*")
  $context.Response.Headers.Add("Cache-Control", "no-store")
  $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  $context.Response.OutputStream.Close()
}

$listener = [Net.HttpListener]::new()
$prefix = "http://127.0.0.1:$Port/"
$listener.Prefixes.Add($prefix)
$listener.Start()

Write-Host ""
Write-Host "Profit Excel bridge is running."
Write-Host "Workbook: $WorkbookPath"
Write-Host "Dashboard: $prefix"
Write-Host "Quotes API: $($prefix)api/quotes"
Write-Host "Keep Excel and this window open. Press Ctrl+C to stop."
Write-Host ""

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $path = $context.Request.Url.AbsolutePath

    try {
      if ($context.Request.HttpMethod -eq "OPTIONS") {
        Write-Response $context 204 "text/plain" ""
      } elseif ($path -eq "/" -or $path -eq "/winfut_bovespa_dashboard.html" -or $path -eq "/winfut_dashboard.html" -or $path -eq "/wdofut_dashboard.html" -or $path -eq "/pro_daytrade_dashboard.html") {
        $fileName = switch ($path) {
          "/pro_daytrade_dashboard.html" { "pro_daytrade_dashboard.html" }
          "/winfut_dashboard.html" { "winfut_dashboard.html" }
          "/wdofut_dashboard.html" { "wdofut_dashboard.html" }
          default { "winfut_bovespa_dashboard.html" }
        }
        $htmlPath = Join-Path $PSScriptRoot $fileName
        $html = [IO.File]::ReadAllText($htmlPath, [Text.Encoding]::UTF8)
        Write-Response $context 200 "text/html; charset=utf-8" $html
      } elseif ($path -eq "/api/quotes") {
        Write-Response $context 200 "application/json; charset=utf-8" (Get-CachedQuotesJson)
      } elseif ($path -eq "/api/external-9am-candles") {
        Write-Response $context 200 "application/json; charset=utf-8" (Get-ExternalNineAmCandlesJson)
      } elseif ($path -eq "/api/tradingview-snapshots") {
        Write-Response $context 200 "application/json; charset=utf-8" (Get-TradingViewSnapshotsJson)
      } else {
        Write-Response $context 404 "text/plain; charset=utf-8" "Not found"
      }
    } catch {
      $errorBody = (@{
        error = $_.Exception.Message
        line = $_.InvocationInfo.ScriptLineNumber
        command = $_.InvocationInfo.Line
        updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
      } | ConvertTo-Json)
      Write-Response $context 500 "application/json; charset=utf-8" $errorBody
    }
  }
} finally {
  if ($listener.IsListening) { $listener.Stop() }
  $listener.Close()
}
