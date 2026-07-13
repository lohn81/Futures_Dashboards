param(
  [string]$WorkbookPath = "C:\Work Hunting\Quotes.xlsx",
  [int]$Port = 8765
)

$ErrorActionPreference = "Stop"
$script:LastQuotesJson = $null
$script:LastReadAt = [DateTime]::MinValue
$script:CacheMs = 900

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
      } elseif ($path -eq "/" -or $path -eq "/winfut_bovespa_dashboard.html" -or $path -eq "/winfut_dashboard.html" -or $path -eq "/wdofut_dashboard.html") {
        $fileName = switch ($path) {
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
