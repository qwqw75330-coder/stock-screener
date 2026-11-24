# 從證交所 (上市) 與櫃買中心 (上櫃) OpenAPI 抓本益比與收盤價，把資料嵌入 template.html 產生 index.html
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-Json($url) {
    $r = Invoke-WebRequest -UseBasicParsing $url
    [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()) | ConvertFrom-Json
}
function To-Num($v) {
    $d = 0.0
    if ($null -ne $v -and [double]::TryParse(("$v").Trim().Replace(',', ''), [ref]$d)) { return $d }
    return $null
}

Write-Host '抓取上市資料...'
$twPe  = Get-Json 'https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL'
$twDay = Get-Json 'https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL'
Write-Host '抓取上櫃資料...'
$otPe  = Get-Json 'https://www.tpex.org.tw/openapi/v1/tpex_mainboard_peratio_analysis'
$otDay = Get-Json 'https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes'

Write-Host '抓取公司基本資料（股本）...'
$twInfo = Get-Json 'https://openapi.twse.com.tw/v1/opendata/t187ap03_L'
$otInfo = Get-Json 'https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O'

# 普通股股數 = (實收資本額 - 特別股股數 x 面額) / 面額
function Get-Shares($capital, $preferred, $parText) {
    $cap = To-Num $capital; if (-not $cap) { return $null }
    $par = 10.0
    $m = [regex]::Match("$parText", '[\d.]+')
    if ($m.Success -and [double]$m.Value -gt 0) { $par = [double]$m.Value }
    $pref = To-Num $preferred; if (-not $pref) { $pref = 0 }
    $s = ($cap - $pref * $par) / $par
    if ($s -gt 0) { return $s } else { return $cap / $par }
}
$shares = @{}
foreach ($i in $twInfo) { $shares[$i.'公司代號'] = Get-Shares $i.'實收資本額' $i.'特別股' $i.'普通股每股面額' }
foreach ($i in $otInfo) { $shares[$i.SecuritiesCompanyCode] = Get-Shares $i.'Paidin.Capital.NTDollars' $i.'PreferredStock.shares' $i.ParValueOfCommonStock }

# 市值（億元）
function Get-MarketCap($code, $price) {
    $s = $shares[$code]
    if ($s -and $price) { return [math]::Round($s * $price / 1e8, 1) }
    return $null
}

$twPx = @{}; foreach ($d in $twDay) { $twPx[$d.Code] = $d }
$otPx = @{}; foreach ($d in $otDay) { $otPx[$d.SecuritiesCompanyCode] = $d }

$rows = New-Object System.Collections.ArrayList
foreach ($p in $twPe) {
    $q = $twPx[$p.Code]
    [void]$rows.Add([ordered]@{
        c = $p.Code; n = $p.Name; m = '上市'
        pe = To-Num $p.PEratio; dy = To-Num $p.DividendYield; pb = To-Num $p.PBratio
        px = if ($q) { To-Num $q.ClosingPrice } else { $null }
        ch = if ($q) { To-Num $q.Change } else { $null }
        vol = if ($q) { To-Num $q.TradeVolume } else { $null }
    })
}
foreach ($p in $otPe) {
    $q = $otPx[$p.SecuritiesCompanyCode]
    [void]$rows.Add([ordered]@{
        c = $p.SecuritiesCompanyCode; n = $p.CompanyName; m = '上櫃'
        pe = To-Num $p.PriceEarningRatio; dy = To-Num $p.YieldRatio; pb = To-Num $p.PriceBookRatio
        px = if ($q) { To-Num $q.Close } else { $null }
        ch = if ($q) { To-Num $q.Change } else { $null }
        vol = if ($q) { To-Num $q.TradingShares } else { $null }
    })
}
foreach ($r in $rows) { $r.mc = Get-MarketCap $r.c $r.px }

# 民國日期 (例 1151123) 轉西元
$raw = "$($twPe[0].Date)"
$date = '{0}-{1}-{2}' -f ([int]$raw.Substring(0, $raw.Length - 4) + 1911), $raw.Substring($raw.Length - 4, 2), $raw.Substring($raw.Length - 2)

# ---------- 每月營收、族群 ----------
Write-Host '抓取每月營收...'
$rev = @{}
foreach ($u in 'https://openapi.twse.com.tw/v1/opendata/t187ap05_L', 'https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap05_O') {
    foreach ($i in (Get-Json $u)) {
        $rev[$i.'公司代號'] = @{
            ind  = "$($i.'產業別')".Trim()
            mom  = To-Num $i.'營業收入-上月比較增減(%)'
            yoy  = To-Num $i.'營業收入-去年同月增減(%)'
            cyoy = To-Num $i.'累計營業收入-前期比較增減(%)'
            ym   = "$($i.'資料年月')"
        }
    }
}
$revMonth = ''
foreach ($r in $rows) {
    $v = $rev[$r.c]
    if ($v) {
        $r.ind = $v.ind
        $r.mom = if ($null -ne $v.mom) { [math]::Round($v.mom, 1) } else { $null }
        $r.yoy = if ($null -ne $v.yoy) { [math]::Round($v.yoy, 1) } else { $null }
        $r.cyoy = if ($null -ne $v.cyoy) { [math]::Round($v.cyoy, 1) } else { $null }
        if ($v.ym -gt $revMonth) { $revMonth = $v.ym }
    }
}
# 民國年月 11510 -> 2026-10
if ($revMonth) { $revMonth = '{0}-{1}' -f ([int]$revMonth.Substring(0, $revMonth.Length - 2) + 1911), $revMonth.Substring($revMonth.Length - 2) }

# ---------- 歷史收盤價快取（近 60 個交易日） ----------
$histFile = Join-Path $PSScriptRoot 'history.json'
$days = @{}; $closed = @{}
if (Test-Path $histFile) {
    $h = [IO.File]::ReadAllText($histFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
    foreach ($c in $h.closed) { $closed["$c"] = $true }
    foreach ($dp in $h.days.PSObject.Properties) {
        $p = @{}; foreach ($x in $dp.Value.p.PSObject.Properties) { $p[$x.Name] = [double]$x.Value }
        $days[$dp.Name] = @{ i = [double]$dp.Value.i; p = $p }
    }
}

function Get-TradingDay([datetime]$d) {
    $key = $d.ToString('yyyyMMdd')
    $tw = Get-Json "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date=$key&type=ALLBUT0999&response=json"
    if ($tw.stat -ne 'OK' -or -not $tw.tables) { return $null }
    $p = @{}; $idx = $null
    foreach ($tb in $tw.tables) {
        if (-not $tb.fields) { continue }
        $f = @($tb.fields)
        $ci = [array]::IndexOf($f, '收盤價')
        if ($f[0] -eq '證券代號' -and $ci -ge 0) {
            foreach ($row in $tb.data) { $v = To-Num $row[$ci]; if ($v) { $p["$($row[0])"] = $v } }
        }
        if ($f[0] -eq '指數' -and -not $idx) {
            foreach ($row in $tb.data) { if ($row[0] -eq '發行量加權股價指數') { $idx = To-Num $row[1] } }
        }
    }
    Start-Sleep -Milliseconds 800
    $ot = Get-Json ("https://www.tpex.org.tw/www/zh-tw/afterTrading/otc?date=" + [uri]::EscapeDataString($d.ToString('yyyy/MM/dd')) + "&type=EW&response=json")
    foreach ($tb in $ot.tables) {
        if (-not $tb.data) { continue }
        foreach ($row in $tb.data) { $v = To-Num $row[2]; if ($v) { $p["$($row[0])"] = $v } }
    }
    if ($p.Count -lt 500) { return $null }
    return @{ i = $idx; p = $p }
}

$need = 61
$latest = [datetime]::ParseExact($date, 'yyyy-MM-dd', $null)
$d = $latest; $have = 0; $fetched = 0
while ($have -lt $need -and $d -gt $latest.AddDays(-140)) {
    $key = $d.ToString('yyyyMMdd')
    if ($d.DayOfWeek -eq 'Saturday' -or $d.DayOfWeek -eq 'Sunday' -or $closed[$key]) { $d = $d.AddDays(-1); continue }
    if (-not $days[$key]) {
        Write-Host "  抓取 $key 歷史股價..."
        $day = $null
        try { $day = Get-TradingDay $d } catch { Write-Host "  $key 抓取失敗：$_" }
        $fetched++
        Start-Sleep -Seconds 3   # 證交所有頻率限制
        if ($day) { $days[$key] = $day } elseif ($d -lt $latest) { $closed[$key] = $true }
    }
    if ($days[$key]) { $have++ }
    $d = $d.AddDays(-1)
}
$keys = @($days.Keys | Sort-Object -Descending | Select-Object -First $need)
$keep = @{}; foreach ($k in $keys) { $keep[$k] = $days[$k] }
$oldest = ($keys | Sort-Object)[0]
$closedKeep = @($closed.Keys | Where-Object { $_ -ge $oldest })
[IO.File]::WriteAllText($histFile, (@{ days = $keep; closed = $closedKeep } | ConvertTo-Json -Depth 5 -Compress), (New-Object Text.UTF8Encoding $false))
Write-Host "歷史股價：$($keys.Count) 個交易日（本次新抓 $fetched 天）"

# 距近 60 日最高收盤的跌幅，以及同期大盤漲跌
$asc = @($keys | Sort-Object)
$lastIdx = $keep[$asc[-1]].i
foreach ($r in $rows) {
    if (-not $r.px) { continue }
    $hi = 0; $hiKey = $null
    foreach ($k in $asc) { $v = $keep[$k].p[$r.c]; if ($v -and $v -gt $hi) { $hi = $v; $hiKey = $k } }
    if ($r.px -gt $hi) { $hi = $r.px; $hiKey = $asc[-1] }
    if ($hi -gt 0) {
        $r.dd = [math]::Round(($r.px / $hi - 1) * 100, 1)
        $r.hd = $hiKey.Substring(4, 2) + '/' + $hiKey.Substring(6, 2)
        $hiIdx = $keep[$hiKey].i
        if ($hiIdx -and $lastIdx) { $r.rel = [math]::Round($r.dd - ($lastIdx / $hiIdx - 1) * 100, 1) }
    }
}

# 預設條件的錯殺候選名單（給雲端 Claude 每日分析用）
$cands = @($rows | Where-Object {
    $_.pe -gt 0 -and $_.pe -lt 20 -and $_.mc -ge 50 -and $_.vol -ge 500000 -and
    $null -ne $_.dd -and $_.dd -le -20 -and $null -ne $_.rel -and $_.rel -le -10 -and $null -ne $_.mom -and $_.mom -gt 0
} | Sort-Object { $_.dd } | ForEach-Object {
    [ordered]@{ code = $_.c; name = $_.n; market = $_.m; industry = $_.ind; close = $_.px; drawdownPct = $_.dd; highDate = $_.hd
        vsMarketPct = $_.rel; revenueMoMPct = $_.mom; revenueYoYPct = $_.yoy; revenueYtdYoYPct = $_.cyoy; pe = $_.pe; pb = $_.pb; marketCap100M = $_.mc }
})
$candJson = [ordered]@{ date = $date; revenueMonth = $revMonth; criteria = 'PE<20, 市值>=50億, 成交量>=500張, 距60日高點跌>=20%, 比大盤多跌>=10個百分點, 營收月增>0'; candidates = $cands } | ConvertTo-Json -Depth 4
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'candidates.json'), $candJson, (New-Object Text.UTF8Encoding $false))
Write-Host "錯殺候選：$($cands.Count) 檔"

$json = @{ date = $date; revMonth = $revMonth; histDays = $keys.Count; histFrom = $oldest; rows = $rows } | ConvertTo-Json -Depth 4 -Compress
$tpl = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'template.html'), [Text.Encoding]::UTF8)
$out = Join-Path $PSScriptRoot 'index.html'
[IO.File]::WriteAllText($out, $tpl.Replace('/*__DATA__*/', "window.STOCK_DATA = $json;"), (New-Object Text.UTF8Encoding $false))
Write-Host "完成：$($rows.Count) 檔，資料日期 $date -> $out"

# 推送到 GitHub Pages（資料有變才 commit）
$git = 'C:\Program Files\Git\cmd\git.exe'
Set-Location $PSScriptRoot
$ErrorActionPreference = 'Continue'   # git 會把進度訊息寫到 stderr，不能當成錯誤
& $git add index.html template.html update-data.ps1 candidates.json .gitignore
& $git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    & $git commit -q -m "Update data $date"
    & $git pull -q --rebase origin main   # 雲端 Claude 會推送 analysis.json，先合併
    & $git push -q origin main
    if ($LASTEXITCODE -eq 0) { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm') 已推送到 GitHub ($date)" }
    else { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm') 推送失敗" }
} else {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm') 資料沒有變化，不需推送"
}
