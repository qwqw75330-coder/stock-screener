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

$json = @{ date = $date; rows = $rows } | ConvertTo-Json -Depth 4 -Compress
$tpl = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'template.html'), [Text.Encoding]::UTF8)
$out = Join-Path $PSScriptRoot 'index.html'
[IO.File]::WriteAllText($out, $tpl.Replace('/*__DATA__*/', "window.STOCK_DATA = $json;"), (New-Object Text.UTF8Encoding $false))
Write-Host "完成：$($rows.Count) 檔，資料日期 $date -> $out"

# 推送到 GitHub Pages（資料有變才 commit）
$git = 'C:\Program Files\Git\cmd\git.exe'
Set-Location $PSScriptRoot
$ErrorActionPreference = 'Continue'   # git 會把進度訊息寫到 stderr，不能當成錯誤
& $git add index.html template.html update-data.ps1
& $git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    & $git commit -q -m "Update data $date"
    & $git push -q origin main
    if ($LASTEXITCODE -eq 0) { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm') 已推送到 GitHub ($date)" }
    else { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm') 推送失敗" }
} else {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm') 資料沒有變化，不需推送"
}
