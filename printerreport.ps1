<#
.SYNOPSIS
    Build an HTML toner/consumables report for network printers via SNMP (Printer‑MIB).

.DESCRIPTION
    - Reads printers from printerlist.txt (CSV with headers: Value,Name,Description).
    - Uses OlePrn.OleSNMP (ISNMP) to query standard Printer‑MIB OIDs.
    - Correlates supplies by row index and colorant index (no brittle string guessing).
    - Columns: K/C/M/Y (% remaining), Waste (% FULL), Drum (% remaining), Fuser (% remaining),
      Maintenance (% remaining), Pages (lifetime), Pages (7d), and Status.
    - Computes % correctly using Unit/Max/Level and handles -1/-2/-3 as non-numeric.
    - Persists lifetime page counts to CSV to compute 7‑day deltas.
    - Produces C:\...\printerreport.html; optional debug CSV.

.NOTES
    Requires Windows printing components (OlePrn.OleSNMP COM).
    OIDs per RFC 3805 Printer‑MIB (generic/standard):
      prtMarkerSuppliesDescription (.6), SupplyUnit (.7), MaxCapacity (.8), Level (.9),
      Class (.4: 3=supplyThatIsConsumed, 4=receptacleThatIsFilled), Type (.5),
      ColorantIndex (.3), ColorantValue (.12.1.1.4)
      Lifetime pages: prtMarkerLifeCount (.10.2.1.4)
      Alerts: prtAlertDescription (.18.1.1.8)
#>

[CmdletBinding()]
param(
    [string]$PrinterListPath = ".\printerlist.txt",                  # CSV: Value,Name,Description
    [string]$ReportDir       = "DIRECTORY OF THE HTML REPORT",
    [string]$Community       = "public",
    [int]$Retries            = 2,
    [int]$TimeoutMs          = 3000,
    [switch]$OpenWhenDone,
    [switch]$WriteDebugCsv
)

# -------------------------
# Helpers
# -------------------------
function Test-HostUp {
    param([string]$Target)
    try {
        return (Test-Connection -ComputerName $Target -Quiet -Count 1 -ErrorAction SilentlyContinue)
    } catch { return $false }
}

function HtmlEncode {
    param([string]$s)
    if ($null -eq $s) { return "" }
    return ($s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'","&#39;")
}

function Format-PercentCell {
    param([Nullable[int]]$Percent)
    if ($null -eq $Percent) { return "&nbsp;" }
    $color = if ($Percent -gt 49) { "green" }
             elseif ($Percent -gt 24) { "#40BB30" }
             elseif ($Percent -gt 10) { "orange" }
             else { "red" }
    return "<b style='font-size:110%;color:$color;'>$Percent</b>%"
}

# For Waste (we show % FULL so thresholds reverse: high=bad)
function Format-PercentCellWaste {
    param([Nullable[int]]$PercentFull)
    if ($null -eq $PercentFull) { return "&nbsp;" }
    $color = if ($PercentFull -ge 90) { "red" }
             elseif ($PercentFull -ge 75) { "orange" }
             elseif ($PercentFull -ge 50) { "#40BB30" }
             else { "green" }
    return "<b style='font-size:110%;color:$color;'>$PercentFull</b>%"
}

# Plain integer with thousands-separators
function Format-Count {
    param([Nullable[long]]$Value)
    if ($null -eq $Value) { return "&nbsp;" }
    return ("{0:N0}" -f $Value)
}

# Compute % remaining (containers) or % full (receptacles) depending on Class.
function TryGetPercentForSupply {
    <#
      Class 3 (supplyThatIsConsumed): we report % REMAINING
      Class 4 (receptacleThatIsFilled): Level is remaining SPACE -> report % FULL = 100 - remaining%
      Unit=19 (percent) => use Level directly (or with Max=100). Else compute from Max/Level.
      Special values (-1/-2/-3) => $null.
    #>
    param(
        [int]$Class,
        [int]$Unit,
        [int]$Max,
        [int]$Level
    )
    if ($Level -lt 0 -or $Max -eq -2) { return $null }

    if ($Unit -eq 19) {
        if ($Level -ge 0 -and $Level -le 100) {
            $pct = [int][math]::Round($Level)
        } elseif ($Max -eq 100 -and $Level -ge 0) {
            $pct = [int][math]::Min(100,[math]::Round($Level))
        } else { $pct = $null }
        if ($null -eq $pct) { return $null }
        if ($Class -eq 4) { return 100 - $pct }
        return $pct
    }

    if ($Max -gt 0 -and $Level -ge 0) {
        $base = [int][math]::Min(100,[math]::Round(($Level/$Max)*100,0))
        if ($Class -eq 4) { return 100 - $base }
        return $base
    }
    return $null
}

function Map-ColorLetter {
    param([string]$label)
    if ([string]::IsNullOrWhiteSpace($label)) { return $null }
    $d = $label.ToLowerInvariant()
    if ($d -match '\bblack\b'  -or $d -match '\bbk\b' -or $d -match '(^|[^a-z])k([^a-z]|$)') { return 'K' }
    if ($d -match '\bcyan\b'   -or $d -match '(^|[^a-z])c([^a-z]|$)') { return 'C' }
    if ($d -match '\bmagenta\b'-or $d -match '(^|[^a-z])m([^a-z]|$)') { return 'M' }
    if ($d -match '\byellow\b' -or $d -match '(^|[^a-z])y([^a-z]|$)') { return 'Y' }
    return $null
}

# Query all supply rows *by index* and join with colorant names.
function Get-PrinterSupplies {
    param(
        [Parameter(Mandatory)]$Snmp,
        [int[]]$HrRange  = (1..4),
        [int[]]$SupRange = (1..48),
        [int[]]$ColRange = (1..16)
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $colorantsByHr = @{}

    foreach ($hr in $HrRange) {
        # Colorant map
        $map = @{}
        foreach ($ci in $ColRange) {
            try {
                $v = $Snmp.Get(".1.3.6.1.2.1.43.12.1.1.4.$hr.$ci")  # prtMarkerColorantValue.hr.colorIdx
                if ($v) { $map[$ci] = "$v" }
            } catch { break }
        }
        if ($map.Count -gt 0) { $colorantsByHr[$hr] = $map }

        # Supplies rows
        $foundAny = $false
        foreach ($si in $SupRange) {
            try { $desc = $Snmp.Get(".1.3.6.1.2.1.43.11.1.1.6.$hr.$si") } catch { $desc = $null }
            if ([string]::IsNullOrWhiteSpace($desc)) {
                if ($foundAny) { break }
                continue
            }
            $foundAny = $true

            $unit=0; $max=-2; $lvl=-2; $class=0; $type=0; $colIx=0
            try { [int]::TryParse("$($Snmp.Get(".1.3.6.1.2.1.43.11.1.1.7.$hr.$si"))", [ref]$unit)  | Out-Null } catch {}
            try { [int]::TryParse("$($Snmp.Get(".1.3.6.1.2.1.43.11.1.1.8.$hr.$si"))", [ref]$max)   | Out-Null } catch {}
            try { [int]::TryParse("$($Snmp.Get(".1.3.6.1.2.1.43.11.1.1.9.$hr.$si"))", [ref]$lvl)   | Out-Null } catch {}
            try { [int]::TryParse("$($Snmp.Get(".1.3.6.1.2.1.43.11.1.1.4.$hr.$si"))", [ref]$class) | Out-Null } catch {}
            try { [int]::TryParse("$($Snmp.Get(".1.3.6.1.2.1.43.11.1.1.5.$hr.$si"))", [ref]$type)  | Out-Null } catch {}
            try { [int]::TryParse("$($Snmp.Get(".1.3.6.1.2.1.43.11.1.1.3.$hr.$si"))", [ref]$colIx) | Out-Null } catch {}

            $colorLabel = $desc
            if ($colIx -gt 0 -and $colorantsByHr.ContainsKey($hr) -and $colorantsByHr[$hr].ContainsKey($colIx)) {
                $colorLabel = $colorantsByHr[$hr][$colIx]
            }
            $colorLetter = Map-ColorLetter -label $colorLabel

            $percent = TryGetPercentForSupply -Class $class -Unit $unit -Max $max -Level $lvl

            $rows.Add([pscustomobject]@{
                Hr=$hr; Index=$si; Desc=$desc; Unit=$unit; Max=$max; Level=$lvl;
                Class=$class; Type=$type; ColorantIndex=$colIx;
                ColorLabel=$colorLabel; ColorLetter=$colorLetter; Percent=$percent
            })
        }
    }

    return ,$rows
}

# Lifetime page count: return MAX prtMarkerLifeCount across rows (generic approach)
function Get-PrinterPageCount {
    param(
        [Parameter(Mandatory)]$Snmp,
        [int[]]$HrRange  = (1..4),
        [int[]]$MarkRange = (1..48)
    )
    [Nullable[Int64]]$maxPages = $null
    foreach ($hr in $HrRange) {
        foreach ($mi in $MarkRange) {
            try { $v = $Snmp.Get(".1.3.6.1.2.1.43.10.2.1.4.$hr.$mi") } catch { $v = $null } # prtMarkerLifeCount
            if ($v -match '^\d+$') {
                $val = [int64]$v
                if ($val -ge 0 -and ($null -eq $maxPages -or $val -gt $maxPages)) {
                    $maxPages = $val
                }
            }
        }
    }
    return $maxPages
}

# -------------------------
# Paths, History and Report setup
# -------------------------
$ReportDir   = (Resolve-Path -LiteralPath $ReportDir).Path
$ReportHtml  = Join-Path $ReportDir "printerreport.html"
$ReportTmp   = "$ReportHtml.tmp"
$ReportBak   = "$ReportHtml.bak.html"
$DebugCsv    = Join-Path $ReportDir "printerreport_debug.csv"
$HistoryCsv  = Join-Path $ReportDir "printerreport_history.csv"   # NEW

# Backup last report if present
if (Test-Path $ReportHtml) {
    Copy-Item -Path $ReportHtml -Destination $ReportBak -Force
}

# Ensure history file exists with header
if (-not (Test-Path $HistoryCsv)) {
    "Timestamp,IP,Pages" | Out-File -FilePath $HistoryCsv -Encoding UTF8
}
# Load history
$history = Import-Csv -LiteralPath $HistoryCsv

# Read printer list
if (-not (Test-Path $PrinterListPath)) {
    throw "Printer list not found at $PrinterListPath"
}
$printerlist = Import-Csv -LiteralPath $PrinterListPath -Header Value,Name,Description

# Start HTML (with sorting UI + JS)
@"
<html>
<head>
<title>Printer Report</title>
<meta charset="utf-8"/>
<style>
  * { font-family:'Trebuchet MS', Arial, sans-serif; }
  table,th,td { border:1px solid #000; border-collapse:collapse; }
  th,td { padding:6px; vertical-align:top; }
  th { background:#f3f3f3; }
  a { color:#5a2dbf; text-decoration:none; font-weight:bold; }
  a:hover { text-decoration:underline; }
  .toolbar { margin: 10px 0 14px 0; display:flex; gap:12px; align-items:center; flex-wrap:wrap; }
  .toolbar label { font-weight:600; }
  .toolbar select, .toolbar button {
      padding:6px 8px; font-size:14px; border:1px solid #bbb; border-radius:4px; background:#fff;
  }
  .toolbar button { cursor:pointer; }
</style>

<script>
(function(){
  function numFromCell(cell, isPercent){
    if(!cell) return Number.POSITIVE_INFINITY;
    var t = (cell.textContent || cell.innerText || "").replace(/\u00A0/g, ' ').trim();
    var m = t.match(/-?\d+(?:[.,]\d+)?/);
    if(!m) return Number.POSITIVE_INFINITY;
    var n = parseFloat(m[0].replace(',', '.'));
    return isNaN(n) ? Number.POSITIVE_INFINITY : n;
  }
  function getSortableRows(tbl){
    var expected = tbl.tHead ? tbl.tHead.rows[0].cells.length : 0;
    var rows = Array.prototype.slice.call(tbl.tBodies[0].rows);
    return rows.filter(function(r){ return r.cells && r.cells.length === expected; });
  }
  function sortTable(tableId, colIndex, direction, isPercent){
    var tbl = document.getElementById(tableId);
    if(!tbl || !tbl.tBodies.length) return;
    var rows = getSortableRows(tbl);
    var asc = (direction === 'asc');
    rows.sort(function(a,b){
      var av = numFromCell(a.cells[colIndex], isPercent);
      var bv = numFromCell(b.cells[colIndex], isPercent);
      return asc ? (av - bv) : (bv - av);
    });
    var tb = tbl.tBodies[0];
    rows.forEach(function(r){ tb.appendChild(r); });
  }
  var colMap = {
    black:{idx:3,pct:true}, cyan:{idx:4,pct:true}, magenta:{idx:5,pct:true}, yellow:{idx:6,pct:true},
    waste:{idx:7,pct:true}, drum:{idx:8,pct:true}, fuser:{idx:9,pct:true}, maintenance:{idx:10,pct:true},
    pages:{idx:11,pct:false}, pages7:{idx:12,pct:false}
  };
  window._printerReportSort = function(){
    var sel = document.getElementById('sort-col');
    var dir = document.getElementById('sort-dir').value;
    if(!sel || !sel.value){ alert('Pick a column to sort.'); return; }
    var meta = colMap[sel.value]; if(!meta) return;
    sortTable('printer-table', meta.idx, dir, meta.pct);
  };
  window.addEventListener('DOMContentLoaded', function(){
    var tbl = document.getElementById('printer-table'); if(!tbl) return;
    var headers = tbl.tHead ? tbl.tHead.rows[0].cells : [];
    var last = {col:-1, dir:'desc'};
    for (var i=0;i<headers.length;i++){
      (function(ix){
        var isNumeric = (ix>=3 && ix<=12);
        if(!isNumeric) return;
        headers[ix].style.cursor = 'pointer';
        headers[ix].title = 'Click to sort';
        headers[ix].addEventListener('click', function(){
          var isPct = (ix>=3 && ix<=10);
          var dir = (last.col===ix && last.dir==='asc') ? 'desc' : 'asc';
          sortTable('printer-table', ix, dir, isPct);
          last = {col:ix, dir:dir};
        });
      })(i);
    }
  });
})();
</script>
</head>
<body>
"@ | Out-File -Encoding UTF8 -FilePath $ReportTmp

# Heading + toolbar
$validCount = ($printerlist.Value | Where-Object {$_ -and ($_ -notlike "-*")}).Count
"Reporting on $validCount printers" | Add-Content $ReportTmp

@"
<div class="toolbar">
  <label for="sort-col">Sort by:</label>
  <select id="sort-col" aria-label="Sort column">
    <option value="" selected disabled>Choose column…</option>
    <option value="black">Black (%)</option>
    <option value="cyan">Cyan (%)</option>
    <option value="magenta">Magenta (%)</option>
    <option value="yellow">Yellow (%)</option>
    <option value="waste">Waste (% full)</option>
    <option value="drum">Drum (%)</option>
    <option value="fuser">Fuser (%)</option>
    <option value="maintenance">Maintenance (%)</option>
    <option value="pages">Pages (count)</option>
    <option value="pages7">Pages (7d)</option>
  </select>

  <label for="sort-dir">Direction:</label>
  <select id="sort-dir" aria-label="Sort direction">
    <option value="asc" selected>Low → High</option>
    <option value="desc">High → Low</option>
  </select>

  <button type="button" onclick="_printerReportSort()">Sort</button>
</div>
"@ | Add-Content $ReportTmp

# Table start
"<table id='printer-table' style='width:100%'>" | Add-Content $ReportTmp
"<thead><tr><th>Description</th><th>Name</th><th>Type</th><th>Black</th><th>Cyan</th><th>Magenta</th><th>Yellow</th><th>Waste</th><th>Drum</th><th>Fuser</th><th>Maintenance</th><th>Pages</th><th>Pages (7d)</th><th>Status</th></tr></thead>" | Add-Content $ReportTmp
"<tbody>" | Add-Content $ReportTmp

# SNMP COM
try {
    $snmp = New-Object -ComObject olePrn.OleSNMP   # ISNMP automation (Microsoft)
} catch {
    throw "Unable to create OlePrn.OleSNMP COM object. Ensure Windows printing components are installed."
}

# Optional debug
if ($WriteDebugCsv) {
    "IP,Hr,Index,Desc,Unit,Max,Level,Class,Type,ColorantIndex,ColorLabel,ColorLetter,Percent" |
        Out-File -FilePath $DebugCsv -Encoding UTF8
}

$index = 0
$now = Get-Date
$weekAgo = $now.AddDays(-7)

foreach ($p in $printerlist) {

    # Section headers for lines starting with '-'
    if ($p.Value -like "-*") {
        "<tr><td colspan='14'><h3>$(HtmlEncode($p.Value.TrimStart('-')))</h3></td></tr>" | Add-Content $ReportTmp
        continue
    }

    $index++
    $ip   = ($p.Value).Trim()
    $name = $p.Name
    $desc = $p.Description

    "<tr>" | Add-Content $ReportTmp
    "<td><b>$(HtmlEncode($desc))</b></td>" | Add-Content $ReportTmp

    if (-not (Test-HostUp $ip)) {
        $href = "<a href=""http://$ip"" target=""_new"">$ip</a>"
        "<td>$href</td><td><b>Offline</b></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>" | Add-Content $ReportTmp
        "</tr>" | Add-Content $ReportTmp
        continue
    }

    $printerType = $null
    $statusText  = "Operational"
    $K=$null; $C=$null; $M=$null; $Y=$null
    $Waste=$null; $Drum=$null; $Fuser=$null; $Maintenance=$null
    [Nullable[Int64]]$Pages = $null
    [Nullable[Int64]]$Pages7 = $null

    try {
        $snmp.Open($ip, $Community, $Retries, $TimeoutMs)

        # Identify printer
        try { $printerType = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1") } catch {}
        $sysName = $null; try { $sysName = $snmp.Get(".1.3.6.1.2.1.1.5.0") } catch {}

        # Supplies (index-aware + colorant-aware)
        $rows = Get-PrinterSupplies -Snmp $snmp

        # Toners (Class=3, Type=3) → K/C/M/Y
        $toners = $rows | Where-Object { $_.Class -eq 3 -and $_.Type -eq 3 }
        foreach ($t in $toners) {
            switch ($t.ColorLetter) {
                'K' { if ($null -eq $K) { $K = $t.Percent } }
                'C' { if ($null -eq $C) { $C = $t.Percent } }
                'M' { if ($null -eq $M) { $M = $t.Percent } }
                'Y' { if ($null -eq $Y) { $Y = $t.Percent } }
            }
        }

        # Waste (explicit waste types; otherwise any Class=4 'waste' row)
        $wasteTypes = @(4,8,14) # wasteToner(4), wasteInk(8), wasteWax(14)
        $wasteRow = $rows | Where-Object {
            ($_.Type -in $wasteTypes) -or
            ($_.Class -eq 4 -and $_.Desc -match '(?i)waste')
        } | Select-Object -First 1
        if ($wasteRow) { $Waste = $wasteRow.Percent }  # already % FULL

        # Drum (OPC / photoconductor)
        $drumRow = $rows | Where-Object {
            ($_.Type -eq 9) -or ($_.Desc -match '(?i)(drum|opc|photoconductor)')
        } | Select-Object -First 1
        if ($drumRow) { $Drum = $drumRow.Percent }      # % remaining

        # Fuser
        $fuserRow = $rows | Where-Object {
            ($_.Type -eq 15) -or ($_.Desc -match '(?i)fuser')
        } | Select-Object -First 1
        if ($fuserRow) { $Fuser = $fuserRow.Percent }   # % remaining

        # Maintenance kit (transfer/cleaner/fuser-related or explicit "maintenance/kit")
        $maintTypes = @(18,19,20,22) # cleanerUnit(18), fuserCleaningPad(19), transferUnit(20), fuserOiler(22)
        $maintRow = $rows | Where-Object {
            ($_.Type -in $maintTypes) -or ($_.Desc -match '(?i)\bmaintenance\b|\bkit\b')
        } | Select-Object -First 1
        if ($maintRow) { $Maintenance = $maintRow.Percent }  # % remaining

        # Lifetime pages (MAX prtMarkerLifeCount across rows)
        $Pages = Get-PrinterPageCount -Snmp $snmp

        # Compute 7-day delta (if a baseline at/before weekAgo exists)
        if ($Pages -ne $null) {
            $histRows = $history | Where-Object { $_.IP -eq $ip } |
                        Select-Object @{n='TS';e={[datetime]::Parse($_.Timestamp)}}, @{n='Pages';e={[int64]$_.Pages}} |
                        Sort-Object TS
            $baseline = $histRows | Where-Object { $_.TS -le $weekAgo } | Sort-Object TS -Descending | Select-Object -First 1
            if ($baseline -and $Pages -ge $baseline.Pages) {
                $Pages7 = $Pages - $baseline.Pages
            } else {
                # No baseline yet (first week) or counter reset/rollover -> show blank
                $Pages7 = $null
            }
        }

        # Write debug rows if requested
        if ($WriteDebugCsv) {
            $rows | ForEach-Object {
                "$ip,$($_.Hr),$($_.Index),""$($_.Desc.Replace('"',''''))"",$($_.Unit),$($_.Max),$($_.Level),$($_.Class),$($_.Type),$($_.ColorantIndex),""$($_.ColorLabel.Replace('"',''''))"",$($_.ColorLetter),$($_.Percent)" |
                    Add-Content -Path $DebugCsv -Encoding UTF8
            }
        }

        # Alerts -> Status
        try {
            $alerts = @($snmp.GetTree(".1.3.6.1.2.1.43.18.1.1.8"))
            $filtered = $alerts | Where-Object { $_ -and $_ -notlike "print*" -and $_ -notlike "*bypass*" }
            if ($filtered.Count -gt 0) { $statusText = ($filtered -join "; ") }
        } catch {}

        # Name/Type cells
        $displayName = if ($sysName) { $sysName } else { $name }
        $href = "<a href=""http://$ip"" target=""_new"">$(HtmlEncode($displayName))</a>"
        "<td>$href</td>" | Add-Content $ReportTmp
        "<td><br/>$(HtmlEncode($printerType))<br/></td>" | Add-Content $ReportTmp

        # Toner cells
        "<td>$(Format-PercentCell $K)</td>" | Add-Content $ReportTmp
        "<td>$(Format-PercentCell $C)</td>" | Add-Content $ReportTmp
        "<td>$(Format-PercentCell $M)</td>" | Add-Content $ReportTmp
        "<td>$(Format-PercentCell $Y)</td>" | Add-Content $ReportTmp

        # Waste (% FULL), Drum/Fuser/Maintenance (% remaining), Pages (lifetime), Pages (7d)
        "<td>$(Format-PercentCellWaste $Waste)</td>" | Add-Content $ReportTmp
        "<td>$(Format-PercentCell $Drum)</td>"       | Add-Content $ReportTmp
        "<td>$(Format-PercentCell $Fuser)</td>"      | Add-Content $ReportTmp
        "<td>$(Format-PercentCell $Maintenance)</td>"| Add-Content $ReportTmp
        "<td>$(Format-Count $Pages)</td>"            | Add-Content $ReportTmp
        "<td>$(Format-Count $Pages7)</td>"           | Add-Content $ReportTmp

        # Status
        "<td><b>$(HtmlEncode($statusText))</b></td>" | Add-Content $ReportTmp

        # Append to history
        if ($Pages -ne $null) {
            "$($now.ToString('o')),$ip,$Pages" | Add-Content -Path $HistoryCsv -Encoding UTF8
        }

    } catch {
        $href = "<a href=""http://$ip"" target=""_new"">$(HtmlEncode(($name,$ip | Where-Object {$_})[0]))</a>"
        "<td>$href</td><td><b>No data</b></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>" | Add-Content $ReportTmp
    } finally {
        try { $snmp.Close() } catch {}
    }

    "</tr>" | Add-Content $ReportTmp
}

"</tbody></table>" | Add-Content $ReportTmp
$stamp = Get-Date -Format "dd/MM HH:mm"
"<h3>$stamp</h3></body></html>" | Add-Content $ReportTmp

# Publish
Move-Item -Path $ReportTmp -Destination $ReportHtml -Force

if ($OpenWhenDone) { Start-Process $ReportHtml }
Write-Host "Report written to: $ReportHtml"
if ($WriteDebugCsv) { Write-Host "Debug rows written to: $DebugCsv" }
``
