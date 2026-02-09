<#
.SYNOPSIS
    Build an HTML toner/consumables report for network printers via SNMP (Printer-MIB).

.DESCRIPTION
    - Reads printers from printerlist.txt (CSV with headers: Value,Name,Description).
    - Uses OlePrn.OleSNMP (ISNMP) to query standard Printer-MIB OIDs.
    - Correlates supplies by row index and colorant index (no brittle string guessing).
    - Columns: 
        * Imaging Kit (% remaining)
        * K/C/M/Y (% remaining) + K/C/M/Y (%/week)
        * Waste (% FULL), Drum (% remaining), Fuser (% remaining), Maintenance (% remaining),
        * Pages (lifetime), Pages (1d), Pages (7d), and Status.
    - Computes % correctly using Unit/Max/Level and handles -1/-2/-3 as non-numeric.
    - Persists lifetime page counts and K/C/M/Y % to CSV to compute 1-day, 7-day, and %/week.
    - Enforces history retention: only the last 180 days are kept.
    - Produces C:\...\printerreport.html; optional debug CSV.

.NOTES
    Requires Windows printing components (OlePrn.OleSNMP COM).
    OIDs per RFC 3805 Printer-MIB (generic/standard):
      prtMarkerSuppliesDescription (.6), SupplyUnit (.7), MaxCapacity (.8), Level (.9),
      Class (.4: 3=supplyThatIsConsumed, 4=receptacleThatIsFilled), Type (.5),
      ColorantIndex (.3), ColorantValue (.12.1.1.4)
      Lifetime pages: prtMarkerLifeCount (.10.2.1.4)
      Alerts: prtAlertDescription (.18.1.1.8)
#>

[CmdletBinding()]
param(
    [string]$PrinterListPath = ".\printerlist.txt",                  # CSV: Value,Name,Description
    [string]$ReportDir       = "C:\Users\IT\Desktop\PS_Printer_Ink_Level_SNMP-master",
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
    # Standard HTML entity encoding (ASCII-safe)
    $s = $s -replace '&', '&amp;'
    $s = $s -replace '<', '&lt;'
    $s = $s -replace '>', '&gt;'
    $s = $s -replace '"', '&quot;'
    $s = $s -replace "'", '&#39;'
    return $s
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

# Display for %/week (neutral color)
function Format-PercentPerWeek {
    param([Nullable[int]]$Value)
    if ($null -eq $Value) { return "&nbsp;" }
    return "<b style='font-size:110%;color:#333;'>$Value</b>%/week"
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

    # Common full names
    if ($d -match '\bblack\b'  -or $d -match '\bmono\b' -or $d -match '\bmonochrome\b') { return 'K' }
    if ($d -match '\bcyan\b')    { return 'C' }
    if ($d -match '\bmagenta\b') { return 'M' }
    if ($d -match '\byellow\b')  { return 'Y' }

    # Single-letter or bracketed letters
    if ($d -match '(^|[^a-z])k([^a-z]|$)' -or $d -match '\bbk\b') { return 'K' }
    if ($d -match '(^|[^a-z])c([^a-z]|$)') { return 'C' }
    if ($d -match '(^|[^a-z])m([^a-z]|$)') { return 'M' }
    if ($d -match '(^|[^a-z])y([^a-z]|$)') { return 'Y' }

    # Lexmark/abbrev variants
    if ($d -match '\bblk\b' -or $d -match '\bblack\s*kit\b' -or $d -match '\bbk\s*(unit|kit|pc)\b') { return 'K' }
    if ($d -match '\bcyn\b' -or $d -match '\bcyan\s*kit\b')    { return 'C' }
    if ($d -match '\bmag\b' -or $d -match '\bmagenta\s*kit\b') { return 'M' }
    if ($d -match '\byel\b' -or $d -match '\byellow\s*kit\b')  { return 'Y' }

    return $null
}
function Is-ImagingLikeSupply {
    param([string]$desc, [int]$type)

    # Type=9 is Drum/Photoconductor in Printer-MIB
    if ($type -eq 9) { return $true }

    if ([string]::IsNullOrWhiteSpace($desc)) { return $false }

    $d = $desc.ToLowerInvariant()

    # Imaging unit/kit, photoconductor, OPC, PC unit/kit, developer unit, photo unit
    return ($d -match '\b(imaging\s*(unit|kit)?)\b' -or
            $d -match '\b(photoconductor|photo-?conductor|opc)\b' -or
            $d -match '\b(pc\s*unit|pc\s*kit)\b' -or
            $d -match '\b(developer\s*(unit|kit))\b' -or
            $d -match '\b(photo\s*unit)\b')
}

function Get-ImagingByColor {
    param([object[]]$rows)

    $result = [ordered]@{ K=$null; C=$null; M=$null; Y=$null; Any=$null }

    if (-not $rows) { return $result }

    $candidates = $rows | Where-Object { Is-ImagingLikeSupply -desc $_.Desc -type $_.Type }

    foreach ($r in $candidates) {
        $letter = $r.ColorLetter
        if (-not $letter) {
            if ($r.ColorLabel) { $letter = Map-ColorLetter -label $r.ColorLabel }
            if (-not $letter -and $r.Desc) { $letter = Map-ColorLetter -label $r.Desc }
        }

        if ($letter -and $r.Percent -ne $null) {
            if ($result[$letter] -eq $null) { $result[$letter] = $r.Percent }
        } elseif ($r.Percent -ne $null) {
            if ($result.Any -eq $null) { $result.Any = $r.Percent }
        }
    }

    return $result
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

# Robust detection for colorless monochrome toner (Lexmark MS510/MS810)
function Is-ColorlessTonerDesc {
    param([string]$desc)
    if ([string]::IsNullOrWhiteSpace($desc)) { return $false }
    $d = $desc.ToLowerInvariant()
    $looksLikeToner = ($d -match '\btoner\b' -or $d -match '\bcartridge\b' -or $d -match '\bctg\b' -or $d -match '\bprint\s*cartridge\b')
    $mentionsAColor = ($d -match '\bblack\b' -or $d -match '\bbk\b' -or $d -match '(^|[^a-z])k([^a-z]|$)' -or
                       $d -match '\bcyan\b' -or $d -match '(^|[^a-z])c([^a-z]|$)' -or
                       $d -match '\bmagenta\b' -or $d -match '(^|[^a-z])m([^a-z]|$)' -or
                       $d -match '\byellow\b' -or $d -match '(^|[^a-z])y([^a-z]|$)')
    return ($looksLikeToner -and -not $mentionsAColor)
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

# === NEW: Cached snapshot helper from history (PS 5.1 safe) ===
function Get-HistorySnapshot {
    param(
        [Parameter(Mandatory)][string]$IP,
        [Nullable[datetime]]$AsOf = $null
    )

    if (-not $script:history) { return $null }

    if (-not $AsOf) {
        $AsOf = [datetime]::MaxValue
    }

    $rows =
        $script:history |
        Where-Object { $_.IP -eq $IP } |
        Select-Object @{
                            n='TS'; e={ try { [datetime]::Parse($_.Timestamp) } catch { $null } }
                        }, @{
                            n='Pages'; e={ if($_.Pages -match '^\d+$'){ [int64]$_.Pages } else { $null } }
                        }, @{
                            n='K'; e={ if($_.K -match '^\d+$'){ [int]$_.K } else { $null } }
                        }, @{
                            n='C'; e={ if($_.C -match '^\d+$'){ [int]$_.C } else { $null } }
                        }, @{
                            n='M'; e={ if($_.M -match '^\d+$'){ [int]$_.M } else { $null } }
                        }, @{
                            n='Y'; e={ if($_.Y -match '^\d+$'){ [int]$_.Y } else { $null } }
                        } |
        Where-Object { $_.TS -ne $null } |
        Sort-Object TS

    if (-not $rows -or $rows.Count -eq 0) { return $null }

    $latest = $rows | Where-Object { $_.TS -le $AsOf } | Select-Object -Last 1
    if (-not $latest) { return $null }

    $baseline7 = $rows | Where-Object { $_.TS -le $latest.TS.AddDays(-7) } | Select-Object -Last 1
    $baseline1 = $rows | Where-Object { $_.TS -le $latest.TS.AddDays(-1) } | Select-Object -Last 1

    [Nullable[Int64]]$Pages7 = $null
    [Nullable[Int64]]$Pages1 = $null
    if ($baseline7 -and $latest.Pages -ne $null -and $baseline7.Pages -ne $null -and $latest.Pages -ge $baseline7.Pages) {
        $Pages7 = $latest.Pages - $baseline7.Pages
    }
    if ($baseline1 -and $latest.Pages -ne $null -and $baseline1.Pages -ne $null -and $latest.Pages -ge $baseline1.Pages) {
        $Pages1 = $latest.Pages - $baseline1.Pages
    }

    function _weekly([Nullable[int]]$latestPct, [Nullable[int]]$baselinePct) {
        if ($latestPct -eq $null -or $baselinePct -eq $null) { return $null }
        $drop = $baselinePct - $latestPct
        if ($drop -lt 0) { $drop = 0 }
        return [int][math]::Min(100, [math]::Round($drop))
    }

    $KWeek = if ($baseline7) { _weekly $latest.K $baseline7.K } else { $null }
    $CWeek = if ($baseline7) { _weekly $latest.C $baseline7.C } else { $null }
    $MWeek = if ($baseline7) { _weekly $latest.M $baseline7.M } else { $null }
    $YWeek = if ($baseline7) { _weekly $latest.Y $baseline7.Y } else { $null }

    return [pscustomobject]@{
        TS      = $latest.TS
        Pages   = $latest.Pages
        Pages1  = $Pages1
        Pages7  = $Pages7
        K       = $latest.K
        C       = $latest.C
        M       = $latest.M
        Y       = $latest.Y
        KWeek   = $KWeek
        CWeek   = $CWeek
        MWeek   = $MWeek
        YWeek   = $YWeek
    }
}

# -------------------------
# Paths, History and Report setup
# -------------------------
$ReportDir   = (Resolve-Path -LiteralPath $ReportDir).Path
$ReportHtml  = Join-Path $ReportDir "printerreport.html"
$ReportTmp   = "$ReportHtml.tmp"
$ReportBak   = "$ReportHtml.bak.html"
$DebugCsv    = Join-Path $ReportDir "printerreport_debug.csv"
$HistoryCsv  = Join-Path $ReportDir "printerreport_history.csv"   # Stores page + color percent history

# Backup last report if present
if (Test-Path $ReportHtml) {
    Copy-Item -Path $ReportHtml -Destination $ReportBak -Force
}

# Ensure history file exists with header (upgraded to include K,C,M,Y)
if (-not (Test-Path $HistoryCsv)) {
    "Timestamp,IP,Pages,K,C,M,Y" | Out-File -FilePath $HistoryCsv -Encoding UTF8
}

# --- History retention: keep last 180 days only ---
$now = Get-Date
$cutoff = $now.AddDays(-180)
try {
    $historyRaw = Import-Csv -LiteralPath $HistoryCsv
} catch {
    $historyRaw = @()
}

# Guarantee columns exist even if old file didn't have them
$historyRaw = $historyRaw | ForEach-Object {
    [pscustomobject]@{
        Timestamp = $_.Timestamp
        IP        = $_.IP
        Pages     = $_.Pages
        K         = $_.K
        C         = $_.C
        M         = $_.M
        Y         = $_.Y
    }
}

if ($historyRaw.Count -gt 0) {
    $historyPruned =
        $historyRaw |
        Where-Object {
            try { [datetime]::Parse($_.Timestamp) -ge $cutoff } catch { $false }
        }

    # Overwrite history with pruned entries (ensure full header)
    if ($historyPruned -is [System.Array] -and $historyPruned.Count -eq 0) {
        "Timestamp,IP,Pages,K,C,M,Y" | Out-File -FilePath $HistoryCsv -Encoding UTF8
    } else {
        $historyPruned | Select-Object Timestamp,IP,Pages,K,C,M,Y |
            Export-Csv -LiteralPath $HistoryCsv -NoTypeInformation -Encoding UTF8
    }

    $history = $historyPruned
} else {
    $history = @()
}
# Make available to helper via script: scope
$script:history = $history

# Read printer list
if (-not (Test-Path $PrinterListPath)) {
    throw "Printer list not found at $PrinterListPath"
}
$printerlist = Import-Csv -LiteralPath $PrinterListPath -Header Value,Name,Description

# Buffer to collect cards markup (so we can write it after the table)
$CardBuffer = New-Object System.Text.StringBuilder
$cardsGridOpen = $false

# =========================
# Start HTML (with sorting UI + JS + View toggle)
# =========================
@"
<html>
<head>
<title>Printer Report</title>
<meta charset="utf-8"/>
<style>
  * { font-family:'Trebuchet MS', Arial, sans-serif; }
  body { margin: 16px; }
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
  .view-switch .btn { padding:6px 10px; border-radius:6px; border:1px solid #bbb; background:#f9f9f9; }
  .view-switch .btn.active { background:#2b3152; color:#fff; border-color:#2b3152; }

  /* Section headers (full row) - table */
  #printer-table td[colspan="20"] {
    background: #2b3152 !important;
    color: #ffffff !important;
  }
  #printer-table td[colspan="20"] h3,
  #printer-table td[colspan="20"] a { color: #ffffff !important; }
  #printer-table td[colspan="20"] h3 { margin: 6px 0; }

  /* Column order (1-based):
     1 Description, 2 Name, 3 Type, 4 Imaging Kit,
     5 Black, 6 Black %/week, 7 Cyan, 8 Cyan %/week,
     9 Magenta, 10 Magenta %/week, 11 Yellow, 12 Yellow %/week,
     13 Waste, 14 Drum, 15 Fuser, 16 Maintenance,
     17 Pages, 18 Pages (1d), 19 Pages (7d), 20 Status
  */

  /* Primary toner % tints (table td only) */
  #printer-table tbody td:nth-child(5)  { background-color: rgba(0, 0, 0, 0.05); }
  #printer-table tbody td:nth-child(7)  { background-color: rgba(0, 174, 239, 0.10); }
  #printer-table tbody td:nth-child(9)  { background-color: rgba(236, 0, 140, 0.10); }
  #printer-table tbody td:nth-child(11) { background-color: rgba(255, 242, 0, 0.20); }

  /* %/week tints (softer) */
  #printer-table tbody td:nth-child(6)  { background-color: rgba(0, 0, 0, 0.03); }
  #printer-table tbody td:nth-child(8)  { background-color: rgba(0, 174, 239, 0.06); }
  #printer-table tbody td:nth-child(10) { background-color: rgba(236, 0, 140, 0.06); }
  #printer-table tbody td:nth-child(12) { background-color: rgba(255, 242, 0, 0.10); }

  /* Supplies tints */
  #printer-table tbody td:nth-child(13) { background-color: rgba(153, 102, 51, 0.12); }
  #printer-table tbody td:nth-child(14) { background-color: rgba(70, 130, 180, 0.12); }
  #printer-table tbody td:nth-child(15) { background-color: rgba(255, 99, 71, 0.12); }
  #printer-table tbody td:nth-child(16) { background-color: rgba(255, 165, 0, 0.12); }

  /* Preserve text coloring set by Format-* helpers */
  #printer-table tbody td:nth-child(5),
  #printer-table tbody td:nth-child(6),
  #printer-table tbody td:nth-child(7),
  #printer-table tbody td:nth-child(8),
  #printer-table tbody td:nth-child(9),
  #printer-table tbody td:nth-child(10),
  #printer-table tbody td:nth-child(11),
  #printer-table tbody td:nth-child(12),
  #printer-table tbody td:nth-child(13),
  #printer-table tbody td:nth-child(14),
  #printer-table tbody td:nth-child(15),
  #printer-table tbody td:nth-child(16) { color: inherit; }

/* Cards view (horizontal scrolling per section) */
.hidden { display:none; }
#cards-view { margin-top: 10px; }
.section-divider {
  background:#2b3152; color:#fff; padding:6px 10px; border-radius:8px; margin:18px 0 10px 0;
}

.cards-grid {
  display: flex;
  flex-wrap: nowrap;
  overflow-x: auto;
  gap: 12px;
  padding-bottom: 8px;
}

/* Make each card a column flexbox so we can control vertical layout uniformly */
.card {
  background:#fff; border:1px solid #ddd; border-radius:10px; box-shadow:0 1px 3px rgba(0,0,0,0.06);
  overflow:hidden;
  min-width: 320px;
  min-height: 360px;
  flex: 0 0 auto;
  display: flex;
  flex-direction: column;
}

.card-head {
  padding:10px 12px; border-bottom:1px solid #eee;
}
.card-name { font-size:16px; font-weight:700; color:#2b3152; }
.card-desc { font-size:13px; color:#333; margin-top:2px; }
.card-type { font-size:12px; color:#666; margin-top:4px; }

/* Let card-body fill the remaining space */
.card-body {
  padding:10px 12px;
  display:flex;
  flex-direction:column;
  gap:8px;
  flex: 1 1 auto;
  min-height: 0;
}

/* Keep these sections at natural height */
.metrics {
  display:grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap:8px;
  flex: 0 0 auto;
}
.metric { font-size:13px; height:15px; background:#fafafa; border:1px solid #eee; border-radius:8px; padding:6px 8px; }
.metric.k    { background-color: rgba(0, 0, 0, 0.05); }
.metric.kw   { background-color: rgba(0, 0, 0, 0.03); }
.metric.c    { background-color: rgba(0, 174, 239, 0.10); }
.metric.cw   { background-color: rgba(0, 174, 239, 0.06); }
.metric.m    { background-color: rgba(236, 0, 140, 0.10); }
.metric.mw   { background-color: rgba(236, 0, 140, 0.06); }
.metric.y    { background-color: rgba(255, 242, 0, 0.20); }
.metric.yw   { background-color: rgba(255, 242, 0, 0.10); }
.metric.waste{ background-color: rgba(153, 102, 51, 0.12); }
.metric.drum { background-color: rgba(70, 130, 180, 0.12); }
.metric.fuser{ background-color: rgba(255, 99, 71, 0.12); }
.metric.maint{ background-color: rgba(255, 165, 0, 0.12); }

/* Keep pages short and on one line of “pills” */
.row.pages {
  display:flex; gap:8px; flex-wrap:wrap;
  flex: 0 0 auto;
}
.pill {
  background:#f3f3f3; border:1px solid #e2e2e2; border-radius:999px; padding:3px 8px; font-size:12px;
}

/* Status: scroll inside, hide overflow, and keep cards consistent */
.row.status {
  flex: 1 1 auto;
  height: 100px;
  overflow: overlay;
  background:#f9fafb;
  border:1px solid #eee;
  border-radius:8px;
  width:300px;
  word-break: break-word;
}
</style>
<script>
(function(){
  // Returns an object { missing: boolean, value: number }
  // missing = true when the cell has no numeric content (or is null/blank/nbsp)
  function parseCellForSort(cell){
    if (!cell) return { missing: true, value: 0 };
    var t = (cell.textContent || cell.innerText || "").replace(/\u00A0/g, " ").trim();
    if (!t) return { missing: true, value: 0 };
    var m = t.match(/-?\d+(?:[.,]\d+)?/);
    if (!m) return { missing: true, value: 0 };
    var n = parseFloat(m[0].replace(",", "."));
    if (isNaN(n)) return { missing: true, value: 0 };
    return { missing: false, value: n };
  }

  function getExpectedColumns(tbl){
    return (tbl && tbl.tHead && tbl.tHead.rows[0]) ? tbl.tHead.rows[0].cells.length : 0;
  }

  function isSectionHeaderRow(row, expectedCols){
    if (!row || !row.cells) return false;
    if (row.cells.length === 1) {
      var c = row.cells[0];
      if ((c.colSpan && c.colSpan >= expectedCols) ||
          (typeof c.hasAttribute === "function" && c.hasAttribute("colspan"))) {
        return true;
      }
    }
    return false;
  }

  function sortTableGrouped(tableId, colIndex, direction, isPercent){
    var tbl = document.getElementById(tableId);
    if (!tbl || !tbl.tBodies || !tbl.tBodies.length) return;

    var expected = getExpectedColumns(tbl);
    if (!expected) return;

    var body = tbl.tBodies[0];
    var allRows = Array.prototype.slice.call(body.rows);

    // Partition into groups to keep section headers intact
    var groups = [];
    var current = { header: null, rows: [] };
    for (var i = 0; i < allRows.length; i++){
      var r = allRows[i];
      if (isSectionHeaderRow(r, expected)) {
        groups.push(current);
        current = { header: r, rows: [] };
      } else {
        current.rows.push(r);
      }
    }
    groups.push(current);

    var asc = (direction === "asc");

    for (var g = 0; g < groups.length; g++){
      var rows = groups[g].rows;
      if (rows && rows.length > 1){
        var sortable = [];
        var nonsortable = [];

        for (var k = 0; k < rows.length; k++){
          if (rows[k].cells && rows[k].cells.length === expected) {
            sortable.push(rows[k]);
          } else {
            // Non-standard row in the middle of a section; preserve their relative order at the end
            nonsortable.push(rows[k]);
          }
        }

        sortable.sort(function(a, b){
          var av = parseCellForSort(a.cells[colIndex]);
          var bv = parseCellForSort(b.cells[colIndex]);

          // Primary key: rows with a value come before rows without a value
          if (av.missing !== bv.missing) {
            return av.missing ? 1 : -1;  // push missing to bottom for both asc/desc
          }

          // Secondary key: numeric comparison according to direction
          if (asc) {
            return av.value - bv.value;
          } else {
            return bv.value - av.value;
          }
        });

        groups[g].rows = sortable.concat(nonsortable);
      }
    }

    var frag = document.createDocumentFragment();
    for (var j = 0; j < groups.length; j++){
      var grp = groups[j];
      if (grp.header) frag.appendChild(grp.header);
      for (var y = 0; y < grp.rows.length; y++){
        frag.appendChild(grp.rows[y]);
      }
    }
    body.innerHTML = "";
    body.appendChild(frag);
  }

  // Column index map unchanged
  var colMap = {
    imaging:{idx:3,pct:true},
    black:{idx:4,pct:true},  blackw:{idx:5,pct:true},
    cyan:{idx:6,pct:true},   cyanw:{idx:7,pct:true},
    magenta:{idx:8,pct:true},magentaw:{idx:9,pct:true},
    yellow:{idx:10,pct:true},yelloww:{idx:11,pct:true},
    waste:{idx:12,pct:true}, drum:{idx:13,pct:true}, fuser:{idx:14,pct:true}, maintenance:{idx:15,pct:true},
    pages:{idx:16,pct:false}, pages1:{idx:17,pct:false}, pages7:{idx:18,pct:false}
  };

  window._printerReportSort = function(){
    var sel = document.getElementById("sort-col");
    var dir = document.getElementById("sort-dir").value;
    if(!sel || !sel.value){ alert("Pick a column to sort."); return; }
    var meta = colMap[sel.value]; if(!meta) return;
    sortTableGrouped("printer-table", meta.idx, dir, meta.pct);
  };

  window.addEventListener("DOMContentLoaded", function(){
    var tbl = document.getElementById("printer-table"); if(!tbl) return;
    var headers = tbl.tHead ? tbl.tHead.rows[0].cells : [];
    var last = {col:-1, dir:"desc"};
    for (var i=0;i<headers.length;i++){
      (function(ix){
        var isNumeric = (ix>=3 && ix<=18);
        if(!isNumeric) return;
        headers[ix].style.cursor = "pointer";
        headers[ix].title = headers[ix].getAttribute("title") || "Click to sort";
        headers[ix].addEventListener("click", function(){
          var isPct = (ix>=3 && ix<=15);
          var dir = (last.col===ix && last.dir==="asc") ? "desc" : "asc";
          sortTableGrouped("printer-table", ix, dir, isPct);
          last = {col:ix, dir:dir};
        });
      })(i);
    }
  });

  window.setView = function(v){
    var tableWrap = document.getElementById("table-wrap");
    var cards = document.getElementById("cards-view");
    var btnTable = document.getElementById("btn-view-table");
    var btnCards = document.getElementById("btn-view-cards");
    var sortControls = document.getElementById("sort-controls");

    if (v === "cards") {
      tableWrap.classList.add("hidden");
      cards.classList.remove("hidden");
      btnCards.classList.add("active");
      btnTable.classList.remove("active");
      if (sortControls) sortControls.style.display = "none";
    } else {
      cards.classList.add("hidden");
      tableWrap.classList.remove("hidden");
      btnTable.classList.add("active");
      btnCards.classList.remove("active");
      if (sortControls) sortControls.style.display = "";
    }
  };

  window.addEventListener("DOMContentLoaded", function(){
    var btnCards = document.getElementById("btn-view-cards");
    var sortControls = document.getElementById("sort-controls");
    var cardsVisible = !document.getElementById("cards-view").classList.contains("hidden");
    if (cardsVisible && sortControls) {
      sortControls.style.display = "none";
      btnCards && btnCards.classList.add("active");
    }
  });
})();
</script>

</head>
<body>
"@ | Out-File -Encoding UTF8 -FilePath $ReportTmp

# Heading + toolbar with view toggle
$validCount = ($printerlist.Value | Where-Object {$_ -and ($_ -notlike "-*")}).Count
"Reporting on $validCount printers" | Add-Content $ReportTmp

@"
<div class="toolbar">
  <div class="view-switch">
    <span>View:</span>
    <button id="btn-view-table" class="btn active" type="button" onclick="setView('table')">Table</button>
    <button id="btn-view-cards" class="btn" type="button" onclick="setView('cards')">Cards</button>
  </div>

  <!-- Sort controls -->
  <div id="sort-controls">
    <label for="sort-col">Sort by:</label>
    <select id="sort-col" aria-label="Sort column">
      <option value="" selected disabled>Choose column...</option>
      <option value="imaging">Imaging Kit (%)</option>
      <option value="black">Black (%)</option>
      <option value="blackw">Black (%/week)</option>
      <option value="cyan">Cyan (%)</option>
      <option value="cyanw">Cyan (%/week)</option>
      <option value="magenta">Magenta (%)</option>
      <option value="magentaw">Magenta (%/week)</option>
      <option value="yellow">Yellow (%)</option>
      <option value="yelloww">Yellow (%/week)</option>
      <option value="waste">Waste (% full)</option>
      <option value="drum">Drum (%)</option>
      <option value="fuser">Fuser (%)</option>
      <option value="maintenance">Maintenance (%)</option>
      <option value="pages">Pages (count)</option>
      <option value="pages1">Pages (1d)</option>
      <option value="pages7">Pages (7d)</option>
    </select>

    <label for="sort-dir">Direction:</label>
    <select id="sort-dir" aria-label="Sort direction">
      <option value="asc" selected>Low -> High</option>
      <option value="desc">High -> Low</option>
    </select>

    <button type="button" onclick="_printerReportSort()">Sort</button>
  </div>
</div>
"@ | Add-Content $ReportTmp

# Table wrapper
"<div id='table-wrap'>" | Add-Content $ReportTmp
"<table id='printer-table' style='width:100%'>" | Add-Content $ReportTmp

@"
<thead>
  <tr>
    <th title="Device description from your list (e.g., location/notes)">Description</th>
    <th title="Clickable name; opens printer web UI in a new tab">Name</th>
    <th title="Model/type as reported by SNMP (hrDeviceDescr)">Type</th>
    <th title="Imaging kit / imaging unit life (% remaining)">Imaging Kit</th>
    <th title="Black toner (% remaining)">Black</th>
    <th title="Estimated black usage this week (%/week)">Black (%/week)</th>
    <th title="Cyan toner (% remaining)">Cyan</th>
    <th title="Estimated cyan usage this week (%/week)">Cyan (%/week)</th>
    <th title="Magenta toner (% remaining)">Magenta</th>
    <th title="Estimated magenta usage this week (%/week)">Magenta (%/week)</th>
    <th title="Yellow toner (% remaining)">Yellow</th>
    <th title="Estimated yellow usage this week (%/week)">Yellow (%/week)</th>
    <th title="Waste container (% full) - higher is worse">Waste</th>
    <th title="Drum/photoconductor life (% remaining)">Drum</th>
    <th title="Fuser life (% remaining)">Fuser</th>
    <th title="Maintenance items (transfer/cleaner/etc.) (% remaining)">Maintenance</th>
    <th title="Total lifetime pages (best available counter)">Pages</th>
    <th title="Pages printed in last 1 day (based on history)">Pages (1d)</th>
    <th title="Pages printed in last 7 days (based on history)">Pages (7d)</th>
    <th title="Alerts/status from prtAlertDescription">Status</th>
  </tr>
</thead>
<tbody>
"@ | Add-Content $ReportTmp

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
$weekAgo = $now.AddDays(-7)
$dayAgo  = $now.AddDays(-1)

foreach ($p in $printerlist) {

    # Section headers for lines starting with '-'
    if ($p.Value -like "-*") {
        $sectionTitle = HtmlEncode($p.Value.TrimStart('-'))
        "<tr><td colspan='20'><h3>$sectionTitle</h3></td></tr>" | Add-Content $ReportTmp

        # CARDS: close previous grid if open, add section divider, and open a new grid
        if ($cardsGridOpen) {
            [void]$CardBuffer.AppendLine("</div>")
            $cardsGridOpen = $false
        }
        [void]$CardBuffer.AppendLine("<h3 class='section-divider'>$sectionTitle</h3>")
        [void]$CardBuffer.AppendLine("<div class='cards-grid'>")
        $cardsGridOpen = $true

        continue
    }

    $index++
    $ip   = ($p.Value).Trim()
    $name = $p.Name
    $desc = $p.Description

    "<tr>" | Add-Content $ReportTmp
    "<td title='Description'><b>$(HtmlEncode($desc))</b></td>" | Add-Content $ReportTmp

    if (-not (Test-HostUp $ip)) {
        # ============================
        # OFFLINE: fallback to history
        # ============================
        $href = "<a href=""http://$ip"" target=""_new"">$ip</a>"
        $cache = Get-HistorySnapshot -IP $ip
        $cachedStamp = $null
        if ($cache -and $cache.TS) {
            $cachedStamp = $cache.TS.ToString("g")
        }

        "<td title='Device web interface'>$href</td>" | Add-Content $ReportTmp
        if ($cachedStamp) {
            "<td title='Model/Type'><b>Offline</b><br/><span style='color:#666;font-size:12px;'>Cached: $cachedStamp</span></td>" | Add-Content $ReportTmp
        } else {
            "<td title='Model/Type'><b>Offline</b></td>" | Add-Content $ReportTmp
        }

        "<td title='Imaging Kit'></td>" | Add-Content $ReportTmp
        "<td title='Black Toner'>$(Format-PercentCell $cache.K)</td>"                 | Add-Content $ReportTmp
        "<td title='Black Toner (%/week)'>$(Format-PercentPerWeek $cache.KWeek)</td>" | Add-Content $ReportTmp
        "<td title='Cyan Toner'>$(Format-PercentCell $cache.C)</td>"                  | Add-Content $ReportTmp
        "<td title='Cyan Toner (%/week)'>$(Format-PercentPerWeek $cache.CWeek)</td>"  | Add-Content $ReportTmp
        "<td title='Magenta Toner'>$(Format-PercentCell $cache.M)</td>"               | Add-Content $ReportTmp
        "<td title='Magenta Toner (%/week)'>$(Format-PercentPerWeek $cache.MWeek)</td>" | Add-Content $ReportTmp
        "<td title='Yellow Toner'>$(Format-PercentCell $cache.Y)</td>"                | Add-Content $ReportTmp
        "<td title='Yellow Toner (%/week)'>$(Format-PercentPerWeek $cache.YWeek)</td>" | Add-Content $ReportTmp

        "<td title='Waste (% full)'></td><td title='Drum'></td><td title='Fuser'></td><td title='Maintenance'></td>" | Add-Content $ReportTmp

        "<td title='Pages (lifetime)'>$(Format-Count $cache.Pages)</td>" | Add-Content $ReportTmp
        "<td title='Pages (1d)'>$(Format-Count $cache.Pages1)</td>"      | Add-Content $ReportTmp
        "<td title='Pages (7d)'>$(Format-Count $cache.Pages7)</td>"      | Add-Content $ReportTmp

        if ($cachedStamp) {
            "<td title='Status / Alerts'><b>Offline</b><br/><span style='color:#666;font-size:12px;'>Showing cached values from $cachedStamp</span></td>" | Add-Content $ReportTmp
        } else {
            "<td title='Status / Alerts'><b>Offline</b></td>" | Add-Content $ReportTmp
        }
        "</tr>" | Add-Content $ReportTmp

        # CARD: ensure a grid exists
        if (-not $cardsGridOpen) {
            [void]$CardBuffer.AppendLine("<div class='cards-grid'>")
            $cardsGridOpen = $true
        }

        $nameOrIp = ($name,$ip | Where-Object {$_})[0]
        $cachedNoteInline = ""
        if ($cachedStamp) {
            $cachedNoteInline = " <span style='color:#666;font-size:12px;'>(cached $cachedStamp)</span>"
        }

        $cardOffline = @"
  <div class='card' data-ip='$ip' title='Printer offline'>
    <div class='card-head'>
      <div class='card-name'><a href="http://$ip" target="_new">$(HtmlEncode($nameOrIp))</a></div>
      <div class='card-desc'>$(HtmlEncode($desc))</div>
      <div class='card-type'><b>Offline</b>$cachedNoteInline</div>
    </div>
    <div class='card-body'>
      <div class='metrics'>
        <div class='metric'    title='Imaging Kit'>Imaging: &nbsp;</div>
        <div class='metric k'  title='Black Toner'>Black: $(Format-PercentCell $cache.K)</div>
        <div class='metric kw' title='Black Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.KWeek)</div>
        <div class='metric c'  title='Cyan Toner'>Cyan: $(Format-PercentCell $cache.C)</div>
        <div class='metric cw' title='Cyan Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.CWeek)</div>
        <div class='metric m'  title='Magenta Toner'>Magenta: $(Format-PercentCell $cache.M)</div>
        <div class='metric mw' title='Magenta Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.MWeek)</div>
        <div class='metric y'  title='Yellow Toner'>Yellow: $(Format-PercentCell $cache.Y)</div>
        <div class='metric yw' title='Yellow Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.YWeek)</div>
        <div class='metric waste' title='Waste (% full)'>Waste: &nbsp;</div>
        <div class='metric drum'  title='Drum'>Drum: &nbsp;</div>
        <div class='metric fuser' title='Fuser'>Fuser: &nbsp;</div>
        <div class='metric maint' title='Maintenance'>Maint: &nbsp;</div>
      </div>
      <div class='row pages'>
        <span class='pill' title='Pages (lifetime)'>Pages: $(Format-Count $cache.Pages)</span>
        <span class='pill' title='Pages (1d)'>1d: $(Format-Count $cache.Pages1)</span>
        <span class='pill' title='Pages (7d)'>7d: $(Format-Count $cache.Pages7)</span>
      </div>
      <div class='row status' title='Status / Alerts'><b>Offline</b></div>
    </div>
  </div>
"@
        [void]$CardBuffer.AppendLine($cardOffline)
        continue
    }

    $printerType = $null
    $statusText  = "Operational"
    $K=$null; $C=$null; $M=$null; $Y=$null
    $Waste=$null; $Drum=$null; $Fuser=$null; $Maintenance=$null
    [Nullable[Int64]]$Pages = $null
    [Nullable[Int64]]$Pages7 = $null
    [Nullable[Int64]]$Pages1 = $null
    [Nullable[int]]$ImagingKit = $null

    [Nullable[int]]$KWeek=$null; [Nullable[int]]$CWeek=$null; [Nullable[int]]$MWeek=$null; [Nullable[int]]$YWeek=$null

    try {
        $snmp.Open($ip, $Community, $Retries, $TimeoutMs)

        # Identify printer
        try { $printerType = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1") } catch {}
        $sysName = $null; try { $sysName = $snmp.Get(".1.3.6.1.2.1.1.5.0") } catch {}

        # Supplies
        $rows = Get-PrinterSupplies -Snmp $snmp

        # === Imaging Kits (per color when available, else fallback) ===
$imaging = Get-ImagingByColor -rows $rows

# Show the "worst (lowest %)" imaging among available K/C/M/Y if any; else mono (Any)
[Nullable[int]]$ImagingKit = $null
$colorImagingValues = @()
if ($imaging.K -ne $null) { $colorImagingValues += $imaging.K }
if ($imaging.C -ne $null) { $colorImagingValues += $imaging.C }
if ($imaging.M -ne $null) { $colorImagingValues += $imaging.M }
if ($imaging.Y -ne $null) { $colorImagingValues += $imaging.Y }

if ($colorImagingValues.Count -gt 0) {
    $ImagingKit = ($colorImagingValues | Measure-Object -Minimum).Minimum
} elseif ($imaging.Any -ne $null) {
    $ImagingKit = $imaging.Any
}

        # Toners: (Class=3 AND (Type=3 OR description suggests toner/cartridge))
$toners = $rows | Where-Object {
    ($_.Class -eq 3 -and $_.Type -eq 3) -or
    ($_.Class -eq 3 -and $_.Desc -match '(?i)\b(toner|cartridge|ctg|print\s*cartridge)\b')
}

foreach ($t in $toners) {
    switch ($t.ColorLetter) {
        'K' { if ($null -eq $K) { $K = $t.Percent } }
        'C' { if ($null -eq $C) { $C = $t.Percent } }
        'M' { if ($null -eq $M) { $M = $t.Percent } }
        'Y' { if ($null -eq $Y) { $Y = $t.Percent } }
    }
}
        if ($null -eq $K) {
            $singleToner = $toners | Where-Object { $_.Percent -ne $null } | Select-Object -First 2
            if ($singleToner.Count -eq 1) { $K = $singleToner[0].Percent }
        }
        if ($null -eq $K) {
            $colorlessK = $toners | Where-Object { $_.Percent -ne $null -and (Is-ColorlessTonerDesc -desc $_.Desc) } | Select-Object -First 1
            if ($colorlessK) { $K = $colorlessK.Percent }
        }
        if ($null -eq $K -and $C -eq $null -and $M -eq $null -and $Y -eq $null) {
            $best = $toners | Where-Object { $_.Percent -ne $null } | Sort-Object Percent -Descending | Select-Object -First 1
            if ($best) { $K = $best.Percent }
        }

        # Waste
        $wasteTypes = @(4,8,14)
        $wasteRow = $rows | Where-Object {
    	($_.Type -in $wasteTypes) -or
    	($_.Class -eq 4 -and $_.Desc -match '(?i)\b(waste|waste\s*(toner|box|bottle|container))\b')
	} | Select-Object -First 1

        # Drum
        $drumRow = $rows | Where-Object { ($_.Type -eq 9) -or ($_.Desc -match '(?i)(drum|opc|photoconductor)') } | Select-Object -First 1
        if ($drumRow) { $Drum = $drumRow.Percent }

        # Fuser
        $fuserRow = $rows | Where-Object { ($_.Type -eq 15) -or ($_.Desc -match '(?i)fuser') } | Select-Object -First 1
        if ($fuserRow) { $Fuser = $fuserRow.Percent }

        # Maintenance kit
        $maintTypes = @(18,19,20,22)
        $maintRow = $rows | Where-Object { ($_.Type -in $maintTypes) -or ($_.Desc -match '(?i)\bmaintenance\b|\bkit\b') } | Select-Object -First 1
        if ($maintRow) { $Maintenance = $maintRow.Percent }

        # Lifetime pages
        $Pages = Get-PrinterPageCount -Snmp $snmp

        # Compute 7-day and 1-day page deltas + %/week per color
        if ($Pages -ne $null) {
            $histRows = $history | Where-Object { $_.IP -eq $ip } |
                        Select-Object @{n='TS';e={[datetime]::Parse($_.Timestamp)}},
                                      @{n='Pages';e={ if($_.Pages -match '^\d+$'){ [int64]$_.Pages } else { $null } }},
                                      @{n='K';e={ if($_.K -match '^\d+$'){ [int]$_.K } else { $null } }},
                                      @{n='C';e={ if($_.C -match '^\d+$'){ [int]$_.C } else { $null } }},
                                      @{n='M';e={ if($_.M -match '^\d+$'){ [int]$_.M } else { $null } }},
                                      @{n='Y';e={ if($_.Y -match '^\d+$'){ [int]$_.Y } else { $null } }} |
                        Sort-Object TS

            $baseline7 = $histRows | Where-Object { $_.TS -le $weekAgo } | Sort-Object TS -Descending | Select-Object -First 1
            if ($baseline7 -and $Pages -ge $baseline7.Pages) { $Pages7 = $Pages - $baseline7.Pages } else { $Pages7 = $null }

            $baseline1 = $histRows | Where-Object { $_.TS -le $dayAgo } | Sort-Object TS -Descending | Select-Object -First 1
            if ($baseline1 -and $Pages -ge $baseline1.Pages) { $Pages1 = $Pages - $baseline1.Pages } else { $Pages1 = $null }

            function Get-WeeklyUsage {
                param(
                    [Nullable[int]]$CurrentPct,
                    [string]$ColorKey
                )
                if ($null -eq $CurrentPct) { return $null }
                if ($null -eq $Pages7 -or $Pages7 -le 0) {
                    if ($baseline7 -and ($baseline7.$ColorKey -ne $null)) {
                        $drop = $baseline7.$ColorKey - $CurrentPct
                        if ($drop -lt 0) { $drop = 0 }
                        return [int][math]::Min(100, [math]::Round($drop))
                    }
                    return $null
                }
                if ($baseline7 -and ($baseline7.Pages -ne $null) -and ($baseline7.$ColorKey -ne $null) -and $Pages -ge $baseline7.Pages) {
                    $deltaPages = $Pages - $baseline7.Pages
                    $deltaPct   = $baseline7.$ColorKey - $CurrentPct
                    if ($deltaPct -lt 0) { $deltaPct = 0 }
                    if ($deltaPages -gt 0) {
                        $pctPerPage = $deltaPct / $deltaPages
                        $weekly = $pctPerPage * $Pages7
                        return [int][math]::Min(100, [math]::Round($weekly))
                    }
                }
                if ($baseline7 -and ($baseline7.$ColorKey -ne $null)) {
                    $drop = $baseline7.$ColorKey - $CurrentPct
                    if ($drop -lt 0) { $drop = 0 }
                    return [int][math]::Min(100, [math]::Round($drop))
                }
                return $null
            }

            $KWeek = Get-WeeklyUsage -CurrentPct $K -ColorKey 'K'
            $CWeek = Get-WeeklyUsage -CurrentPct $C -ColorKey 'C'
            $MWeek = Get-WeeklyUsage -CurrentPct $M -ColorKey 'M'
            $YWeek = Get-WeeklyUsage -CurrentPct $Y -ColorKey 'Y'
        }

        # Alerts -> Status
        try {
            $alerts = @($snmp.GetTree(".1.3.6.1.2.1.43.18.1.1.8"))
            $filtered = $alerts | Where-Object { $_ -and $_ -notlike "print*" -and $_ -notlike "*bypass*" }
            if ($filtered.Count -gt 0) { $statusText = ($filtered -join "; ") }
        } catch {}

        # TABLE cells: Name/Type
        $displayName = if ($sysName) { $sysName } else { $name }
        $href = "<a href=""http://$ip"" target=""_new"">$(HtmlEncode($displayName))</a>"
        "<td title='Device web interface'>$href</td>" | Add-Content $ReportTmp
        "<td title='Model/Type'><br/>$(HtmlEncode($printerType))<br/></td>" | Add-Content $ReportTmp

        # TABLE cells: Imaging + Toners + Supplies + Pages + Status
        "<td title='Imaging Kit'>$(Format-PercentCell $ImagingKit)</td>" | Add-Content $ReportTmp
        "<td title='Black Toner'>$(Format-PercentCell $K)</td>" | Add-Content $ReportTmp
        "<td title='Black Toner (%/week)'>$(Format-PercentPerWeek $KWeek)</td>" | Add-Content $ReportTmp
        "<td title='Cyan Toner'>$(Format-PercentCell $C)</td>" | Add-Content $ReportTmp
        "<td title='Cyan Toner (%/week)'>$(Format-PercentPerWeek $CWeek)</td>" | Add-Content $ReportTmp
        "<td title='Magenta Toner'>$(Format-PercentCell $M)</td>" | Add-Content $ReportTmp
        "<td title='Magenta Toner (%/week)'>$(Format-PercentPerWeek $MWeek)</td>" | Add-Content $ReportTmp
        "<td title='Yellow Toner'>$(Format-PercentCell $Y)</td>" | Add-Content $ReportTmp
        "<td title='Yellow Toner (%/week)'>$(Format-PercentPerWeek $YWeek)</td>" | Add-Content $ReportTmp
        "<td title='Waste (% full)'>$(Format-PercentCellWaste $Waste)</td>" | Add-Content $ReportTmp
        "<td title='Drum'>$(Format-PercentCell $Drum)</td>" | Add-Content $ReportTmp
        "<td title='Fuser'>$(Format-PercentCell $Fuser)</td>" | Add-Content $ReportTmp
        "<td title='Maintenance'>$(Format-PercentCell $Maintenance)</td>" | Add-Content $ReportTmp
        "<td title='Pages (lifetime)'>$(Format-Count $Pages)</td>" | Add-Content $ReportTmp
        "<td title='Pages (1d)'>$(Format-Count $Pages1)</td>" | Add-Content $ReportTmp
        "<td title='Pages (7d)'>$(Format-Count $Pages7)</td>" | Add-Content $ReportTmp
        "<td title='Status / Alerts'><b>$(HtmlEncode($statusText))</b></td>" | Add-Content $ReportTmp

        # CARD view
        if (-not $cardsGridOpen) {
            [void]$CardBuffer.AppendLine("<div class='cards-grid'>")
            $cardsGridOpen = $true
        }
        $card = @"
  <div class='card' data-ip='$ip' title='Printer'>
    <div class='card-head'>
      <div class='card-name'><a href="http://$ip" target="_new">$(HtmlEncode($displayName))</a></div>
      <div class='card-desc'>$(HtmlEncode($desc))</div>
      <div class='card-type'>$(HtmlEncode($printerType))</div>
    </div>
    <div class='card-body'>
      <div class='metrics'>
        <div class='metric'   title='Imaging Kit'>Imaging: $(Format-PercentCell $ImagingKit)</div>
        <div class='metric k' title='Black Toner'>Black: $(Format-PercentCell $K)</div>
        <div class='metric kw' title='Black Toner (%/week)'>/wk: $(Format-PercentPerWeek $KWeek)</div>
        <div class='metric c' title='Cyan Toner'>Cyan: $(Format-PercentCell $C)</div>
        <div class='metric cw' title='Cyan Toner (%/week)'>/wk: $(Format-PercentPerWeek $CWeek)</div>
        <div class='metric m' title='Magenta Toner'>Magenta: $(Format-PercentCell $M)</div>
        <div class='metric mw' title='Magenta Toner (%/week)'>/wk: $(Format-PercentPerWeek $MWeek)</div>
        <div class='metric y' title='Yellow Toner'>Yellow: $(Format-PercentCell $Y)</div>
        <div class='metric yw' title='Yellow Toner (%/week)'>/wk: $(Format-PercentPerWeek $YWeek)</div>
        <div class='metric waste' title='Waste (% full)'>Waste: $(Format-PercentCellWaste $Waste)</div>
        <div class='metric drum'  title='Drum'>Drum: $(Format-PercentCell $Drum)</div>
        <div class='metric fuser' title='Fuser'>Fuser: $(Format-PercentCell $Fuser)</div>
        <div class='metric maint' title='Maintenance'>Maint: $(Format-PercentCell $Maintenance)</div>
      </div>
      <div class='row pages'>
        <span class='pill' title='Pages (lifetime)'>Pages: $(Format-Count $Pages)</span>
        <span class='pill' title='Pages (1d)'>1d: $(Format-Count $Pages1)</span>
        <span class='pill' title='Pages (7d)'>7d: $(Format-Count $Pages7)</span>
      </div>
      <div class='row status' title='Status / Alerts'><b>$(HtmlEncode($statusText))</b></div>
    </div>
  </div>
"@
        [void]$CardBuffer.AppendLine($card)

        # Append to history
        if ($Pages -ne $null) {
            $kOut = if ($K -ne $null) { $K } else { "" }
            $cOut = if ($C -ne $null) { $C } else { "" }
            $mOut = if ($M -ne $null) { $M } else { "" }
            $yOut = if ($Y -ne $null) { $Y } else { "" }
            "$($now.ToString('o')),$ip,$Pages,$kOut,$cOut,$mOut,$yOut" | Add-Content -Path $HistoryCsv -Encoding UTF8
        }

    } catch {
        # ============================
        # SNMP failure: use cached data
        # ============================
        $href = "<a href=""http://$ip"" target=""_new"">$(HtmlEncode(($name,$ip | Where-Object {$_})[0]))</a>"
        $cache = Get-HistorySnapshot -IP $ip
        $cachedStamp = $null
        if ($cache -and $cache.TS) {
            $cachedStamp = $cache.TS.ToString("g")
        }

        "<td title='Device web interface'>$href</td>" | Add-Content $ReportTmp
        if ($cachedStamp) {
            "<td title='Model/Type'><b>No live data</b><br/><span style='color:#666;font-size:12px;'>Cached: $cachedStamp</span></td>" | Add-Content $ReportTmp
        } else {
            "<td title='Model/Type'><b>No data</b></td>" | Add-Content $ReportTmp
        }

        "<td title='Imaging Kit'></td>" | Add-Content $ReportTmp
        "<td title='Black Toner'>$(Format-PercentCell $cache.K)</td>"                  | Add-Content $ReportTmp
        "<td title='Black Toner (%/week)'>$(Format-PercentPerWeek $cache.KWeek)</td>"  | Add-Content $ReportTmp
        "<td title='Cyan Toner'>$(Format-PercentCell $cache.C)</td>"                   | Add-Content $ReportTmp
        "<td title='Cyan Toner (%/week)'>$(Format-PercentPerWeek $cache.CWeek)</td>"   | Add-Content $ReportTmp
        "<td title='Magenta Toner'>$(Format-PercentCell $cache.M)</td>"                | Add-Content $ReportTmp
        "<td title='Magenta Toner (%/week)'>$(Format-PercentPerWeek $cache.MWeek)</td>"| Add-Content $ReportTmp
        "<td title='Yellow Toner'>$(Format-PercentCell $cache.Y)</td>"                 | Add-Content $ReportTmp
        "<td title='Yellow Toner (%/week)'>$(Format-PercentPerWeek $cache.YWeek)</td>" | Add-Content $ReportTmp

        "<td title='Waste (% full)'></td><td title='Drum'></td><td title='Fuser'></td><td title='Maintenance'></td>" | Add-Content $ReportTmp

        "<td title='Pages (lifetime)'>$(Format-Count $cache.Pages)</td>" | Add-Content $ReportTmp
        "<td title='Pages (1d)'>$(Format-Count $cache.Pages1)</td>"      | Add-Content $ReportTmp
        "<td title='Pages (7d)'>$(Format-Count $cache.Pages7)</td>"      | Add-Content $ReportTmp

        if ($cachedStamp) {
            "<td title='Status / Alerts'><b>No live data</b><br/><span style='color:#666;font-size:12px;'>Showing cached values from $cachedStamp</span></td>" | Add-Content $ReportTmp
        } else {
            "<td title='Status / Alerts'><b>No data</b></td>" | Add-Content $ReportTmp
        }

        # CARD: ensure a grid exists
        if (-not $cardsGridOpen) {
            [void]$CardBuffer.AppendLine("<div class='cards-grid'>")
            $cardsGridOpen = $true
        }

        $nameOrIp = ($name,$ip | Where-Object {$_})[0]
        $state = "No data"
        if ($cachedStamp) { $state = "No live data (cached $cachedStamp)" }

        $cardNoData = @"
  <div class='card' data-ip='$ip' title='$state'>
    <div class='card-head'>
      <div class='card-name'><a href="http://$ip" target="_new">$(HtmlEncode($nameOrIp))</a></div>
      <div class='card-desc'>$(HtmlEncode($desc))</div>
      <div class='card-type'><b>$state</b></div>
    </div>
    <div class='card-body'>
      <div class='metrics'>
        <div class='metric'    title='Imaging Kit'>Imaging: &nbsp;</div>
        <div class='metric k'  title='Black Toner'>Black: $(Format-PercentCell $cache.K)</div>
        <div class='metric kw' title='Black Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.KWeek)</div>
        <div class='metric c'  title='Cyan Toner'>Cyan: $(Format-PercentCell $cache.C)</div>
        <div class='metric cw' title='Cyan Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.CWeek)</div>
        <div class='metric m'  title='Magenta Toner'>Magenta: $(Format-PercentCell $cache.M)</div>
        <div class='metric mw' title='Magenta Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.MWeek)</div>
        <div class='metric y'  title='Yellow Toner'>Yellow: $(Format-PercentCell $cache.Y)</div>
        <div class='metric yw' title='Yellow Toner (%/week)'>/wk: $(Format-PercentPerWeek $cache.YWeek)</div>
        <div class='metric waste' title='Waste (% full)'>Waste: &nbsp;</div>
        <div class='metric drum'  title='Drum'>Drum: &nbsp;</div>
        <div class='metric fuser' title='Fuser'>Fuser: &nbsp;</div>
        <div class='metric maint' title='Maintenance'>Maint: &nbsp;</div>
      </div>
      <div class='row pages'>
        <span class='pill' title='Pages (lifetime)'>Pages: $(Format-Count $cache.Pages)</span>
        <span class='pill' title='Pages (1d)'>1d: $(Format-Count $cache.Pages1)</span>
        <span class='pill' title='Pages (7d)'>7d: $(Format-Count $cache.Pages7)</span>
      </div>
      <div class='row status' title='Status / Alerts'><b>$state</b></div>
    </div>
  </div>
"@
        [void]$CardBuffer.AppendLine($cardNoData)
    } finally {
        try { $snmp.Close() } catch {}
    }

    "</tr>" | Add-Content $ReportTmp
}

"</tbody></table>" | Add-Content $ReportTmp
"</div>" | Add-Content $ReportTmp  # close #table-wrap

# Close any open cards grid and output Cards view
if ($cardsGridOpen) {
    [void]$CardBuffer.AppendLine("</div>")
    $cardsGridOpen = $false
}

@"
<div id="cards-view" class="hidden">
  $($CardBuffer.ToString())
</div>
"@ | Add-Content $ReportTmp

$stamp = Get-Date -Format "dd/MM HH:mm"
"<h3>$stamp</h3></body></html>" | Add-Content $ReportTmp

# Publish
Move-Item -Path $ReportTmp -Destination $ReportHtml -Force

if ($OpenWhenDone) { Start-Process $ReportHtml }
Write-Host "Report written to: $ReportHtml"
if ($WriteDebugCsv) { Write-Host "Debug rows written to: $DebugCsv" }
