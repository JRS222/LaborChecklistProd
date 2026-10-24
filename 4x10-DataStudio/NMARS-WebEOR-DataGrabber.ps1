#Requires -Version 5.1
<#
.SYNOPSIS
    NMARS + WebEOR Data Grabber — downloads all the reports for the 4/10
    Modified Work Week proposal in one run.

.DESCRIPTION
    Built from the saved NMARS report forms (eCBM completion, tasks bypassed,
    PM trends, contract cost, equipment inventory, parts issues, national stock
    room, acronym listing) and the WebEOR EOR/machine-mapping flow from
    WEBEOR.ps1.  Runs on an on-intranet USPS workstation under Windows
    PowerShell 5.1 (NOT pwsh 7).

    GUI mode (default): pick site, periods, date range, and which reports to
    pull, then Run.  CLI mode: -NoGui.  -DryRun prints every URL/request the
    script WOULD make without touching the network (safe to test anywhere).

.PARAMETER NoGui
    Run headless instead of showing the WinForms window.

.PARAMETER SiteId
    Numeric eMARS/WebEOR site id (e.g., 19101).  Required in CLI mode unless
    only site-independent reports are selected.

.PARAMETER OutputDir
    Root folder.  Each run creates a stamped subfolder (NMARS-Reports_yyyyMMdd_HHmmss).

.PARAMETER Reports
    Comma-separated report ids for CLI mode, or ALL.  Ids:
    PM4WKS_SITE  PM4WKS_NAT  PM4WKS_AREA  PM4WKS_DIST  PMAREA_SITE  PMAREA_AREA
    BYPASSED     TRENDS_SITE TRENDS_NAT   TRENDS_AREA  CONTRACT_SITE
    CONTRACT_NAT CONTRACT_AREA  EQUIP_SITE  EQUIP_NAT  ISSUES_SITE  ISSUES_NAT
    STOCK        ACRONYMS    EOR_RUNS  EOR_MAINT  RPG_RUNVSPLAN  MAPPING
    (Legacy aliases: EOR = EOR_RUNS+EOR_MAINT; EOR_PERF = EOR_RUNS)

.PARAMETER FromDate / ToDate
    Date range for WebEOR EOR and Contract Cost (mm/dd/yyyy).  Defaults to the
    last 12 months.

.PARAMETER Fy / FromPeriod / ToPeriod
    Fiscal year (default current FY) and accounting periods (1-13, default
    1..current) for the FY-based NMARS reports.

.PARAMETER AreaCode / DistrictCode
    Region (1V Eastern / 1W Western) and Division (D01..D17) for area/division
    benchmark reports.

.PARAMETER DryRun
    Print the plan (URLs + WebEOR request descriptions) and exit — no network.

.PARAMETER ListSites
    Fetch and print the NMARS site list (requires intranet access), then exit.

.PARAMETER TestConnection
    Run the connection self-test only: checks NMARS reachability and WebEOR login,
    writes samples of any failed response, then exits.  Use this FIRST from the
    intranet to prove connectivity before the full run.

.PARAMETER OpenStudio
    After a successful run, open the Data Studio and Coverage Studio pages with the
    data folder ready to drop.

.PARAMETER Bundle
    Zip the run folder (log, manifest, samples, diagnostics) into a single
    NMARS-WebEOR_<stamp>_bundle.zip for reporting back.

.EXAMPLE
    powershell -File NMARS-WebEOR-DataGrabber.ps1 -NoGui -SiteId 19010 -TestConnection

    From the intranet, verifies NMARS + WebEOR before the big run.

.EXAMPLE
    powershell -File NMARS-WebEOR-DataGrabber.ps1 -NoGui -SiteId 19010 -Reports ALL -Quarter 4 -Bundle -OpenStudio

    Full network run for one site: every NMARS report + WebEOR CSV, with a
    diagnostics bundle and the studio pages opened afterwards.

.EXAMPLE
    powershell -File NMARS-WebEOR-DataGrabber.ps1 -NoGui -SiteId 19101 -OutputDir C:\data -DryRun

.EXAMPLE
    powershell -File NMARS-WebEOR-DataGrabber.ps1 -NoGui -SiteId 19101 -Reports "PM4WKS_SITE,TRENDS_SITE,CONTRACT_SITE,EOR,MAPPING"
#>
[CmdletBinding()]
param(
    [switch]$NoGui,
    [string]$SiteId,
    [string]$OutputDir,
    [string]$Reports = "ALL",
    [string]$FromDate,
    [string]$ToDate,
    [int]$Fy,
    [int]$FromPeriod,
    [int]$ToPeriod,
    [string]$AreaCode = "1V",
    [string]$DistrictCode = "",
    [int]$Quarter,
    [switch]$DryRun,
    [switch]$ListSites,
    [switch]$TestConnection,
    [switch]$OpenStudio,
    [switch]$Bundle
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
$script:NMARS_BASE  = "http://emarssu3.eng.usps.gov/pemarsnp"   # eMARS National Reporting
$script:WEBEOR_BASE = "https://webeor.usps.gov"                 # WebEOR
$script:UA          = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
                      "AppleWebKit/537.36 (KHTML, like Gecko) " +
                      "Chrome/155.0.0.0 Safari/537.36 Edg/155.0.0.0"
$script:ChunkSize          = 14      # WebEOR EOR: days per CSV chunk
$script:ChunkRetries       = 2
$script:RetrySleepSec      = 5
$script:InterChunkSleepSec = 1
$script:InterReportSleepMs = 800
$script:NmarsRetries       = 2
$script:CancelRequested    = $false
$script:RunStamp           = Get-Date -Format "yyyyMMdd_HHmmss"
$script:LogFile            = $null
$script:LogBox             = $null
$script:ProgressBar        = $null

# --- WebEOR EOR variants (from WebEOR User Guide v1.1.62) --------------------
# Filter values seen on the EOR Viewer: "All Runs", "All Runs - Maint. Runs Included",
# "Maint-Runs - Maint. Runs Only", "All Runs wo/Maint Runs", "Split Days...", etc.
# NOTE: maintenance runs are kept SEPARATE from production runs — maint runs do not
# drive production staffing, but they show whether maintenance windows are being kept.
$script:EorFilterRuns  = "All Runs wo/Maint Runs"            # production runs only
$script:EorFilterMaint = "Maint-Runs - Maint. Runs Only"     # maintenance runs only
# END_OF_RUN table columns (guide Appendix E): DOWN_TIME, IDLE_TIME, JAM_TIME,
# MAINT_TIME, AP_NBR, FISCAL_YR, MODS_DATE.  If your viewer rejects one, trim it here.
$script:EorPerfFieldNames = "SITE_CODE,MACH_TYPE_DESC,MACHINE_NBR,OP_NBR,MSP,TOUR_CODE,RUN_NBR,START_DATE,END_DATE,TOTAL_PIECES_FED_CNT,MODSTRANS,DSIS,DOWN_TIME,IDLE_TIME,JAM_TIME,MAINT_TIME,AP_NBR,FISCAL_YR,MODS_DATE"
$script:EorPerfFieldLabels = "Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS,DownTime,IdleTime,JamTime,MaintTime,AP,FY,MODS Date"
$script:EorBaseFieldNames  = "SITE_CODE,MACH_TYPE_DESC,MACHINE_NBR,OP_NBR,MSP,TOUR_CODE,RUN_NBR,START_DATE,END_DATE,TOTAL_PIECES_FED_CNT,MODSTRANS,DSIS"
$script:EorBaseFieldLabels = "Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS"
$script:EorMaintFieldNames  = $script:EorBaseFieldNames  + ",AP_NBR,FISCAL_YR,MODS_DATE"
$script:EorMaintFieldLabels = $script:EorBaseFieldLabels + ",AP,FY,MODS Date"
# RPG "Run vs Plan Volume" viewer (guide ch.3): fed volume, run time, throughput by
# site / MODS date / machine type / operation / tour.  The page is discovered from the
# WebEOR main menu; this is the query string template ({site} is substituted).
$script:RpgRunVsPlanQuery = "SITEID={site}"
# Local fiscal-calendar adjustment (days) used ONLY when the NMARS date server
# is unreachable.  0 = AP1 starts the Saturday on/before October 1.
$script:ApStartOverrideDays = 0

$script:SettingsDir  = Join-Path $env:LOCALAPPDATA "NMARSWebEORDataGrabber"
$script:SettingsFile = Join-Path $script:SettingsDir "settings.json"

# --- diagnostics (network-capable logging: everything is recorded so issues
#     can be reported back and fixed) ------------------------------------------
$script:JsonlPath  = $null    # structured log (one JSON object per line)
$script:SampleDir  = $null    # folder for raw response samples of failed items
$script:DiagNotes  = New-Object System.Collections.Generic.List[string]

function Save-Sample {
    # Writes a truncated raw response snippet for later review.
    param([string]$Name, [string]$Content, [int]$MaxChars = 800)
    if (-not $script:SampleDir) { return $null }
    try {
        if (-not (Test-Path $script:SampleDir)) {
            New-Item -ItemType Directory -Force -Path $script:SampleDir | Out-Null
        }
        $safe = $Name -replace '[^\w\-\.]', '_'
        $p = Join-Path $script:SampleDir ("{0}_{1}.txt" -f (Get-Date -Format 'HHmmssfff'), $safe)
        if ($null -eq $Content) { $Content = '' }
        if ($Content.Length -gt $MaxChars) { $Content = $Content.Substring(0, $MaxChars) + "`n... [truncated]" }
        [System.IO.File]::WriteAllText($p, $Content, [System.Text.Encoding]::UTF8)
        return $p
    } catch { return $null }
}

function Write-RunLog {
    # Structured entry -> JSONL + human log + console.
    param(
        [string]$Level = 'Info',
        [string]$System = 'APP',
        [string]$Event = 'log',
        [string]$Message = '',
        [string]$Detail = '',
        [string]$Sample = ''
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $entry = [pscustomobject]@{
        ts      = $ts
        level   = $Level
        system  = $System
        event   = $Event
        message = $Message
        detail  = $Detail
        sample  = $Sample
    }
    $color = switch ($Level) { 'Success' { 'Green' } 'Warn' { 'Yellow' } 'Error' { 'Red' } 'Debug' { 'DarkGray' } default { 'White' } }
    Write-Host ("{0} [{1}] [{2}] {3}" -f $ts, $Level, $System, $Message) -ForegroundColor $color
    if ($Detail) { Write-Host ("      {0}" -f $Detail) -ForegroundColor 'DarkGray' }
    if ($script:JsonlPath) {
        try { ($entry | ConvertTo-Json -Compress -Depth 5) | Out-File -Append -FilePath $script:JsonlPath -Encoding UTF8 } catch { }
    }
    if ($script:LogFile) {
        try { ("{0} [{1}] [{2}] {3}" -f $ts, $Level, $System, $Message) | Out-File -Append -FilePath $script:LogFile -Encoding UTF8 } catch { }
        if ($Detail) {
            try { ("      " + $Detail) | Out-File -Append -FilePath $script:LogFile -Encoding UTF8 } catch { }
        }
    }
}

function Write-HttpLog {
    # Per-request instrumentation: system, endpoint, status, bytes, duration, outcome.
    param(
        [string]$System,
        [string]$Event,
        [string]$Url,
        [string]$Method = 'GET',
        [int]$Status = 0,
        [long]$Bytes = 0,
        [long]$Ms = 0,
        [int]$Retry = 0,
        [bool]$Ok = $true,
        [string]$Error = '',
        [string]$SampleFile = ''
    )
    $level = if ($Ok) { 'Info' } else { 'Error' }
    Write-RunLog -Level $level -System $System -Event $Event `
        -Message ("{0} {1} -> HTTP {2} ({3} bytes, {4} ms{5})" -f $Method, $Url, $(if ($Status) { $Status } else { '?' }), $Bytes, $Ms, $(if ($Retry) { ", retry $Retry" } else { '' })) `
        -Detail $Error -Sample $SampleFile
}

# ---------------------------------------------------------------------------
# Logging (console + GUI box + per-run file)
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warn','Error','Debug')]
        [string]$Level = 'Info'
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$Level] $Message"
    Write-Host $line
    if ($script:LogFile) {
        try { $line | Out-File -Append -FilePath $script:LogFile -Encoding UTF8 } catch { }
    }
    if ($script:JsonlPath) {
        try {
            $e = [pscustomobject]@{ ts = $ts; level = $Level; system = 'APP'; event = 'log'; message = $Message; detail = ''; sample = '' }
            ($e | ConvertTo-Json -Compress -Depth 5) | Out-File -Append -FilePath $script:JsonlPath -Encoding UTF8
        } catch { }
    }
    if ($script:LogBox) {
        $script:LogBox.AppendText($line + "`r`n")
        $script:LogBox.SelectionStart = $script:LogBox.TextLength
        $script:LogBox.ScrollToCaret()
    }
    if ($script:ProgressBar) {
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
function Get-Settings {
    $defaults = @{
        SiteId        = $null
        OutputDir     = Join-Path $env:USERPROFILE "Downloads"
        FromOffset    = -365
        ToOffset      = 0
        Fy            = 0
        FromPeriod    = 1
        ToPeriod      = 1
        AreaCode      = "1V"
        DistrictCode  = ""
        MappingScope  = "ALL"
        CheckedItems  = $null
    }
    if (-not (Test-Path $script:SettingsFile)) { return $defaults }
    try {
        $raw = Get-Content $script:SettingsFile -Raw | ConvertFrom-Json
        $merged = $defaults.Clone()
        foreach ($p in $raw.PSObject.Properties) {
            if ($defaults.ContainsKey($p.Name)) { $merged[$p.Name] = $p.Value }
        }
        return $merged
    } catch {
        return $defaults
    }
}

function Save-Settings {
    param([hashtable]$Settings)
    try {
        if (-not (Test-Path $script:SettingsDir)) {
            New-Item -ItemType Directory -Force -Path $script:SettingsDir | Out-Null
        }
        $Settings | ConvertTo-Json | Set-Content -Path $script:SettingsFile -Encoding UTF8
    } catch {
        Write-Log "Could not save settings: $($_.Exception.Message)" -Level Debug
    }
}

# ---------------------------------------------------------------------------
# Fiscal-year helpers
# ---------------------------------------------------------------------------
function Get-CurrentFy {
    $now = Get-Date
    if ($now.Month -ge 10) { return $now.Year + 1 }
    return $now.Year
}

function Get-CurrentAp {
    # USPS accounting period: 13 periods of 4 weeks starting Oct 1.
    $now = Get-Date
    $fyStart = Get-Date -Year ($(if ($now.Month -ge 10) { $now.Year } else { $now.Year - 1 })) -Month 10 -Day 1
    $weeks = [math]::Floor((($now.Date - $fyStart.Date).Days) / 28) + 1
    return [math]::Min(13, [math]::Max(1, [int]$weeks))
}

function Get-MaxAp {
    param([int]$Fy)
    # The NMARS forms give FY2022 and FY2028 a 14th accounting period.
    if ($Fy -eq 2022 -or $Fy -eq 2028) { return 14 }
    return 13
}

function Get-ApStartLocal {
    # Local fallback: AP1 = Saturday on/before Oct 1 of the prior calendar year;
    # each AP is 4 weeks.  (Server dates override this when reachable.)
    param([int]$Fy, [int]$Ap)
    $oct1 = Get-Date -Year ($Fy - 1) -Month 10 -Day 1
    $back = ([int]$oct1.DayOfWeek + 1) % 7
    $ap1 = $oct1.Date.AddDays(-$back).AddDays($script:ApStartOverrideDays)
    return $ap1.AddDays(28 * ($Ap - 1))
}

function Get-ApDates {
    # Returns @{ Start = <datetime>; End = <datetime> } for a FY/AP range.
    # Prefers the NMARS convert_apwk_to_date endpoint (authoritative); falls
    # back to the local fiscal calendar.
    param(
        [int]$Fy,
        [int]$FromAp,
        [int]$ToAp,
        [switch]$UseServer,
        [switch]$Quiet
    )
    $start = $null; $end = $null; $source = "LOCAL calendar"
    if ($UseServer) {
        try {
            $s1 = Invoke-NmarsGet -RelUrl "nm_pm_ecbm_completion_4wks.convert_apwk_to_date?p_fy=$Fy&p_ap=$FromAp&p_wk=1&p_isStart=1" -TimeoutSec 60
            $s2 = Invoke-NmarsGet -RelUrl "nm_pm_ecbm_completion_4wks.convert_apwk_to_date?p_fy=$Fy&p_ap=$ToAp&p_wk=4&p_isStart=0" -TimeoutSec 60
            $m1 = [regex]::Match($s1, '<STARTDATE[^>]*>([^<]+)</STARTDATE>', 'IgnoreCase')
            $m2 = [regex]::Match($s2, '<STARTDATE[^>]*>([^<]+)</STARTDATE>', 'IgnoreCase')
            if ($m1.Success -and $m2.Success) {
                $start = [datetime]::ParseExact($m1.Groups[1].Value.Trim(), 'MM/dd/yyyy', $null)
                $end   = [datetime]::ParseExact($m2.Groups[1].Value.Trim(), 'MM/dd/yyyy', $null)
                $source = "NMARS server"
            }
        } catch {
            Write-Log "AP date server lookup failed ($($_.Exception.Message)) - using local fiscal calendar" -Level Warn
        }
    }
    if ($null -eq $start -or $null -eq $end) {
        $start = Get-ApStartLocal -Fy $Fy -Ap $FromAp
        $end   = (Get-ApStartLocal -Fy $Fy -Ap $ToAp).AddDays(27)
    }
    if ($end -lt $start) { $t = $start; $start = $end; $end = $t }
    if (-not $Quiet) {
        Write-Log "Timeline: FY$Fy AP$FromAp-$ToAp -> $($start.ToString('MM/dd/yyyy')) .. $($end.ToString('MM/dd/yyyy'))  [$source]" -Level Info
    }
    return @{ Start = $start; End = $end; Source = $source }
}

# ---------------------------------------------------------------------------
# NMARS — HTTP + report catalog
# ---------------------------------------------------------------------------
function Invoke-NmarsGet {
    param(
        [Parameter(Mandatory)][string]$RelUrl,
        [int]$TimeoutSec = 300
    )
    $url = "$($script:NMARS_BASE)/$RelUrl"
    $params = @{
        Uri             = $url
        UseBasicParsing = $true
        UseDefaultCredentials = $true
        TimeoutSec      = $TimeoutSec
        Headers         = @{ Referer = "$($script:NMARS_BASE)/" }
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $resp = Invoke-WebRequest @params
        $sw.Stop()
        Write-HttpLog -System 'NMARS' -Event 'http' -Url $RelUrl -Method 'GET' `
            -Status ([int]$resp.StatusCode) -Bytes ([long]$resp.RawContentLength) -Ms $sw.ElapsedMilliseconds -Ok $true
        return $resp.Content
    } catch {
        $sw.Stop()
        $status = 0
        $ex = $_.Exception
        if ($ex.Response) { try { $status = [int]$ex.Response.StatusCode } catch { } }
        Write-HttpLog -System 'NMARS' -Event 'http' -Url $RelUrl -Method 'GET' `
            -Status $status -Bytes 0 -Ms $sw.ElapsedMilliseconds -Ok $false -Error $ex.Message
        throw
    }
}

function Test-NmarsReportContent {
    param([string]$Content)
    $out = [pscustomobject]@{ ok = $false; reason = 'empty' }
    if ([string]::IsNullOrWhiteSpace($Content)) { return $out }
    if ($Content -imatch 'login|session expired|unauthorized' -and $Content -inotmatch '<TABLE') {
        $out.reason = 'login/expired page'; return $out
    }
    if ($Content -imatch 'Period Selection' -and $Content -inotmatch '<TABLE') {
        $out.reason = 'criteria form served back (parameters rejected?)'; return $out
    }
    if ($Content -imatch '<TABLE' -or $Content -imatch 'NO DATA|No Data|NODATA') {
        $out.ok = $true; $out.reason = 'ok'; return $out
    }
    $out.reason = 'no <TABLE> and no no-data marker'; return $out
}

function Get-HtmlHeaders {
    # Extracts the first few <TH>/<TD> header texts from a response, for parser tuning.
    param([string]$Content, [int]$Max = 12)
    $ths = [regex]::Matches($Content, '<th[^>]*>([\s\S]*?)</th>', 'IgnoreCase')
    if ($ths.Count -eq 0) { $ths = [regex]::Matches($Content, '<td[^>]*class="[^"]*header[^"]*"[^>]*>([\s\S]*?)</td>', 'IgnoreCase') }
    $out = @()
    foreach ($m in $ths) {
        $t = ([regex]::Replace($m.Groups[1].Value, '<[^>]+>', ' ') -replace '\s+', ' ').Trim()
        if ($t) { $out += $t }
        if ($out.Count -ge $Max) { break }
    }
    return ($out -join ' | ')
}

function Save-NmarsReport {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$RelUrl,
        [Parameter(Mandatory)][string]$OutFile
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            Write-Log "  GET $RelUrl" -Level Info
            $content = Invoke-NmarsGet -RelUrl $RelUrl
            $check = Test-NmarsReportContent -Content $content
            if (-not $check.ok) {
                $hdrs = Get-HtmlHeaders -Content $content
                $sample = Save-Sample -Name $Name -Content $content
                throw "response does not look like a report ($($check.reason); bytes=$($content.Length); headers=[$hdrs])"
            }
            [System.IO.File]::WriteAllText($OutFile, $content, [System.Text.Encoding]::UTF8)
            $size = (Get-Item $OutFile).Length
            Write-Log "  SAVED $([math]::Round($size/1KB,1)) KB -> $(Split-Path $OutFile -Leaf)" -Level Success
            Write-RunLog -System 'NMARS' -Event 'save' -Message "saved $Name" `
                -Detail ("bytes=$size; headers=[{0}]" -f (Get-HtmlHeaders -Content $content))
            return $OutFile
        } catch {
            if ($script:CancelRequested) { throw "Cancelled by user" }
            $errMsg = $_.Exception.Message
            $sample = Save-Sample -Name $Name -Content $(try { $content } catch { $null })
            Write-RunLog -Level 'Error' -System 'NMARS' -Event 'report-fail' -Message "$Name failed (attempt $attempt)" `
                -Detail $errMsg -Sample $sample
            if ($attempt -le $script:NmarsRetries) {
                Write-Log "  attempt $attempt failed: $errMsg — retrying in $($script:RetrySleepSec)s" -Level Warn
                Start-Sleep -Seconds $script:RetrySleepSec
            } else {
                Write-Log "  FAILED after $($attempt-1) retry(ies): $errMsg" -Level Error
                $script:DiagNotes.Add("NMARS $Name : $errMsg")
                return $null
            }
        }
    }
}

function Get-NmarsSiteList {
    Write-Log "Loading NMARS site list..." -Level Info
    $content = Invoke-NmarsGet -RelUrl "nm_pm_ecbm_completion_4wks"
    $items = [regex]::Matches($content,
        '<OPTION\s+value="(?<id>\d+)"\s*>\s*(?<name>[^<]+)', 'IgnoreCase') |
        ForEach-Object {
            [pscustomobject]@{
                Id   = $_.Groups['id'].Value
                Name = [System.Net.WebUtility]::HtmlDecode($_.Groups['name'].Value.Trim())
            }
        } |
        Sort-Object -Property @{ Expression = { $_.Name }; Ascending = $true }
    Write-Log "$($items.Count) sites loaded (alphabetical)" -Level Success
    return $items
}

function Get-BypassAcronyms {
    param([Parameter(Mandatory)][string]$Site)
    $xmlText = Invoke-NmarsGet -RelUrl "NM_ecbm_tasks_bypassed.getAcronyms?p_site=$Site"
    $acronyms = @()
    try {
        $xml = [xml]$xmlText
        $nodes = $xml.SelectNodes('//ROW')
        if (-not $nodes -or $nodes.Count -eq 0) { $nodes = $xml.SelectNodes('//*[local-name()="ROW"]') }
        foreach ($n in $nodes) {
            $t = $null
            if ($n.FirstChild) { $t = $n.FirstChild.InnerText }
            if (-not [string]::IsNullOrWhiteSpace($t) -and $t -notmatch 'acronym') {
                $acronyms += $t.Trim()
            }
        }
    } catch {
        Write-Log "getAcronyms parse failed: $($_.Exception.Message)" -Level Warn
    }
    if ($acronyms.Count -eq 0) {
        Write-Log "  no acronyms returned for site $Site — skipping BYPASSED report" -Level Warn
    }
    return ($acronyms | Sort-Object -Unique)
}

# --- report id -> label map (order matters for the GUI list) ----------------
$script:ReportDefs = [ordered]@{
    PM4WKS_SITE   = @{ Label = "eCBM Completion % (last 4 wks) - SITE";       Group = "NMARS" }
    PM4WKS_NAT    = @{ Label = "eCBM Completion % (last 4 wks) - NATIONAL";   Group = "NMARS" }
    PM4WKS_AREA   = @{ Label = "eCBM Completion % (last 4 wks) - AREA";       Group = "NMARS" }
    PM4WKS_DIST   = @{ Label = "eCBM Completion % (last 4 wks) - DISTRICT";   Group = "NMARS" }
    PMAREA_SITE   = @{ Label = "eCBM Completion Stats - SITE";                Group = "NMARS" }
    PMAREA_AREA   = @{ Label = "eCBM Completion Stats - AREA";                Group = "NMARS" }
    BYPASSED      = @{ Label = "eCBM Tasks Bypassed (per machine acronym)";   Group = "NMARS" }
    TRENDS_SITE   = @{ Label = "PM Trends - SITE";                            Group = "NMARS" }
    TRENDS_NAT    = @{ Label = "PM Trends - NATIONAL";                        Group = "NMARS" }
    TRENDS_AREA   = @{ Label = "PM Trends - AREA";                            Group = "NMARS" }
    CONTRACT_SITE = @{ Label = "Contract Cost - SITE";                        Group = "NMARS" }
    CONTRACT_NAT  = @{ Label = "Contract Cost - NATIONAL";                    Group = "NMARS" }
    CONTRACT_AREA = @{ Label = "Contract Cost - AREA";                        Group = "NMARS" }
    EQUIP_SITE    = @{ Label = "Equipment Inventory - SITE (all conditions)"; Group = "NMARS" }
    EQUIP_NAT     = @{ Label = "Equipment Inventory - NATIONAL";              Group = "NMARS" }
    ISSUES_SITE   = @{ Label = "Parts Issues (Stocked) - SITE";               Group = "NMARS" }
    ISSUES_NAT    = @{ Label = "Parts Issues (Stocked) - NATIONAL";           Group = "NMARS" }
    STOCK         = @{ Label = "National Stock Room - SITE (parts on hand)";  Group = "NMARS" }
    ACRONYMS      = @{ Label = "Acronym Listing (machine-class dictionary)";  Group = "NMARS" }
    EOR_RUNS      = @{ Label = "WebEOR - PRODUCTION runs (down/idle/jam/maint time + AP) (CSV)"; Group = "WEBEOR" }
    EOR_MAINT     = @{ Label = "WebEOR - MAINTENANCE runs only (CSV, verifies maint windows)"; Group = "WEBEOR" }
    RPG_RUNVSPLAN = @{ Label = "WebEOR - Run vs RPG Plan Volume (CSV, plan-vs-actual efficiency)"; Group = "WEBEOR" }
    MAPPING       = @{ Label = "WebEOR - Machine Mapping (CSV)";              Group = "WEBEOR" }
}

function New-JobList {
    param(
        [hashtable]$Cfg,
        [string[]]$SelectedIds,
        [switch]$PlanOnly
    )
    $jobs = New-Object System.Collections.Generic.List[object]
    $enc = { param($v) [uri]::EscapeDataString([string]$v) }
    $fy = $Cfg.Fy; $fp = $Cfg.FromPeriod; $tp = $Cfg.ToPeriod
    $fw = 1; $tw = 4
    $dFrom = $Cfg.From; $dTo = $Cfg.To
    $ci = [System.Globalization.CultureInfo]::InvariantCulture

    $periodBase = "p_from_fy=$fy&p_from_period=$fp&p_from_wk=$fw&p_to_fy=$fy&p_to_period=$tp&p_to_wk=$tw"
    $sDate = if ($Cfg.StartDateStr) { (& $enc $Cfg.StartDateStr) } else { "" }
    $eDate = if ($Cfg.EndDateStr)   { (& $enc $Cfg.EndDateStr) }   else { "" }
    $dates = "&p_start_date=$sDate&p_end_date=$eDate"

    foreach ($id in $SelectedIds) {
        if (-not $script:ReportDefs.Contains($id)) { continue }
        $out = Join-Path $Cfg.RunDir "$($script:RunStamp)_$id"
        switch ($id) {

            "PM4WKS_SITE" {
                $u = "nm_pm_ecbm_completion_4wks.pm_comp_by_site?$periodBase&p_site_id=$($Cfg.SiteId)&p_back_b=0$dates"
                $jobs.Add(@{ Id=$id; Name="PM 4wks site"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "PM4WKS_NAT" {
                $u = "nm_pm_ecbm_completion_4wks.pm_comp_by_national?$periodBase$dates"
                $jobs.Add(@{ Id=$id; Name="PM 4wks national"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "PM4WKS_AREA" {
                $u = "nm_pm_ecbm_completion_4wks.pm_comp_by_area?$periodBase&p_area_code=$($Cfg.AreaCode)&p_back_b=0$dates"
                $jobs.Add(@{ Id=$id; Name="PM 4wks area $($Cfg.AreaCode)"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "PM4WKS_DIST" {
                if (-not $Cfg.DistrictCode) {
                    Write-Log "PM4WKS_DIST skipped — no District selected" -Level Warn
                    break
                }
                $u = "nm_pm_ecbm_completion_4wks.pm_comp_by_district?$periodBase&p_district_code=$($Cfg.DistrictCode)&p_back_b=0$dates"
                $jobs.Add(@{ Id=$id; Name="PM 4wks district $($Cfg.DistrictCode)"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "PMAREA_SITE" {
                $u = "nm_pm_ecbm_completion_area.pm_comp_by_site?$periodBase&p_site_id=$($Cfg.SiteId)&p_back_b=0$dates"
                $jobs.Add(@{ Id=$id; Name="PM area stats site"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "PMAREA_AREA" {
                $u = "nm_pm_ecbm_completion_area.pm_comp_by_area?$periodBase&p_area_code=$($Cfg.AreaCode)&p_back_b=0$dates"
                $jobs.Add(@{ Id=$id; Name="PM area stats area $($Cfg.AreaCode)"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "BYPASSED" {
                if ($PlanOnly) {
                    $u = "nm_ecbm_tasks_bypassed.tasks_bypassed?$periodBase$dates&p_site_id=$($Cfg.SiteId)&p_acronym=<ACRONYM>&p_equipno=&p_sort=1,2"
                    $jobs.Add(@{ Id=$id; Name="Tasks bypassed [<ACRONYM> x N]"; File="$($out)_<ACRONYM>.html"; Url=$u; Kind="NMARS" })
                    break
                }
                $acronyms = Get-BypassAcronyms -Site $Cfg.SiteId
                foreach ($a in $acronyms) {
                    $safeA = (& $enc $a)
                    $u = "nm_ecbm_tasks_bypassed.tasks_bypassed?$periodBase$dates&p_site_id=$($Cfg.SiteId)&p_acronym=$safeA&p_equipno=&p_sort=1,2"
                    $jobs.Add(@{ Id=$id; Name="Tasks bypassed [$a]"; File="$($out)_$safeA.html"; Url=$u; Kind="NMARS" })
                }
            }
            "TRENDS_SITE" {
                $u = "nm_pm_trends.pm_trend_by_site?p_site_id=$($Cfg.SiteId)&p_from_fy=$fy&p_from_period=$fp&p_to_fy=$fy&p_to_period=$tp"
                $jobs.Add(@{ Id=$id; Name="PM trends site"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "TRENDS_NAT" {
                $u = "nm_pm_trends.pm_trend_by_national?p_from_fy=$fy&p_from_period=$fp&p_to_fy=$fy&p_to_period=$tp"
                $jobs.Add(@{ Id=$id; Name="PM trends national"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "TRENDS_AREA" {
                $u = "nm_pm_trends.pm_trend_by_area?p_area_id=$($Cfg.AreaCode)&p_from_fy=$fy&p_from_period=$fp&p_to_fy=$fy&p_to_period=$tp"
                $jobs.Add(@{ Id=$id; Name="PM trends area $($Cfg.AreaCode)"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "CONTRACT_SITE" {
                $u = "nm_contract_cost.ccs_report?pi_rpt_param=$($Cfg.SiteId)&pi_start_dt=$(& $enc $dFrom.ToString('MM/dd/yyyy',$ci))&pi_end_dt=$(& $enc $dTo.ToString('MM/dd/yyyy',$ci))"
                $jobs.Add(@{ Id=$id; Name="Contract cost site"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "CONTRACT_NAT" {
                $u = "nm_contract_cost.ccan_report?pi_rpt_param=N&pi_start_dt=$(& $enc $dFrom.ToString('MM/dd/yyyy',$ci))&pi_end_dt=$(& $enc $dTo.ToString('MM/dd/yyyy',$ci))&pi_grp=D"
                $jobs.Add(@{ Id=$id; Name="Contract cost national"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "CONTRACT_AREA" {
                $u = "nm_contract_cost.ccan_report?pi_rpt_param=$($Cfg.AreaCode)&pi_start_dt=$(& $enc $dFrom.ToString('MM/dd/yyyy',$ci))&pi_end_dt=$(& $enc $dTo.ToString('MM/dd/yyyy',$ci))&pi_grp=D"
                $jobs.Add(@{ Id=$id; Name="Contract cost area $($Cfg.AreaCode)"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "EQUIP_SITE" {
                $u = "nm_equip_inventory.eis_report?pi_site=$($Cfg.SiteId)&pi_condition_code=ALL"
                $jobs.Add(@{ Id=$id; Name="Equip inventory site"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "EQUIP_NAT" {
                $u = "nm_equip_inventory.eian_report"
                $jobs.Add(@{ Id=$id; Name="Equip inventory national"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "ISSUES_SITE" {
                $u = "nm_issues_stocked.stock_by_site_summ?p_site_id=$($Cfg.SiteId)&p_from_fy=$($dFrom.Year)&p_from_period=$($dFrom.Month)&p_to_fy=$($dTo.Year)&p_to_period=$($dTo.Month)"
                $jobs.Add(@{ Id=$id; Name="Parts issues site"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "ISSUES_NAT" {
                $u = "nm_issues_stocked.stock_by_national_summ?p_from_fy=$($dFrom.Year)&p_from_period=$($dFrom.Month)&p_to_fy=$($dTo.Year)&p_to_period=$($dTo.Month)"
                $jobs.Add(@{ Id=$id; Name="Parts issues national"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "STOCK" {
                $u = "nm_national_stock.stockroom_by_site?p_site_id=$($Cfg.SiteId)&p_search_type=DESC&p_search_string=&p_boh_radio=-1"
                $jobs.Add(@{ Id=$id; Name="National stock room site"; File="$out.html"; Url=$u; Kind="NMARS" })
            }
            "ACRONYMS" {
                foreach ($cand in @("nm_acronym_listing","nm_acronym_listing.acronym_report","nm_acronym_listing.al_report","nm_acronym_listing.acronym_listing")) {
                    $jobs.Add(@{ Id=$id; Name="Acronym listing [$cand]"; File="$out.html"; Url=$cand; Kind="NMARS_TRY" })
                }
            }
            "EOR_RUNS" {
                $jobs.Add(@{ Id=$id; Name="WebEOR production runs CSV"; File="$out.csv"; Kind="EOR";
                             Filter=$script:EorFilterRuns; FieldNames=$script:EorPerfFieldNames; FieldLabels=$script:EorPerfFieldLabels })
            }
            "EOR_MAINT" {
                $jobs.Add(@{ Id=$id; Name="WebEOR maintenance-runs CSV"; File="$out.csv"; Kind="EOR";
                             Filter=$script:EorFilterMaint; FieldNames=$script:EorMaintFieldNames; FieldLabels=$script:EorMaintFieldLabels })
            }
            "RPG_RUNVSPLAN" {
                $jobs.Add(@{ Id=$id; Name="WebEOR Run vs RPG Plan CSV"; File="$out.csv"; Kind="WEBEOR_RPG" })
            }
            "MAPPING" {
                $jobs.Add(@{ Id=$id; Name="WebEOR machine mapping"; File="$out.csv"; Kind="MAPPING" })
            }
        }
    }
    return $jobs
}

# ---------------------------------------------------------------------------
# WebEOR — ported from WEBEOR.ps1 (proven flow: login -> viewer -> CSV)
# ---------------------------------------------------------------------------
function Connect-WebEor {
    param([Parameter(Mandatory)][int]$SiteId)
    Write-Log "Connecting to WebEOR (site $SiteId)..." -Level Info
    $s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $s.UserAgent = $script:UA
    $null = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/GUI_Login_View.asp" `
        -WebSession $s -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60
    $body = @{
        User="View"; Pword="View"; ChildSites=""; ServerType="Production"
        GroupList=""; FullName=""; EMail=""; ParentSiteID="$SiteId"; btnLogin="View"
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $r = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/GUI_Login_Validate.asp" `
            -Method POST -Body $body -WebSession $s -UseDefaultCredentials -UseBasicParsing `
            -Headers @{ Origin = $script:WEBEOR_BASE; Referer = "$($script:WEBEOR_BASE)/GUI_Login_View.asp" } `
            -TimeoutSec 60
        $sw.Stop()
        Write-HttpLog -System 'WebEOR' -Event 'login' -Url 'GUI_Login_Validate.asp' -Method 'POST' `
            -Status 200 -Bytes ([long]$r.RawContentLength) -Ms $sw.ElapsedMilliseconds -Ok $true
    } catch {
        $sw.Stop()
        $status = 0
        if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch { } }
        Write-HttpLog -System 'WebEOR' -Event 'login' -Url 'GUI_Login_Validate.asp' -Method 'POST' `
            -Status $status -Bytes 0 -Ms $sw.ElapsedMilliseconds -Ok $false -Error $_.Exception.Message
        throw
    }
    if ($r.Content -notmatch 'DMS/EOR Main Menu') {
        $sample = Save-Sample -Name "webeor_login_response" -Content $r.Content
        Write-RunLog -Level 'Error' -System 'WebEOR' -Event 'login' -Message "Login validation failed (no main menu)" `
            -Detail ("bytes=$($r.RawContentLength)") -Sample $sample
        throw "WebEOR login failed for site $SiteId (got $($r.RawContentLength) bytes)"
    }
    $script:LastWebEorMenu = $r.Content   # menu HTML for link discovery (RPG Run vs Plan, etc.)
    Write-Log "WebEOR login OK" -Level Success
    return $s
}

# --- JS string decoding helpers (used by both EOR and mapping viewers) --------
function Decode-JsString {
    param([string]$Raw)
    if ($null -eq $Raw) { return $null }
    $sentinel = [string][char]0x0001
    $s = $Raw
    $s = $s -replace '\\\\', $sentinel
    $s = [regex]::Replace($s, '\\x([0-9A-Fa-f]{2})', {
        param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16)
    })
    $s = [regex]::Replace($s, '\\u([0-9A-Fa-f]{4})', {
        param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16)
    })
    $s = $s -replace '\\"',  '"'
    $s = $s -replace "\\'",  "'"
    $s = $s -replace '\\n',  "`n"
    $s = $s -replace '\\r',  "`r"
    $s = $s -replace '\\t',  "`t"
    $s = $s.Replace($sentinel, '\')
    return $s
}

function Get-JsVarValue {
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$VarName
    )
    $m = [regex]::Match($Html,
        '\bvar\s+' + [regex]::Escape($VarName) + '\s*=\s*',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) { return $null }

    $start = $m.Index + $m.Length
    $inStr = $false
    $end   = -1
    $i     = $start
    while ($i -lt $Html.Length) {
        $c = $Html[$i]
        if ($inStr) {
            if ($c -eq '\') { $i += 2; continue }
            if ($c -eq '"') { $inStr = $false }
        } else {
            if     ($c -eq '"') { $inStr = $true }
            elseif ($c -eq ';') { $end = $i; break }
        }
        $i++
    }
    if ($end -lt 0) { return $null }

    $rhs = $Html.Substring($start, $end - $start)
    $sb = New-Object System.Text.StringBuilder
    $tokenPattern = '(?:"((?:[^"\\]|\\.)*)")|(?:\bString\s*\.\s*fromCharCode\s*\(\s*(0x[0-9A-Fa-f]+|\d+)\s*\))'
    foreach ($tok in [regex]::Matches($rhs, $tokenPattern)) {
        if ($tok.Groups[1].Success) {
            [void]$sb.Append((Decode-JsString $tok.Groups[1].Value))
        } elseif ($tok.Groups[2].Success) {
            $codeStr = $tok.Groups[2].Value
            $code = if ($codeStr -match '^0x') {
                [Convert]::ToInt32($codeStr, 16)
            } else {
                [int]$codeStr
            }
            [void]$sb.Append([char]$code)
        }
    }
    return $sb.ToString()
}

function Save-EorCsvSingleChunk {
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [int]      $SiteId,
        [Parameter(Mandatory)] [datetime] $From,
        [Parameter(Mandatory)] [datetime] $To,
        [Parameter(Mandatory)] [string]   $OutFile,
        [string]$Filter       = "All Runs - Maint. Runs Included",
        [string]$FieldNames   = "SITE_CODE,MACH_TYPE_DESC,MACHINE_NBR,OP_NBR,MSP,TOUR_CODE,RUN_NBR,START_DATE,END_DATE,TOTAL_PIECES_FED_CNT,MODSTRANS,DSIS",
        [string]$FieldLabels  = "Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS"
    )
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    Write-Log "  EOR viewer POST $($From.ToString('yyyy-MM-dd'))..$($To.ToString('yyyy-MM-dd')) [Filter=$Filter]" -Level Info

    $viewerBody = @{
        Help_Loc          = "/WebHelp/WebEOR_Help.htm#Viewing_End_of_Runs.htm"
        FirstLoad         = "FALSE"
        ClickDeleteButton = "FALSE"
        SiteID            = "$SiteId"
        ActiveSites       = "on"
        DisplaySiteOrZIP  = "EORZIP"
        MachineGroup      = "ALL"
        MachineType       = "ALL"
        MachineNo         = "ALL"
        OperationNo       = ""
        TourCode          = "ALL"
        SortProgram       = ""
        Filter            = $Filter
        DsisStatus        = ""
        ModsStatus        = ""
        StartDateFrom     = $From.ToString("MM/dd/yyyy", $ci)
        StartDateTo       = $To.ToString("MM/dd/yyyy", $ci)
        SF                = ""
        ST                = ""
        MultiSort         = "A_mach_type_desc,A_machine_nbr,A_start_date"
        FieldNames        = $FieldNames
        FieldLabels       = $FieldLabels
        RecordsListed     = "0"
        ARG1              = ""
        ARG2              = ""
        Active            = "Y"
        Update            = ""
        FullQuery         = ""
    }

    $viewer = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $viewer = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/EORViewer/GUI_EORView_Viewer.asp" `
            -Method POST -Body $viewerBody -WebSession $Session `
            -UseDefaultCredentials -UseBasicParsing `
            -Headers @{ Origin = $script:WEBEOR_BASE; Referer = "$($script:WEBEOR_BASE)/EORViewer/GUI_EORView_Viewer.asp" } `
            -TimeoutSec 900
        $sw.Stop()
        Write-HttpLog -System 'WebEOR' -Event 'eor-viewer' -Url 'EORViewer/GUI_EORView_Viewer.asp' -Method 'POST' `
            -Status 200 -Bytes ([long]$viewer.RawContentLength) -Ms $sw.ElapsedMilliseconds -Ok $true
    } catch {
        $sw.Stop()
        $status = 0
        if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch { } }
        Write-HttpLog -System 'WebEOR' -Event 'eor-viewer' -Url 'EORViewer/GUI_EORView_Viewer.asp' -Method 'POST' `
            -Status $status -Bytes 0 -Ms $sw.ElapsedMilliseconds -Ok $false -Error $_.Exception.Message
        throw "Viewer POST failed: $($_.Exception.Message)"
    }

    if ($viewer.Content -notmatch 'gstrTableSQL\s*=\s*"([^"]+)"') {
        $sample = Save-Sample -Name "eor_viewer_response" -Content $viewer.Content
        Write-RunLog -Level 'Error' -System 'WebEOR' -Event 'eor-viewer' -Message "EOR viewer missing gstrTableSQL" `
            -Detail ("bytes=$($viewer.RawContentLength)") -Sample $sample
        throw "EOR viewer response missing gstrTableSQL (bytes=$($viewer.RawContentLength))"
    }
    $viewerSql = $Matches[1] `
        -replace '\\x3C', '<' -replace '\\x3E', '>' `
        -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&'

    $expectedDate = $From.ToString("dd-MMM-yyyy", $ci)
    if ($viewerSql -notmatch [regex]::Escape($expectedDate)) {
        Write-Log "  WARN EOR viewer SQL missing '$expectedDate'" -Level Warn
    }

    # Prefer the viewer's own table metadata (works for any FieldNames set);
    # fall back to the legacy hardcoded mapping used by WEBEOR.ps1.
    $vHeader = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeader'
    $vGroup  = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeaderGroup'
    $vCols   = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableColumns'
    $vSrc    = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableSource'

    if (-not [string]::IsNullOrWhiteSpace($vHeader) -and -not [string]::IsNullOrWhiteSpace($vCols)) {
        $body = @{
            tabCsvGroup   = $vGroup
            tabCsvHeader  = $vHeader
            tabCsvColumns = $vCols
            tabCsvSource  = if ([string]::IsNullOrWhiteSpace($vSrc)) { "CLIENT" } else { $vSrc }
            tabCsvSQL     = $viewerSql
        }
        $csvHeader = $vHeader
        Write-Log "  using viewer table metadata ($($vHeader.Split(',').Count) columns)" -Level Debug
    } else {
        $body = @{
            tabCsvGroup   = " "
            tabCsvHeader  = "Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS"
            tabCsvColumns = "0,1,2,7,3,4,12,5,6,9,10,11"
            tabCsvSource  = "CLIENT"
            tabCsvSQL     = $viewerSql
        }
        $csvHeader = "Site,MType,MNo"
        Write-Log "  viewer metadata unavailable - using legacy column mapping" -Level Debug
    }

    $tmp = "$OutFile.tmp"
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    $null = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/ReportPage/GUI_Report_CSV.asp" `
        -Method POST -Body $body -WebSession $Session `
        -UseDefaultCredentials -UseBasicParsing `
        -Headers @{ Origin = $script:WEBEOR_BASE; Referer = "$($script:WEBEOR_BASE)/EORViewer/GUI_EORView_Viewer.asp" } `
        -OutFile $tmp -TimeoutSec 900

    if (-not (Test-Path $tmp)) { throw "CSV POST produced no file" }
    $size = (Get-Item $tmp).Length
    if ($size -lt 100) {
        $head = Get-Content $tmp -Raw -ErrorAction SilentlyContinue
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV response too small ($size bytes). Content: $head"
    }

    $fs = [System.IO.File]::OpenRead($tmp)
    try {
        $headBytes = New-Object byte[] ([Math]::Min(300, $size))
        $null = $fs.Read($headBytes, 0, $headBytes.Length)
        $headText = [System.Text.Encoding]::ASCII.GetString($headBytes)
    } finally { $fs.Close() }

    if ($headText -match '<HTML|Unrecongized|Unrecognized|Data Source') {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV endpoint returned HTML ($size bytes)"
    }
    $firstCol = ($csvHeader -split ',')[0]
    if ([string]::IsNullOrWhiteSpace($firstCol) -or $headText -notmatch "^$([regex]::Escape($firstCol))") {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV header mismatch ($size bytes)"
    }

    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    Move-Item $tmp $OutFile
    Write-Log "  EOR chunk OK $size bytes" -Level Success
    return (Get-Item $OutFile)
}

function Save-EorCsv {
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [int]      $SiteId,
        [Parameter(Mandatory)] [datetime] $From,
        [Parameter(Mandatory)] [datetime] $To,
        [Parameter(Mandatory)] [string]   $OutFile,
        [string]$Filter,
        [string]$FieldNames,
        [string]$FieldLabels,
        [scriptblock]$Progress
    )
    $dayCount = ($To.Date - $From.Date).Days + 1
    Write-Log "WebEOR EOR: $dayCount day(s): $($From.ToString('yyyy-MM-dd')) -> $($To.ToString('yyyy-MM-dd')) [Filter=$Filter]" -Level Info

    if ($dayCount -le $script:ChunkSize) {
        $f = Save-EorCsvSingleChunk -Session $Session -SiteId $SiteId -From $From -To $To `
            -OutFile $OutFile -Filter $Filter -FieldNames $FieldNames -FieldLabels $FieldLabels
        return $f
    }

    $chunks = @()
    $cursor = $From.Date
    while ($cursor -le $To.Date) {
        $end = $cursor.AddDays($script:ChunkSize - 1)
        if ($end -gt $To.Date) { $end = $To.Date }
        $chunks += @{ From = $cursor; To = $end }
        $cursor = $end.AddDays(1)
    }
    Write-Log "$dayCount days -> $($chunks.Count) chunks of <=$($script:ChunkSize) days" -Level Info

    $tempFiles    = @()
    $totalRows    = 0
    $failedChunks = @()

    for ($i = 0; $i -lt $chunks.Count; $i++) {
        if ($script:CancelRequested) {
            foreach ($tf in $tempFiles) { Remove-Item $tf -Force -ErrorAction SilentlyContinue }
            throw "Cancelled by user"
        }
        if ($i -gt 0 -and $script:InterChunkSleepSec -gt 0) {
            Start-Sleep -Seconds $script:InterChunkSleepSec
        }
        $c = $chunks[$i]
        $pct = [int](($i / $chunks.Count) * 100)
        if ($Progress) { & $Progress -Percent $pct -Message "EOR chunk $($i+1)/$($chunks.Count) — $($c.From.ToString('yyyy-MM-dd'))..$($c.To.ToString('yyyy-MM-dd'))" }
        Write-Log "EOR chunk $($i+1)/$($chunks.Count): $($c.From.ToString('yyyy-MM-dd'))..$($c.To.ToString('yyyy-MM-dd'))" -Level Info

        $chunkOut = "$OutFile.chunk$i.csv"
        $attempt  = 0
        $ok       = $false
        while (-not $ok -and $attempt -le $script:ChunkRetries) {
            if ($script:CancelRequested) {
                foreach ($tf in $tempFiles) { Remove-Item $tf -Force -ErrorAction SilentlyContinue }
                Remove-Item $chunkOut -Force -ErrorAction SilentlyContinue
                throw "Cancelled by user"
            }
            $attempt++
            try {
                $f = Save-EorCsvSingleChunk -Session $Session -SiteId $SiteId -From $c.From -To $c.To `
                    -OutFile $chunkOut -Filter $Filter -FieldNames $FieldNames -FieldLabels $FieldLabels
                $lines = (Get-Content $f.FullName | Measure-Object -Line).Lines
                $dataRows = [Math]::Max(0, $lines - 1)
                $totalRows += $dataRows
                Write-Log "  $dataRows rows, $([math]::Round($f.Length/1KB,1)) KB" -Level Info
                $tempFiles += $chunkOut
                $ok = $true
            } catch {
                Write-Log "  attempt $attempt failed: $($_.Exception.Message)" -Level Warn
                Remove-Item $chunkOut -Force -ErrorAction SilentlyContinue
                if ($attempt -le $script:ChunkRetries) {
                    Write-Log "  retrying in $($script:RetrySleepSec)s..." -Level Info
                    Start-Sleep -Seconds $script:RetrySleepSec
                }
            }
        }
        if (-not $ok) { $failedChunks += "$($c.From.ToString('yyyy-MM-dd'))..$($c.To.ToString('yyyy-MM-dd'))" }
    }

    if ($tempFiles.Count -eq 0) { throw "All EOR chunks failed - no CSV produced" }

    Write-Log "Merging $($tempFiles.Count) EOR chunk file(s)..." -Level Info
    $header  = Get-Content $tempFiles[0] -TotalCount 1
    $allData = New-Object System.Collections.Generic.List[string]
    foreach ($tf in $tempFiles) {
        $lines = Get-Content $tf
        for ($j = 1; $j -lt $lines.Count; $j++) { $allData.Add($lines[$j]) }
    }

    $tmpFinal = "$OutFile.tmpmerge"
    $writer = New-Object System.IO.StreamWriter($tmpFinal, $false, [System.Text.Encoding]::UTF8)
    try {
        $writer.WriteLine($header)
        foreach ($row in $allData) { $writer.WriteLine($row) }
    } finally { $writer.Close() }

    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    Move-Item $tmpFinal $OutFile
    foreach ($tf in $tempFiles) { Remove-Item $tf -Force -ErrorAction SilentlyContinue }

    $finalSize = (Get-Item $OutFile).Length
    Write-Log "SAVED $([math]::Round($finalSize/1KB,1)) KB (~$totalRows rows) -> $(Split-Path $OutFile -Leaf)" -Level Success
    if ($failedChunks.Count -gt 0) {
        Write-Log "WARNING: $($failedChunks.Count) EOR chunk(s) failed and are NOT in the merged file:" -Level Warn
        foreach ($r in $failedChunks) { Write-Log "  - $r" -Level Warn }
    }
    return (Get-Item $OutFile)
}

function Save-MachineMappingCsv {
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [string] $SiteId,
        [Parameter(Mandatory)] [string] $OutFile
    )
    if ($SiteId -ne "ALL" -and $SiteId -notmatch '^\d{1,5}$') {
        throw "Mapping SiteId must be a numeric site id or 'ALL' (got '$SiteId')"
    }
    $scopeLabel = if ($SiteId -eq "ALL") { "ALL sites" } else { "site $SiteId" }
    Write-Log "Machine mapping: GET viewer ($scopeLabel, MACHTYPE=ALL, Status=Y)..." -Level Info

    $viewerFields = [ordered]@{
        Help_Loc      = "/WebHelp/WebEOR_Help.htm#Mapping_Machine_Information.htm"
        MACHTYPE      = "ALL"
        SITEID        = $SiteId
        Status        = "Y"
        cmdRefresh    = "Refresh Table"
        SF            = ""
        ST            = ""
        ARG1          = ""
        ARG2          = ""
        MultiSort     = "A_SITE_NAME,A_MACHINE,A_MNO"
        FieldNames    = "GLOBAL_MACHINE_ID,SITE_NAME,MACHINE,EQUIPMENT_CODE,MNO,MAX_BIN_NBR,EOR_SW_VER,MODS_MACH_ID,NDSS_MACH_ID,MPE_ID,IOS_MACH_ID,LIFT_UNIT_NBR,FORMAT_EFFECTIVE_START_DATE,FORMAT_EFFECTIVE_END_DATE,SERIAL_NO,CAPITAL_ID_NO,D_ENABLED,FORMAT_DATETIME_DTM,SOURCE_LINK"
        FieldLabels   = "Global ID,Site Name,Machine Type,Model,MNO,Max Bin#,S/W Ver,MODS,NDSS,MPE ID,IOS,Lift Unit,Eff Start Date,Eff End Date,Serial,Capital ID,Status,Changed Date,Source"
        RecordsListed = "0"
    }
    $qs = ($viewerFields.GetEnumerator() | ForEach-Object {
        "$([uri]::EscapeDataString($_.Key))=$([uri]::EscapeDataString([string]$_.Value))"
    }) -join '&'

    $viewer = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/mapping/GUI_MachineInfo_Viewer.asp?$qs" `
        -WebSession $Session -UseDefaultCredentials -UseBasicParsing `
        -Headers @{ Referer = "$($script:WEBEOR_BASE)/mapping/GUI_MachineInfo_Viewer.asp" } `
        -TimeoutSec 900
    Write-Log "  mapping viewer returned $($viewer.RawContentLength) bytes" -Level Info

    $sql    = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableSQL'
    $header = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeader'
    $group  = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeaderGroup'
    $cols   = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableColumns'
    $src    = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableSource'

    if ([string]::IsNullOrWhiteSpace($sql))    { throw "Mapping viewer missing gstrTableSQL" }
    if ([string]::IsNullOrWhiteSpace($header)) { throw "Mapping viewer missing gstrTableHeader" }
    if ([string]::IsNullOrWhiteSpace($cols))   { throw "Mapping viewer missing gstrTableColumns" }

    if ($sql -notmatch '\bSELECT\b' -or $sql -notmatch '\bFROM\b' -or $sql -notmatch '\bORDER BY\b') {
        throw "Decoded mapping SQL is not a complete SELECT"
    }
    Write-Log "  mapping SQL decoded OK ($($sql.Length) chars)" -Level Success

    $body = @{
        tabCsvGroup   = $group
        tabCsvHeader  = $header
        tabCsvColumns = $cols
        tabCsvSource  = if ([string]::IsNullOrWhiteSpace($src)) { "CLIENT" } else { $src }
        tabCsvSQL     = $sql
    }

    $tmp = "$OutFile.tmp"
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    $null = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/ReportPage/GUI_Report_CSV.asp" `
        -Method POST -Body $body -WebSession $Session `
        -UseDefaultCredentials -UseBasicParsing `
        -Headers @{ Origin = $script:WEBEOR_BASE; Referer = "$($script:WEBEOR_BASE)/mapping/GUI_MachineInfo_Viewer.asp" } `
        -OutFile $tmp -TimeoutSec 900

    if (-not (Test-Path $tmp)) { throw "Mapping CSV POST produced no file" }
    $size = (Get-Item $tmp).Length
    if ($size -lt 100) {
        $head = Get-Content $tmp -Raw -ErrorAction SilentlyContinue
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "Mapping CSV response too small ($size bytes). Content: $head"
    }

    $fs = [System.IO.File]::OpenRead($tmp)
    try {
        $headBytes = New-Object byte[] ([Math]::Min(300, $size))
        $null = $fs.Read($headBytes, 0, $headBytes.Length)
        $headText = [System.Text.Encoding]::ASCII.GetString($headBytes)
    } finally { $fs.Close() }

    if ($headText -match '<HTML|Unrecongized|Unrecognized|Invalid Data Source|Data Source') {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "Mapping CSV endpoint returned HTML ($size bytes)"
    }
    if ($headText -notmatch 'Global ID|Site Name|Machine Type') {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "Mapping CSV header mismatch ($size bytes)"
    }

    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    Move-Item $tmp $OutFile
    Write-Log "  SAVED $([math]::Round($size/1KB,1)) KB -> $(Split-Path $OutFile -Leaf)" -Level Success
    return (Get-Item $OutFile)
}

# ---------------------------------------------------------------------------
# WebEOR — generic GET-based viewer -> CSV (used for RPG Run vs Plan, etc.)
# Same framework as the mapping viewer: viewer page embeds gstrTableSQL /
# gstrTableHeader / gstrTableColumns, and GUI_Report_CSV.asp renders the CSV.
# ---------------------------------------------------------------------------
function Save-WebEorGetViewerCsv {
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [string] $Url,
        [Parameter(Mandatory)] [string] $OutFile
    )
    Write-Log "  viewer GET $Url" -Level Info
    $viewer = Invoke-WebRequest -Uri $Url `
        -WebSession $Session -UseDefaultCredentials -UseBasicParsing `
        -Headers @{ Referer = "$($script:WEBEOR_BASE)/" } `
        -TimeoutSec 900

    $sql    = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableSQL'
    $header = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeader'
    $group  = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeaderGroup'
    $cols   = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableColumns'
    $src    = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableSource'

    if ([string]::IsNullOrWhiteSpace($sql))    { throw "viewer page has no gstrTableSQL (criteria page? bytes=$($viewer.RawContentLength))" }
    if ([string]::IsNullOrWhiteSpace($header)) { throw "viewer page has no gstrTableHeader" }
    if ([string]::IsNullOrWhiteSpace($cols))   { throw "viewer page has no gstrTableColumns" }

    $body = @{
        tabCsvGroup   = $group
        tabCsvHeader  = $header
        tabCsvColumns = $cols
        tabCsvSource  = if ([string]::IsNullOrWhiteSpace($src)) { "CLIENT" } else { $src }
        tabCsvSQL     = $sql
    }

    $tmp = "$OutFile.tmp"
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    $null = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/ReportPage/GUI_Report_CSV.asp" `
        -Method POST -Body $body -WebSession $Session `
        -UseDefaultCredentials -UseBasicParsing `
        -Headers @{ Origin = $script:WEBEOR_BASE; Referer = $Url } `
        -OutFile $tmp -TimeoutSec 900

    if (-not (Test-Path $tmp)) { throw "CSV POST produced no file" }
    $size = (Get-Item $tmp).Length
    if ($size -lt 100) {
        $head = Get-Content $tmp -Raw -ErrorAction SilentlyContinue
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV response too small ($size bytes). Content: $head"
    }
    $fs = [System.IO.File]::OpenRead($tmp)
    try {
        $headBytes = New-Object byte[] ([Math]::Min(300, $size))
        $null = $fs.Read($headBytes, 0, $headBytes.Length)
        $headText = [System.Text.Encoding]::ASCII.GetString($headBytes)
    } finally { $fs.Close() }

    if ($headText -match '<HTML|Unrecongized|Unrecognized|Invalid Data Source|Data Source') {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV endpoint returned HTML ($size bytes)"
    }
    $firstCol = ($header -split ',')[0]
    if ([string]::IsNullOrWhiteSpace($firstCol) -or $headText -notmatch "^$([regex]::Escape($firstCol))") {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV header mismatch ($size bytes). Head: $($headText.Substring(0,[Math]::Min(200,$headText.Length)))"
    }

    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    Move-Item $tmp $OutFile
    Write-Log "  SAVED $([math]::Round($size/1KB,1)) KB -> $(Split-Path $OutFile -Leaf)" -Level Success
    return (Get-Item $OutFile)
}

function Find-WebEorLink {
    # Searches the cached main-menu HTML for a link whose text or href matches
    # any of the patterns (first match wins).
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [string[]]$Patterns
    )
    if (-not $script:LastWebEorMenu) {
        $m = Invoke-WebRequest -Uri "$($script:WEBEOR_BASE)/" -WebSession $Session `
            -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60
        $script:LastWebEorMenu = $m.Content
    }
    $links = [regex]::Matches($script:LastWebEorMenu,
        '<a\s+[^>]*href\s*=\s*["''](?<href>[^"'']+)["''][^>]*>(?<text>.*?)</a>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::Singleline) |
        ForEach-Object {
            [pscustomobject]@{
                Href = $_.Groups['href'].Value
                Text = ([regex]::Replace($_.Groups['text'].Value, '<[^>]+>', ' ')).Trim()
            }
        }
    foreach ($p in $Patterns) {
        $hit = $links | Where-Object { $_.Text -match $p -or $_.Href -match $p } | Select-Object -First 1
        if ($hit) {
            $url = $hit.Href
            if ($url -notmatch '^https?://') {
                $url = "$($script:WEBEOR_BASE.TrimEnd('/'))/$($url.TrimStart('/'))"
            }
            Write-Log "  menu link '$($hit.Text)' -> $url" -Level Info
            return $url
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Job runner
# ---------------------------------------------------------------------------
function Invoke-Jobs {
    param(
        [hashtable]$Cfg,
        [object[]]$Jobs,
        [scriptblock]$Progress
    )
    $manifest = New-Object System.Collections.Generic.List[object]
    $webSession = $null
    $total = $Jobs.Count
    $done  = 0
    $failed = 0
    $acronymOk = $false

    foreach ($job in $Jobs) {
        if ($script:CancelRequested) {
            Write-Log "Cancel requested - stopping" -Level Warn
            break
        }
        $done++
        $pct = [int](($done / [math]::Max(1, $total)) * 100)
        if ($Progress) { & $Progress -Percent $pct -Message "$($job.Name) ($done/$total)" }

        switch ($job.Kind) {
            "NMARS" {
                try {
                    $f = Save-NmarsReport -Name $job.Name -RelUrl $job.Url -OutFile $job.File
                    $manifest.Add([pscustomobject]@{
                        Report = $job.Name; Source = "NMARS"; File = Split-Path $job.File -Leaf
                        Bytes = if ($f) { (Get-Item $f).Length } else { 0 }; Status = if ($f) { "OK" } else { "FAILED" }
                    })
                } catch {
                    $failed++
                    Write-Log "$($job.Name) FAILED: $($_.Exception.Message)" -Level Error
                    $manifest.Add([pscustomobject]@{ Report=$job.Name; Source="NMARS"; File=""; Bytes=0; Status="FAILED" })
                }
            }
            "NMARS_TRY" {
                # Acronym listing: try candidates until one looks like the listing page.
                if ($acronymOk) { break }
                try {
                    $content = Invoke-NmarsGet -RelUrl $job.Url
                    if ($content -match '<TABLE' -and $content -match 'Acronym') {
                        [System.IO.File]::WriteAllText($job.File, $content, [System.Text.Encoding]::UTF8)
                        $acronymOk = $true
                        Write-Log "  SAVED acronym listing via $($job.Url) -> $(Split-Path $job.File -Leaf)" -Level Success
                        $manifest.Add([pscustomobject]@{ Report=$job.Name; Source="NMARS"; File=Split-Path $job.File -Leaf; Bytes=(Get-Item $job.File).Length; Status="OK" })
                    } else {
                        Write-Log "  acronym candidate $($job.Url) did not return the listing - trying next" -Level Warn
                    }
                } catch {
                    Write-Log "  acronym candidate $($job.Url) failed: $($_.Exception.Message)" -Level Warn
                }
            }
            "EOR" {
                try {
                    if ($null -eq $webSession) { $webSession = Connect-WebEor -SiteId ([int]$Cfg.SiteId) }
                    $f = Save-EorCsv -Session $webSession -SiteId ([int]$Cfg.SiteId) `
                        -From $Cfg.From -To $Cfg.To -OutFile $job.File `
                        -Filter $job.Filter -FieldNames $job.FieldNames -FieldLabels $job.FieldLabels `
                        -Progress $Progress
                    $manifest.Add([pscustomobject]@{
                        Report = $job.Name; Source = "WebEOR"; File = Split-Path $job.File -Leaf
                        Bytes = $f.Length; Status = "OK"
                    })
                } catch {
                    $failed++
                    Write-Log "$($job.Name) FAILED: $($_.Exception.Message)" -Level Error
                    $manifest.Add([pscustomobject]@{ Report=$job.Name; Source="WebEOR"; File=""; Bytes=0; Status="FAILED" })
                }
            }
            "MAPPING" {
                try {
                    if ($null -eq $webSession) { $webSession = Connect-WebEor -SiteId ([int]$Cfg.SiteId) }
                    $scope = $Cfg.MappingScope
                    $f = Save-MachineMappingCsv -Session $webSession -SiteId $scope -OutFile $job.File
                    $manifest.Add([pscustomobject]@{
                        Report = $job.Name; Source = "WebEOR"; File = Split-Path $job.File -Leaf
                        Bytes = $f.Length; Status = "OK"
                    })
                } catch {
                    $failed++
                    Write-Log "$($job.Name) FAILED: $($_.Exception.Message)" -Level Error
                    $manifest.Add([pscustomobject]@{ Report=$job.Name; Source="WebEOR"; File=""; Bytes=0; Status="FAILED" })
                }
            }
            "WEBEOR_RPG" {
                # Run vs RPG Plan Volume: the page URL is discovered from the WebEOR
                # main menu (link text "Run vs Plan Volume"); best-effort by design.
                try {
                    if ($null -eq $webSession) { $webSession = Connect-WebEor -SiteId ([int]$Cfg.SiteId) }
                    $url = Find-WebEorLink -Session $webSession -Patterns @('Run vs Plan', 'Run vs RPG', 'RPG.*Plan')
                    if (-not $url) {
                        throw "No 'Run vs Plan' link found on the WebEOR main menu - open the RPG section in the browser to confirm the page name"
                    }
                    $qs = $script:RpgRunVsPlanQuery -replace '\{site\}', $Cfg.SiteId
                    $fullUrl = $url
                    if ($qs) { $fullUrl = $url + $(if ($url -match '\?') { '&' } else { '?' }) + $qs }
                    $f = Save-WebEorGetViewerCsv -Session $webSession -Url $fullUrl -OutFile $job.File
                    $manifest.Add([pscustomobject]@{
                        Report = $job.Name; Source = "WebEOR"; File = Split-Path $job.File -Leaf
                        Bytes = $f.Length; Status = "OK"
                    })
                    Write-Log "  Run-vs-Plan keys to join with production EOR: site, MODS date, machine type, operation, tour" -Level Info
                } catch {
                    $failed++
                    Write-Log "$($job.Name) FAILED: $($_.Exception.Message)" -Level Error
                    Write-Log "  hint: check the viewer's query string in a browser and set `$script:RpgRunVsPlanQuery at the top of this script" -Level Warn
                    $manifest.Add([pscustomobject]@{ Report=$job.Name; Source="WebEOR"; File=""; Bytes=0; Status="FAILED" })
                }
            }
        }

        if ($script:InterReportSleepMs -gt 0) {
            Start-Sleep -Milliseconds $script:InterReportSleepMs
        }
    }

    if ($manifest.Count -gt 0) {
        $manifestFile = Join-Path $Cfg.RunDir "$($script:RunStamp)_manifest.csv"
        $manifest | Export-Csv -Path $manifestFile -NoTypeInformation -Encoding UTF8
        Write-Log "Manifest written: $manifestFile" -Level Info
    }
    if ($Progress) { & $Progress -Percent 100 -Message "Done ($failed failed)" }
    return $failed
}

function Show-Plan {
    param([hashtable]$Cfg, [object[]]$Jobs)
    Write-Log "===== DRY RUN - planned requests ($($Jobs.Count)) =====" -Level Info
    foreach ($job in $Jobs) {
        switch ($job.Kind) {
            "NMARS"       { Write-Log ("GET  {0}/{1}  ->  {2}" -f $script:NMARS_BASE, $job.Url, (Split-Path $job.File -Leaf)) }
            "NMARS_TRY"   { Write-Log ("TRY  {0}/{1}  ->  {2}" -f $script:NMARS_BASE, $job.Url, (Split-Path $job.File -Leaf)) }
            "EOR"         { Write-Log ("WEBEOR  EOR CSV  site={0}  {1}..{2}  Filter='{3}'  ({4} cols)  (chunks of <=$($script:ChunkSize) days)  ->  {5}" -f $Cfg.SiteId, $Cfg.From.ToString('yyyy-MM-dd'), $Cfg.To.ToString('yyyy-MM-dd'), $job.Filter, ($job.FieldLabels -split ',').Count, (Split-Path $job.File -Leaf)) }
            "WEBEOR_RPG"  { Write-Log ("WEBEOR  Run vs RPG Plan (menu-discovered, query='{0}')  ->  {1}" -f ($script:RpgRunVsPlanQuery -replace '\{site\}', $Cfg.SiteId), (Split-Path $job.File -Leaf)) }
            "MAPPING"     { Write-Log ("WEBEOR  machine mapping  scope={0}  ->  {1}" -f $Cfg.MappingScope, (Split-Path $job.File -Leaf)) }
        }
    }
    Write-Log "===== end of dry run =====" -Level Success
}

# ---------------------------------------------------------------------------
# Config assembly (shared by GUI and CLI)
# ---------------------------------------------------------------------------
function Get-DefaultCfg {
    param([hashtable]$Settings)
    $now = Get-Date
    return @{
        SiteId       = if ($Settings.SiteId) { [string]$Settings.SiteId } else { "" }
        OutputDir    = [string]$Settings.OutputDir
        From         = $now.Date.AddDays([double]$Settings.FromOffset)
        To           = $now.Date.AddDays([double]$Settings.ToOffset)
        Fy           = if ($Settings.Fy -gt 2010) { [int]$Settings.Fy } else { Get-CurrentFy }
        FromPeriod   = if ($Settings.FromPeriod -ge 1) { [int]$Settings.FromPeriod } else { 1 }
        ToPeriod     = if ($Settings.ToPeriod -ge 1) { [int]$Settings.ToPeriod } else { Get-CurrentAp }
        AreaCode     = [string]$Settings.AreaCode
        DistrictCode = [string]$Settings.DistrictCode
        MappingScope = [string]$Settings.MappingScope
        RunDir       = $null
        StartDateStr = ""
        EndDateStr   = ""
    }
}

function Test-NmarsConnection {
    Write-Log "Testing NMARS connectivity..." -Level Info
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $c = Invoke-NmarsGet -RelUrl "nm_pm_ecbm_completion_4wks" -TimeoutSec 60
        $sw.Stop()
        $sample = Save-Sample -Name "nmars_form" -Content $c
        $sites = ([regex]::Matches($c, '<OPTION\s+value="\d+"') | Measure-Object).Count
        Write-RunLog -System 'NMARS' -Event 'connect-test' -Message "NMARS reachable" `
            -Detail ("ms=$($sw.ElapsedMilliseconds); bytes=$($c.Length); site-options=$sites") -Sample $sample
        return $true
    } catch {
        $sw.Stop()
        Write-RunLog -Level 'Error' -System 'NMARS' -Event 'connect-test' -Message "NMARS unreachable" -Detail $_.Exception.Message
        return $false
    }
}

function Test-WebEorConnection {
    param([string]$SiteId)
    Write-Log "Testing WebEOR connectivity (site $SiteId)..." -Level Info
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $null = Connect-WebEor -SiteId ([int]$SiteId)
        $sw.Stop()
        Write-RunLog -System 'WebEOR' -Event 'connect-test' -Message "WebEOR login OK" -Detail "ms=$($sw.ElapsedMilliseconds)"
        return $true
    } catch {
        $sw.Stop()
        $sample = Save-Sample -Name "webeor_login" -Content $script:LastWebEorMenu
        Write-RunLog -Level 'Error' -System 'WebEOR' -Event 'connect-test' -Message "WebEOR login failed" -Detail $_.Exception.Message -Sample $sample
        return $false
    }
}

function Write-DiagnosticsReport {
    param([string]$RunDir, [int]$FailedCount, [int]$JobCount)
    $p = Join-Path $RunDir "DIAGNOSTICS.md"
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# Diagnostics — NMARS + WebEOR grab run")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("- Ran: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("- Jobs: $JobCount planned, $FailedCount failed")
    [void]$sb.AppendLine("- Log: run_log.txt (human) + run_log.jsonl (structured)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## Issues found during this run")
    if ($script:DiagNotes.Count -eq 0) { [void]$sb.AppendLine("- (none logged)") }
    else { foreach ($n in $script:DiagNotes) { [void]$sb.AppendLine("- $n") } }
    $samples = @()
    if ($script:SampleDir -and (Test-Path $script:SampleDir)) {
        $samples = @(Get-ChildItem $script:SampleDir -Filter *.txt -ErrorAction SilentlyContinue)
    }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## Raw response samples")
    if ($samples.Count -eq 0) { [void]$sb.AppendLine("- (no failed responses were sampled)") }
    else { foreach ($s in $samples) { [void]$sb.AppendLine("- $($s.Name) ($($s.Length) bytes)") } }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## What to send back")
    [void]$sb.AppendLine("Attach run_log.jsonl + run_log.txt + DIAGNOSTICS.md (or the whole zip) so the")
    [void]$sb.AppendLine("parser/endpoint issues can be fixed.")
    [System.IO.File]::WriteAllText($p, $sb.ToString(), [System.Text.Encoding]::UTF8)
    return $p
}

# ---------------------------------------------------------------------------
# CLI entry
# ---------------------------------------------------------------------------
function Invoke-Cli {
    $settings = Get-Settings

    $manualDates = $false
    $fromD = $null
    if ($FromDate) {
        try { $fromD = [datetime]::ParseExact($FromDate, 'MM/dd/yyyy', $null); $manualDates = $true }
        catch { Write-Log "Bad -FromDate '$FromDate' (use MM/dd/yyyy)" -Level Error; return 2 }
    }
    $toD = $null
    if ($ToDate) {
        try { $toD = [datetime]::ParseExact($ToDate, 'MM/dd/yyyy', $null); $manualDates = $true }
        catch { Write-Log "Bad -ToDate '$ToDate' (use MM/dd/yyyy)" -Level Error; return 2 }
    }

    $cfg = Get-DefaultCfg -Settings $settings
    if ($SiteId) { $cfg.SiteId = $SiteId }
    if ($OutputDir) { $cfg.OutputDir = $OutputDir }
    if ($Fy)    { $cfg.Fy   = $Fy }
    if ($FromPeriod) { $cfg.FromPeriod = $FromPeriod }
    if ($ToPeriod)   { $cfg.ToPeriod   = $ToPeriod }
    if ($AreaCode)   { $cfg.AreaCode   = $AreaCode }
    if ($DistrictCode) { $cfg.DistrictCode = $DistrictCode }

    # Quarter preset: Q1=AP1-3, Q2=AP4-6, Q3=AP7-9, Q4=AP10-(13|14)
    if ($Quarter) {
        if ($Quarter -lt 1 -or $Quarter -gt 4) { Write-Log "Quarter must be 1..4" -Level Error; return 2 }
        $maxAp = Get-MaxAp -Fy $cfg.Fy
        $cfg.FromPeriod = 1 + 3 * ($Quarter - 1)
        $cfg.ToPeriod   = [math]::Min($maxAp, 3 * $Quarter)
    }

    $maxAp = Get-MaxAp -Fy $cfg.Fy
    if ($cfg.FromPeriod -lt 1 -or $cfg.FromPeriod -gt $maxAp) {
        Write-Log "FromPeriod must be 1..$maxAp for FY$($cfg.Fy) (got $($cfg.FromPeriod))" -Level Error; return 2
    }
    if ($cfg.ToPeriod -lt $cfg.FromPeriod -or $cfg.ToPeriod -gt $maxAp) {
        Write-Log "ToPeriod must be $($cfg.FromPeriod)..$maxAp for FY$($cfg.Fy) (got $($cfg.ToPeriod))" -Level Error; return 2
    }

    # One timeline drives every report: FY/AP -> real dates -> WebEOR EOR,
    # Contract Cost, and Parts Issues all use the same window.
    if ($manualDates) {
        $cfg.From = $fromD; $cfg.To = $toD
        $cfg.StartDateStr = $fromD.ToString('MM/dd/yyyy')
        $cfg.EndDateStr   = $toD.ToString('MM/dd/yyyy')
        Write-Log "Timeline: MANUAL override $($cfg.StartDateStr) .. $($cfg.EndDateStr)" -Level Info
    } else {
        $r = Get-ApDates -Fy $cfg.Fy -FromAp $cfg.FromPeriod -ToAp $cfg.ToPeriod -UseServer:(-not $DryRun)
        $cfg.From = $r.Start; $cfg.To = $r.End
        $cfg.StartDateStr = $r.Start.ToString('MM/dd/yyyy')
        $cfg.EndDateStr   = $r.End.ToString('MM/dd/yyyy')
    }

    if (-not $cfg.OutputDir) { $cfg.OutputDir = Join-Path $env:USERPROFILE "Downloads" }
    if (-not (Test-Path $cfg.OutputDir)) {
        try { New-Item -ItemType Directory -Force -Path $cfg.OutputDir | Out-Null }
        catch { Write-Log "Cannot create output dir $($cfg.OutputDir): $($_.Exception.Message)" -Level Error; return 2 }
    }
    $cfg.RunDir = Join-Path $cfg.OutputDir "NMARS-WebEOR_$($script:RunStamp)"
    if (-not $DryRun) {
        New-Item -ItemType Directory -Force -Path $cfg.RunDir | Out-Null
        $script:LogFile   = Join-Path $cfg.RunDir "run_log.txt"
        $script:JsonlPath = Join-Path $cfg.RunDir "run_log.jsonl"
        $script:SampleDir = Join-Path $cfg.RunDir "samples"
        New-Item -ItemType Directory -Force -Path $script:SampleDir | Out-Null
    }

    if ($TestConnection) {
        Write-RunLog -System 'APP' -Event 'start' -Message "Connection self-test started" -Detail "site=$($cfg.SiteId)"
        $okN = Test-NmarsConnection
        $okW = $true
        if ($cfg.SiteId) { $okW = Test-WebEorConnection -SiteId $cfg.SiteId }
        $howto = Join-Path $cfg.RunDir "HOW_TO_REPORT.txt"
        @(
            "HOW TO REPORT BACK",
            "===================",
            "",
            "NMARS reachable : $okN",
            "WebEOR login OK : $okW",
            "",
            "Send these files back for diagnosis:",
            "  - $($script:JsonlPath)",
            "  - $($script:LogFile)",
            "  - everything in $($script:SampleDir)"
        ) | Set-Content -Path $howto -Encoding UTF8
        Write-RunLog -System 'APP' -Event 'done' -Message "Self-test finished (NMARS=$okN, WebEOR=$okW)"
        if (-not $okN -and -not $okW) { return 1 }
        return 0
    }

    $selected = @()
    if ($Reports -eq "ALL") {
        $selected = @($script:ReportDefs.Keys)
    } else {
        $selected = @($Reports -split ',' | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ })
    }
    # Legacy aliases: EOR -> both run CSVs; EOR_PERF -> production runs CSV.
    if ($selected -contains "EOR") {
        $selected = @($selected | Where-Object { $_ -ne "EOR" }) + @("EOR_RUNS", "EOR_MAINT")
    }
    if ($selected -contains "EOR_PERF") {
        $selected = @($selected | Where-Object { $_ -ne "EOR_PERF" }) + @("EOR_RUNS")
    }

    $siteScoped = @($selected | Where-Object { $_ -in @("PM4WKS_SITE","PMAREA_SITE","BYPASSED","TRENDS_SITE","CONTRACT_SITE","EQUIP_SITE","ISSUES_SITE","STOCK","EOR_RUNS","EOR_MAINT","RPG_RUNVSPLAN") })
    if ($siteScoped.Count -gt 0 -and -not $cfg.SiteId) {
        Write-Log "Site-scoped reports selected but no -SiteId given." -Level Error
        Write-Log "  Site-scoped reports: $($siteScoped -join ', ')" -Level Error
        return 2
    }
    if ($cfg.SiteId -and $cfg.SiteId -notmatch '^\d{1,5}$') {
        Write-Log "SiteId must be numeric (got '$($cfg.SiteId)')" -Level Error
        return 2
    }

    Write-Log "Selected reports ($($selected.Count)): $($selected -join ', ')" -Level Info

    $jobs = New-JobList -Cfg $cfg -SelectedIds $selected -PlanOnly:$DryRun

    if ($DryRun) {
        Show-Plan -Cfg $cfg -Jobs $jobs
        return 0
    }

    Write-Log "Output folder: $($cfg.RunDir)" -Level Info
    Write-RunLog -System 'APP' -Event 'start' -Message "Run started" -Detail "site=$($cfg.SiteId); timeline=$($cfg.StartDateStr)..$($cfg.EndDateStr)"
    $failed = Invoke-Jobs -Cfg $cfg -Jobs $jobs -Progress $null
    Write-Log "Run complete. $($jobs.Count) job(s), $failed failed." -Level Info
    Write-RunLog -System 'APP' -Event 'done' -Message "Run complete" -Detail "jobs=$($jobs.Count); failed=$failed"

    $diag = Write-DiagnosticsReport -RunDir $cfg.RunDir -FailedCount $failed -JobCount $jobs.Count
    Write-Log "Diagnostics written: $diag" -Level Info

    $howto = Join-Path $cfg.RunDir "HOW_TO_REPORT.txt"
    @(
        "HOW TO REPORT BACK",
        "===================",
        "",
        "Run folder: $($cfg.RunDir)",
        "",
        "If anything failed or looks wrong, send back:",
        "  1. run_log.jsonl   (structured log of every request/response)",
        "  2. run_log.txt     (human-readable log)",
        "  3. DIAGNOSTICS.md  (issue summary)",
        "  4. samples\        (raw response snippets of failed items)",
        "  ... or just the whole zip if -Bundle was used.",
        "",
        "Tip: paste the tail of run_log.txt around any [Error] lines."
    ) | Set-Content -Path $howto -Encoding UTF8

    if ($Bundle) {
        $zip = Join-Path $cfg.OutputDir ("NMARS-WebEOR_{0}_bundle.zip" -f $script:RunStamp)
        try {
            Compress-Archive -Path (Join-Path $cfg.RunDir '*') -DestinationPath $zip -Force
            Write-RunLog -System 'APP' -Event 'bundle' -Message "Diagnostics bundle created" -Detail $zip
        } catch {
            Write-RunLog -Level 'Error' -System 'APP' -Event 'bundle' -Message "Could not create bundle" -Detail $_.Exception.Message
        }
    }

    if ($OpenStudio) {
        $studioRoot = $PSScriptRoot
        foreach ($page in @('index.html', 'Coverage-Studio.html')) {
            $cand = @(Join-Path $studioRoot $page, (Join-Path $studioRoot "4x10-DataStudio\$page")) |
                    Where-Object { Test-Path $_ } | Select-Object -First 1
            if ($cand) {
                Start-Process $cand
                Write-RunLog -System 'APP' -Event 'open' -Message "Opened studio page" -Detail $cand
            }
        }
    }

    if ($failed -gt 0) { return 1 }
    return 0
}

# ---------------------------------------------------------------------------
# List sites (CLI only)
# ---------------------------------------------------------------------------
if ($ListSites) {
    $sites = Get-NmarsSiteList
    $sites | ForEach-Object { "{0}`t{1}" -f $_.Id, $_.Name }
    exit 0
}

if ($NoGui) {
    exit (Invoke-Cli)
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
$script:Form    = New-Object System.Windows.Forms.Form
$script:Form.Text = "NMARS + WebEOR Data Grabber — 4/10 proposal evidence"
$script:Form.StartPosition = "CenterScreen"
$script:Form.Size = New-Object System.Drawing.Size(1000, 830)
$script:Form.MinimumSize = New-Object System.Drawing.Size(980, 770)

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:Form.Font = $font

$settings = Get-Settings
$cfgGui = Get-DefaultCfg -Settings $settings

# --- site row ----------------------------------------------------------------
$y = 10
$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Text = "Site:"; $lblSite.Location = New-Object System.Drawing.Point(14, $y + 5); $lblSite.Size = New-Object System.Drawing.Size(40, 22)
$script:Form.Controls.Add($lblSite)

$script:cmbSite = New-Object System.Windows.Forms.ComboBox
$script:cmbSite.Location = New-Object System.Drawing.Point(58, $y)
$script:cmbSite.Size = New-Object System.Drawing.Size(430, 24)
$script:cmbSite.DropDownStyle = "DropDown"   # editable: shows name, value holds id
$script:cmbSite.DisplayMember = "Name"
$script:cmbSite.ValueMember   = "Id"
$script:Form.Controls.Add($script:cmbSite)

$btnLoadSites = New-Object System.Windows.Forms.Button
$btnLoadSites.Text = "Load sites"; $btnLoadSites.Location = New-Object System.Drawing.Point(496, $y - 1); $btnLoadSites.Size = New-Object System.Drawing.Size(90, 26)
$script:Form.Controls.Add($btnLoadSites)

$lblArea = New-Object System.Windows.Forms.Label
$lblArea.Text = "Region:"; $lblArea.Location = New-Object System.Drawing.Point(598, $y + 5); $lblArea.Size = New-Object System.Drawing.Size(48, 22)
$script:Form.Controls.Add($lblArea)

$cmbArea = New-Object System.Windows.Forms.ComboBox
$cmbArea.Location = New-Object System.Drawing.Point(650, $y); $cmbArea.Size = New-Object System.Drawing.Size(120, 24)
$cmbArea.DropDownStyle = "DropDownList"
$cmbArea.Items.AddRange(@("1V — Eastern","1W — Western"))
$cmbArea.SelectedIndex = if ($cfgGui.AreaCode -eq "1W") { 1 } else { 0 }
$script:Form.Controls.Add($cmbArea)

$lblDist = New-Object System.Windows.Forms.Label
$lblDist.Text = "Division:"; $lblDist.Location = New-Object System.Drawing.Point(780, $y + 5); $lblDist.Size = New-Object System.Drawing.Size(52, 22)
$script:Form.Controls.Add($lblDist)

$cmbDist = New-Object System.Windows.Forms.ComboBox
$cmbDist.Location = New-Object System.Drawing.Point(836, $y); $cmbDist.Size = New-Object System.Drawing.Size(140, 24)
$cmbDist.DropDownStyle = "DropDownList"
$cmbDist.Items.Add("(none)")
$cmbDist.Items.AddRange(@("D01 — Chesapeake","D02 — Southeast","D03 — Lakeshores","D04 — South Atlantic","D05 — Western","D06 — Midwest","D07 — New England","D08 — New York Metro","D09 — Pacific Coastal","D10 — Southern California","D11 — Texas","D12 — Westshores","D13 — Gulf South","D14 — Central","D15 — Northern Plains","D16 — Florida","D17 — Northwest"))
$cmbDist.SelectedIndex = 0
$script:Form.Controls.Add($cmbDist)

# --- FY / periods row --------------------------------------------------------
$y = 44
$lblFy = New-Object System.Windows.Forms.Label
$lblFy.Text = "FY:"; $lblFy.Location = New-Object System.Drawing.Point(14, $y + 5); $lblFy.Size = New-Object System.Drawing.Size(40, 22)
$script:Form.Controls.Add($lblFy)

$numFy = New-Object System.Windows.Forms.NumericUpDown
$numFy.Location = New-Object System.Drawing.Point(58, $y); $numFy.Size = New-Object System.Drawing.Size(70, 22)
$numFy.Minimum = 2015; $numFy.Maximum = 2030; $numFy.Value = [int]$cfgGui.Fy
$script:Form.Controls.Add($numFy)

$lblFp = New-Object System.Windows.Forms.Label
$lblFp.Text = "From Period:"; $lblFp.Location = New-Object System.Drawing.Point(140, $y + 5); $lblFp.Size = New-Object System.Drawing.Size(76, 22)
$script:Form.Controls.Add($lblFp)

$numFp = New-Object System.Windows.Forms.NumericUpDown
$numFp.Location = New-Object System.Drawing.Point(218, $y); $numFp.Size = New-Object System.Drawing.Size(52, 22)
$numFp.Minimum = 1; $numFp.Maximum = 13; $numFp.Value = [int]$cfgGui.FromPeriod
$script:Form.Controls.Add($numFp)

$lblTp = New-Object System.Windows.Forms.Label
$lblTp.Text = "To Period:"; $lblTp.Location = New-Object System.Drawing.Point(282, $y + 5); $lblTp.Size = New-Object System.Drawing.Size(64, 22)
$script:Form.Controls.Add($lblTp)

$numTp = New-Object System.Windows.Forms.NumericUpDown
$numTp.Location = New-Object System.Drawing.Point(348, $y); $numTp.Size = New-Object System.Drawing.Size(52, 22)
$numTp.Minimum = 1; $numTp.Maximum = 14; $numTp.Value = [int](Get-MaxAp -Fy $cfgGui.Fy)
$script:Form.Controls.Add($numTp)

$lblPreset = New-Object System.Windows.Forms.Label
$lblPreset.Text = "Range:"; $lblPreset.Location = New-Object System.Drawing.Point(420, $y + 5); $lblPreset.Size = New-Object System.Drawing.Size(46, 22)
$script:Form.Controls.Add($lblPreset)

$cmbPreset = New-Object System.Windows.Forms.ComboBox
$cmbPreset.Location = New-Object System.Drawing.Point(468, $y); $cmbPreset.Size = New-Object System.Drawing.Size(96, 24)
$cmbPreset.DropDownStyle = "DropDownList"
$cmbPreset.Items.AddRange(@("Full FY","YTD","Q1","Q2","Q3","Q4","Custom"))
$cmbPreset.SelectedIndex = 0
$script:Form.Controls.Add($cmbPreset)

$lblWindow = New-Object System.Windows.Forms.Label
$lblWindow.Text = "Computed window: —"
$lblWindow.Location = New-Object System.Drawing.Point(574, $y + 5)
$lblWindow.Size = New-Object System.Drawing.Size(280, 22)
$lblWindow.AutoEllipsis = $true
$script:Form.Controls.Add($lblWindow)
$script:lblWindow = $lblWindow

$chkManual = New-Object System.Windows.Forms.CheckBox
$chkManual.Text = "Manual dates"
$chkManual.Location = New-Object System.Drawing.Point(862, $y + 3)
$chkManual.Size = New-Object System.Drawing.Size(112, 22)
$script:Form.Controls.Add($chkManual)

# Manual override row (hidden unless checked)
$yManual = 70
$lblM1 = New-Object System.Windows.Forms.Label
$lblM1.Text = "From:"; $lblM1.Location = New-Object System.Drawing.Point(14, $yManual + 5); $lblM1.Size = New-Object System.Drawing.Size(42, 22)
$script:Form.Controls.Add($lblM1)

$dtFrom = New-Object System.Windows.Forms.DateTimePicker
$dtFrom.Location = New-Object System.Drawing.Point(58, $yManual); $dtFrom.Size = New-Object System.Drawing.Size(110, 22)
$dtFrom.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtFrom.CustomFormat = "MM/dd/yyyy"
$dtFrom.Value = $cfgGui.From
$script:Form.Controls.Add($dtFrom)

$lblM2 = New-Object System.Windows.Forms.Label
$lblM2.Text = "to"; $lblM2.Location = New-Object System.Drawing.Point(174, $yManual + 5); $lblM2.Size = New-Object System.Drawing.Size(20, 22)
$script:Form.Controls.Add($lblM2)

$dtTo = New-Object System.Windows.Forms.DateTimePicker
$dtTo.Location = New-Object System.Drawing.Point(196, $yManual); $dtTo.Size = New-Object System.Drawing.Size(110, 22)
$dtTo.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtTo.CustomFormat = "MM/dd/yyyy"
$dtTo.Value = $cfgGui.To
$script:Form.Controls.Add($dtTo)

$lblM1.Visible = $false; $lblM2.Visible = $false
$dtFrom.Visible = $false; $dtTo.Visible = $false

# --- reports checklist -------------------------------------------------------
$y = 112
$grpReports = New-Object System.Windows.Forms.GroupBox
$grpReports.Text = "Reports"
$grpReports.Location = New-Object System.Drawing.Point(14, $y)
$grpReports.Size = New-Object System.Drawing.Size(960, 330)
$script:Form.Controls.Add($grpReports)

$clb = New-Object System.Windows.Forms.CheckedListBox
$clb.Location = New-Object System.Drawing.Point(14, 22)
$clb.Size = New-Object System.Drawing.Size(930, 296)
$clb.CheckOnClick = $true
$grpReports.Controls.Add($clb)

$script:ReportIdList = @($script:ReportDefs.Keys)
foreach ($k in $script:ReportDefs.Keys) {
    $null = $clb.Items.Add($script:ReportDefs[$k].Label)
    $clb.SetItemChecked($clb.Items.Count - 1, $true)
}
if ($settings.CheckedItems) {
    for ($i = 0; $i -lt $script:ReportIdList.Count; $i++) {
        $id = $script:ReportIdList[$i]
        $clb.SetItemChecked($i, ($settings.CheckedItems -contains $id))
    }
}

# --- output row --------------------------------------------------------------
$y = 452
$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = "Output"
$grpOut.Location = New-Object System.Drawing.Point(14, $y)
$grpOut.Size = New-Object System.Drawing.Size(960, 58)
$script:Form.Controls.Add($grpOut)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(14, 22)
$txtOut.Size = New-Object System.Drawing.Size(760, 22)
$txtOut.Text = [string]$cfgGui.OutputDir
$grpOut.Controls.Add($txtOut)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."; $btnBrowse.Location = New-Object System.Drawing.Point(786, 20); $btnBrowse.Size = New-Object System.Drawing.Size(80, 26)
$grpOut.Controls.Add($btnBrowse)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "Open"; $btnOpen.Location = New-Object System.Drawing.Point(872, 20); $btnOpen.Size = New-Object System.Drawing.Size(76, 26)
$grpOut.Controls.Add($btnOpen)

# --- mapping scope + action buttons ------------------------------------------
$y = 520
$lblScope = New-Object System.Windows.Forms.Label
$lblScope.Text = "WebEOR mapping scope:"; $lblScope.Location = New-Object System.Drawing.Point(14, $y + 6); $lblScope.Size = New-Object System.Drawing.Size(150, 22)
$script:Form.Controls.Add($lblScope)

$cmbScope = New-Object System.Windows.Forms.ComboBox
$cmbScope.Location = New-Object System.Drawing.Point(168, $y + 2); $cmbScope.Size = New-Object System.Drawing.Size(160, 24)
$cmbScope.DropDownStyle = "DropDownList"
$cmbScope.Items.AddRange(@("ALL (parent + children)","This site only"))
$cmbScope.SelectedIndex = if ($cfgGui.MappingScope -eq "ALL") { 0 } else { 1 }
$script:Form.Controls.Add($cmbScope)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"; $btnRun.Location = New-Object System.Drawing.Point(560, $y); $btnRun.Size = New-Object System.Drawing.Size(120, 30)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(31, 56, 100)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:Form.Controls.Add($btnRun)

$btnDry = New-Object System.Windows.Forms.Button
$btnDry.Text = "Dry Run (plan only)"; $btnDry.Location = New-Object System.Drawing.Point(690, $y); $btnDry.Size = New-Object System.Drawing.Size(130, 30)
$script:Form.Controls.Add($btnDry)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"; $btnCancel.Location = New-Object System.Drawing.Point(830, $y); $btnCancel.Size = New-Object System.Drawing.Size(70, 30)
$script:Form.Controls.Add($btnCancel)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"; $btnClose.Location = New-Object System.Drawing.Point(906, $y); $btnClose.Size = New-Object System.Drawing.Size(70, 30)
$script:Form.Controls.Add($btnClose)

# --- progress + log ----------------------------------------------------------
$y = 562
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(14, $y)
$progress.Size = New-Object System.Drawing.Size(960, 16)
$script:Form.Controls.Add($progress)
$script:ProgressBar = $progress

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(14, $y + 24)
$logBox.Size = New-Object System.Drawing.Size(960, 170)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::FromArgb(20, 22, 26)
$logBox.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 240)
$script:Form.Controls.Add($logBox)
$script:LogBox = $logBox

# --- event handlers ----------------------------------------------------------
$btnLoadSites.Add_Click({
    try {
        $btnLoadSites.Enabled = $false
        $sites = Get-NmarsSiteList
        $script:cmbSite.DataSource = $null
        $script:cmbSite.DataSource = @($sites)
        if ($sites.Count -gt 0) { $script:cmbSite.SelectedIndex = 0 }
    } catch {
        Write-Log "Could not load sites: $($_.Exception.Message)" -Level Error
    } finally {
        $btnLoadSites.Enabled = $true
    }
})

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $txtOut.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOut.Text = $dlg.SelectedPath
    }
})

$btnOpen.Add_Click({
    if (Test-Path $txtOut.Text) { Start-Process $txtOut.Text }
})

$btnClose.Add_Click({ $script:Form.Close() })

$btnCancel.Add_Click({
    $script:CancelRequested = $true
    Write-Log "Cancel requested - will stop after the current item" -Level Warn
})

function Read-GuiCfg {
    $siteId = ""
    if ($script:cmbSite.Text -match '^\d{1,5}$') {
        $siteId = $script:cmbSite.Text
    } elseif ($script:cmbSite.SelectedItem -and $script:cmbSite.SelectedItem.Id) {
        $siteId = [string]$script:cmbSite.SelectedItem.Id
    }
    $area = if ($cmbArea.SelectedIndex -eq 1) { "1W" } else { "1V" }
    $dist = ""
    if ($cmbDist.SelectedIndex -gt 0) { $dist = ($cmbDist.SelectedItem -split '\s+')[0] }

    $cfg = @{
        SiteId       = $siteId
        OutputDir    = $txtOut.Text.Trim()
        From         = $dtFrom.Value.Date
        To           = $dtTo.Value.Date
        Fy           = [int]$numFy.Value
        FromPeriod   = [int]$numFp.Value
        ToPeriod     = [int]$numTp.Value
        AreaCode     = $area
        DistrictCode = $dist
        MappingScope = if ($cmbScope.SelectedIndex -eq 0) { "ALL" } else { $siteId }
        RunDir       = $null
    }
    return $cfg
}

function Save-GuiSettings {
    $checked = @()
    for ($i = 0; $i -lt $script:ReportIdList.Count; $i++) {
        if ($clb.GetItemChecked($i)) { $checked += $script:ReportIdList[$i] }
    }
    $scope = if ($cmbScope.SelectedIndex -eq 0) { "ALL" } else { $cfgSiteId0 }
    Save-Settings -Settings @{
        SiteId       = $cfgSiteId0
        OutputDir    = $txtOut.Text.Trim()
        FromOffset   = ($dtFrom.Value.Date - (Get-Date).Date).Days
        ToOffset     = ($dtTo.Value.Date - (Get-Date).Date).Days
        Fy           = [int]$numFy.Value
        FromPeriod   = [int]$numFp.Value
        ToPeriod     = [int]$numTp.Value
        AreaCode     = if ($cmbArea.SelectedIndex -eq 1) { "1W" } else { "1V" }
        DistrictCode = if ($cmbDist.SelectedIndex -gt 0) { ($cmbDist.SelectedItem -split '\s+')[0] } else { "" }
        MappingScope = $scope
        CheckedItems = $checked
    }
}
$script:cfgSiteId0 = ""

$script:RunHandler = {
    param($dry)
    $cfg = Read-GuiCfg
    $selected = @()
    for ($i = 0; $i -lt $script:ReportIdList.Count; $i++) {
        if ($clb.GetItemChecked($i)) { $selected += $script:ReportIdList[$i] }
    }
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one report.", "Data Grabber", "OK", "Information") | Out-Null
        return
    }
    $siteScoped = @($selected | Where-Object { $_ -in @("PM4WKS_SITE","PMAREA_SITE","BYPASSED","TRENDS_SITE","CONTRACT_SITE","EQUIP_SITE","ISSUES_SITE","STOCK","EOR_RUNS","EOR_MAINT","RPG_RUNVSPLAN") })
    if ($siteScoped.Count -gt 0 -and -not $cfg.SiteId) {
        [System.Windows.Forms.MessageBox]::Show("Site-scoped reports need a Site.  Click 'Load sites' (on the intranet) or type the numeric site id.", "Data Grabber", "OK", "Information") | Out-Null
        return
    }
    if ($cfg.From -gt $cfg.To) {
        [System.Windows.Forms.MessageBox]::Show("From date must be on or before To date.", "Data Grabber", "OK", "Warning") | Out-Null
        return
    }
    $script:cfgSiteId0 = $cfg.SiteId
    Save-GuiSettings

    if (-not (Test-Path $cfg.OutputDir)) {
        try { New-Item -ItemType Directory -Force -Path $cfg.OutputDir | Out-Null }
        catch { Write-Log "Cannot create output dir: $($_.Exception.Message)" -Level Error; return }
    }
    $script:RunStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $cfg.RunDir = Join-Path $cfg.OutputDir "NMARS-WebEOR_$($script:RunStamp)"

    if (-not $dry) {
        New-Item -ItemType Directory -Force -Path $cfg.RunDir | Out-Null
        $script:LogFile = Join-Path $cfg.RunDir "run_log.txt"
    } else {
        $script:LogFile = $null
    }

    $script:CancelRequested = $false
    $btnRun.Enabled = $false; $btnDry.Enabled = $false; $btnCancel.Enabled = $true
    $progress.Value = 0

    $progCb = {
        param($Percent, $Message)
        if ($script:ProgressBar) { $script:ProgressBar.Value = [math]::Min(100, [int]$Percent) }
        if ($Message) { Write-Log $Message -Level Info }
    }

    try {
        Write-Log "Building job list..." -Level Info
        $jobs = New-JobList -Cfg $cfg -SelectedIds $selected -PlanOnly:$dry
        if ($dry) {
            Show-Plan -Cfg $cfg -Jobs $jobs
        } else {
            Write-Log "Output folder: $($cfg.RunDir)" -Level Info
            $failed = Invoke-Jobs -Cfg $cfg -Jobs $jobs -Progress $progCb
            Write-Log "Run complete. $($jobs.Count) job(s), $failed failed." -Level Info
            if ($failed -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Done. Files are in:`n$($cfg.RunDir)", "Data Grabber", "OK", "Information") | Out-Null
            } else {
                [System.Windows.Forms.MessageBox]::Show("Finished with $failed failed item(s). See the log.", "Data Grabber", "OK", "Warning") | Out-Null
            }
        }
    } catch {
        Write-Log "RUN FAILED: $($_.Exception.Message)" -Level Error
        [System.Windows.Forms.MessageBox]::Show("Run failed: $($_.Exception.Message)", "Data Grabber", "OK", "Error") | Out-Null
    } finally {
        $btnRun.Enabled = $true; $btnDry.Enabled = $true; $btnCancel.Enabled = $false
    }
}

$btnRun.Add_Click({ & $script:RunHandler -dry $false })
$btnDry.Add_Click({ & $script:RunHandler -dry $true })

# --- go ----------------------------------------------------------------------
Write-Log "NMARS + WebEOR Data Grabber ready.  Run on an intranet USPS workstation." -Level Info
Write-Log "Use 'Load sites' to populate the site list (requires intranet)." -Level Info
[System.Windows.Forms.Application]::Run($script:Form)
