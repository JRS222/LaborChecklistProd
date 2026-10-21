#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone WebEOR EOR CSV downloader + machine mapping downloader.

.NOTES
    Requires an on-intranet USPS workstation with Windows auth to
    https://webeor.usps.gov. Run under Windows PowerShell 5.1 (NOT pwsh 7).

    Set $env:WEBEOR_DEBUG_MAPPING = "1" before launching to dump
    machine-mapping diagnostics to %LOCALAPPDATA%\WebEORDownloader\.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'    # kills Invoke-WebRequest chatter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
$script:Base               = "https://webeor.usps.gov"
$script:UA                 = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
                             "AppleWebKit/537.36 (KHTML, like Gecko) " +
                             "Chrome/155.0.0.0 Safari/537.36 Edg/155.0.0.0"
$script:ChunkSize          = 14
$script:ChunkRetries       = 2
$script:RetrySleepSec      = 5
$script:InterChunkSleepSec = 1
$script:Session            = $null
$script:SessionSite        = $null
$script:CancelRequested    = $false

$script:SettingsDir  = Join-Path $env:LOCALAPPDATA "WebEORDownloader"
$script:SettingsFile = Join-Path $script:SettingsDir "settings.json"

[System.Net.ServicePointManager]::Expect100Continue = $false

$script:DefaultDownloads = Join-Path $env:USERPROFILE "Downloads"
if (-not (Test-Path $script:DefaultDownloads)) {
    New-Item -ItemType Directory -Force -Path $script:DefaultDownloads | Out-Null
}

# ---------------------------------------------------------------------------
# Debug dump
# ---------------------------------------------------------------------------
$script:DebugMapping = ($env:WEBEOR_DEBUG_MAPPING -eq '1')

function Write-MappingDebugDump {
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowEmptyString()][string]$Content,
        [string]$Extension = 'txt'
    )
    if (-not $script:DebugMapping) { return }
    try {
        if (-not (Test-Path $script:SettingsDir)) {
            New-Item -ItemType Directory -Force -Path $script:SettingsDir | Out-Null
        }
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $p = Join-Path $script:SettingsDir "mapping_debug_${Name}_${stamp}.${Extension}"
        if ($null -eq $Content) { $Content = '' }
        [System.IO.File]::WriteAllText($p, $Content)
        $bytes = [System.Text.Encoding]::UTF8.GetByteCount($Content)
        Write-DlLog "  [debug] wrote $p ($bytes bytes)" -Level Debug
    } catch {
        Write-DlLog "  [debug] dump failed for $Name : $($_.Exception.Message)" -Level Debug
    }
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
$script:LogBox = $null

function Write-DlLog {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warn','Error','Debug')]
        [string]$Level = 'Info'
    )
    $ts   = Get-Date -Format 'HH:mm:ss.fff'
    $line = "[$ts] $Message"

    $consoleColor = switch ($Level) {
        'Error'   { 'Red' }
        'Warn'    { 'Yellow' }
        'Success' { 'Green' }
        'Debug'   { 'DarkGray' }
        default   { 'Gray' }
    }
    Write-Host $line -ForegroundColor $consoleColor

    if ($script:LogBox -and -not $script:LogBox.IsDisposed) {
        $uiColor = switch ($Level) {
            'Error'   { [System.Drawing.Color]::Firebrick }
            'Warn'    { [System.Drawing.Color]::DarkGoldenrod }
            'Success' { [System.Drawing.Color]::ForestGreen }
            'Debug'   { [System.Drawing.Color]::Gray }
            default   { [System.Drawing.Color]::FromArgb(50,50,50) }
        }
        $script:LogBox.SelectionStart  = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0
        $script:LogBox.SelectionColor  = $uiColor
        $script:LogBox.AppendText($line + "`r`n")
        $script:LogBox.SelectionColor  = $script:LogBox.ForeColor
        $script:LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ---------------------------------------------------------------------------
# WebEOR — login
# ---------------------------------------------------------------------------
function Connect-WebEor {
    param([int]$SiteId)

    Write-DlLog "Connecting to WebEOR (site $SiteId)..." -Level Info
    $s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $s.UserAgent = $script:UA

    $null = Invoke-WebRequest -Uri "$($script:Base)/GUI_Login_View.asp" `
        -WebSession $s -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60

    $body = @{
        User="View"; Pword="View"; ChildSites=""; ServerType="Production"
        GroupList=""; FullName=""; EMail=""; ParentSiteID="$SiteId"; btnLogin="View"
    }
    $r = Invoke-WebRequest -Uri "$($script:Base)/GUI_Login_Validate.asp" `
        -Method POST -Body $body -WebSession $s -UseDefaultCredentials -UseBasicParsing `
        -Headers @{ Origin = $script:Base; Referer = "$($script:Base)/GUI_Login_View.asp" } `
        -TimeoutSec 60

    if ($r.Content -notmatch 'DMS/EOR Main Menu') {
        throw "Login failed for site $SiteId (got $($r.RawContentLength) bytes)"
    }
    Write-DlLog "Login OK" -Level Success
    return $s
}

# ---------------------------------------------------------------------------
# WebEOR — site list scraper (sorted by NAME, case-insensitive)
# ---------------------------------------------------------------------------
function Get-WebEorSiteList {
    Write-DlLog "Loading site list from WebEOR..." -Level Info
    $s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $s.UserAgent = $script:UA
    $landing = Invoke-WebRequest -Uri "$($script:Base)/GUI_Login_View.asp" `
        -WebSession $s -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60

    $items = [regex]::Matches($landing.Content,
        '<option\s+value\s*="(?<id>\d+)"\s*>(?<name>[^<]+)</option>', 'IgnoreCase') |
        ForEach-Object {
            [pscustomobject]@{
                Id   = $_.Groups['id'].Value
                Name = [System.Net.WebUtility]::HtmlDecode($_.Groups['name'].Value.Trim())
            }
        } |
        Sort-Object -Property @{ Expression = { $_.Name }; Ascending = $true }

    Write-DlLog "$($items.Count) sites loaded (alphabetical)" -Level Success
    return $items
}

# ---------------------------------------------------------------------------
# WebEOR — single-chunk EOR download
# ---------------------------------------------------------------------------
function Save-EorCsvSingleChunk {
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [int]      $SiteId,
        [Parameter(Mandatory)] [datetime] $From,
        [Parameter(Mandatory)] [datetime] $To,
        [Parameter(Mandatory)] [string]   $OutFile
    )

    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    Write-DlLog "  viewer POST $($From.ToString('yyyy-MM-dd'))..$($To.ToString('yyyy-MM-dd'))" -Level Info

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
        Filter            = "All Runs"
        DsisStatus        = ""
        ModsStatus        = ""
        StartDateFrom     = $From.ToString("MM/dd/yyyy", $ci)
        StartDateTo       = $To.ToString("MM/dd/yyyy", $ci)
        SF                = ""
        ST                = ""
        MultiSort         = "A_mach_type_desc,A_machine_nbr,A_start_date"
        FieldNames        = "SITE_CODE,MACH_TYPE_DESC,MACHINE_NBR,OP_NBR,MSP,TOUR_CODE,RUN_NBR,START_DATE,END_DATE,TOTAL_PIECES_FED_CNT,MODSTRANS,DSIS"
        FieldLabels       = "Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS"
        RecordsListed     = "0"
        ARG1              = ""
        ARG2              = ""
        Active            = "Y"
        Update            = ""
        FullQuery         = ""
    }

    try {
        $viewer = Invoke-WebRequest -Uri "$($script:Base)/EORViewer/GUI_EORView_Viewer.asp" `
            -Method POST -Body $viewerBody -WebSession $Session `
            -UseDefaultCredentials -UseBasicParsing `
            -Headers @{ Origin = $script:Base; Referer = "$($script:Base)/EORViewer/GUI_EORView_Viewer.asp" } `
            -TimeoutSec 900
    } catch {
        throw "Viewer POST failed: $($_.Exception.Message)"
    }

    if ($viewer.Content -notmatch 'gstrTableSQL\s*=\s*"([^"]+)"') {
        throw "Viewer response missing gstrTableSQL (bytes=$($viewer.RawContentLength))"
    }
    $viewerSql = $Matches[1] `
        -replace '\\x3C', '<' -replace '\\x3E', '>' `
        -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&'

    $expectedDate = $From.ToString("dd-MMM-yyyy", $ci)
    if ($viewerSql -notmatch [regex]::Escape($expectedDate)) {
        Write-DlLog "  WARN viewer SQL missing '$expectedDate'" -Level Warn
    }

    $body = @{
        tabCsvGroup   = " "
        tabCsvHeader  = "Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS"
        tabCsvColumns = "0,1,2,7,3,4,12,5,6,9,10,11"
        tabCsvSource  = "CLIENT"
        tabCsvSQL     = $viewerSql
    }

    $tmp = "$OutFile.tmp"
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue

    Write-DlLog "  csv POST ($($viewerSql.Length) chars of SQL)" -Level Info
    try {
        $null = Invoke-WebRequest -Uri "$($script:Base)/ReportPage/GUI_Report_CSV.asp" `
            -Method POST -Body $body -WebSession $Session `
            -UseDefaultCredentials -UseBasicParsing `
            -Headers @{ Origin = $script:Base; Referer = "$($script:Base)/EORViewer/GUI_EORView_Viewer.asp" } `
            -OutFile $tmp -TimeoutSec 900
    } catch {
        $partial = if (Test-Path $tmp) { (Get-Item $tmp).Length } else { 0 }
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV POST failed after $partial bytes: $($_.Exception.Message)"
    }

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
    if ($headText -notmatch '^Site,MType,MNo') {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "CSV header mismatch ($size bytes)"
    }

    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    Move-Item $tmp $OutFile
    Write-DlLog "  OK $size bytes -> $(Split-Path $OutFile -Leaf)" -Level Success
    return (Get-Item $OutFile)
}

# ---------------------------------------------------------------------------
# WebEOR — chunked EOR download + merge
# ---------------------------------------------------------------------------
function Save-EorCsv {
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [int]      $SiteId,
        [Parameter(Mandatory)] [datetime] $From,
        [Parameter(Mandatory)] [datetime] $To,
        [Parameter(Mandatory)] [string]   $OutFile,
        [scriptblock]$Progress
    )

    $dayCount = ($To.Date - $From.Date).Days + 1
    Write-DlLog "Fetching $dayCount day(s): $($From.ToString('yyyy-MM-dd')) → $($To.ToString('yyyy-MM-dd'))" -Level Info

    if ($Progress) { & $Progress -Percent 0 -Message "Starting download..." }

    if ($dayCount -le $script:ChunkSize) {
        $f = Save-EorCsvSingleChunk -Session $Session -SiteId $SiteId -From $From -To $To -OutFile $OutFile
        if ($Progress) { & $Progress -Percent 100 -Message "Done" }
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

    Write-DlLog "$dayCount days → $($chunks.Count) chunks of <=$($script:ChunkSize) days" -Level Info

    $tempFiles    = @()
    $totalRows    = 0
    $failedChunks = @()

    for ($i = 0; $i -lt $chunks.Count; $i++) {
        if ($i -gt 0 -and $script:InterChunkSleepSec -gt 0) {
            Start-Sleep -Seconds $script:InterChunkSleepSec
        }
        if ($script:CancelRequested) {
            Write-DlLog "Cancel requested — cleaning up partial chunks" -Level Warn
            foreach ($tf in $tempFiles) { Remove-Item $tf -Force -ErrorAction SilentlyContinue }
            throw "Cancelled by user"
        }

        $c = $chunks[$i]
        $pct = [int](($i / $chunks.Count) * 100)
        if ($Progress) {
            & $Progress -Percent $pct -Message "Chunk $($i+1) of $($chunks.Count) — $($c.From.ToString('yyyy-MM-dd')) to $($c.To.ToString('yyyy-MM-dd'))"
        }
        Write-DlLog "Chunk $($i+1)/$($chunks.Count): $($c.From.ToString('yyyy-MM-dd'))..$($c.To.ToString('yyyy-MM-dd'))" -Level Info

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
                $f = Save-EorCsvSingleChunk -Session $Session -SiteId $SiteId `
                    -From $c.From -To $c.To -OutFile $chunkOut
                $lines = (Get-Content $f.FullName | Measure-Object -Line).Lines
                $dataRows = [Math]::Max(0, $lines - 1)
                $totalRows += $dataRows
                Write-DlLog "  $dataRows rows, $([math]::Round($f.Length/1KB,1)) KB" -Level Info
                $tempFiles += $chunkOut
                $ok = $true
            } catch {
                Write-DlLog "  attempt $attempt failed: $($_.Exception.Message)" -Level Warn
                Remove-Item $chunkOut -Force -ErrorAction SilentlyContinue
                if ($attempt -le $script:ChunkRetries) {
                    Write-DlLog "  retrying in $($script:RetrySleepSec)s..." -Level Info
                    Start-Sleep -Seconds $script:RetrySleepSec
                }
            }
        }
        if (-not $ok) {
            $failedChunks += "$($c.From.ToString('yyyy-MM-dd'))..$($c.To.ToString('yyyy-MM-dd'))"
        }
    }

    if ($tempFiles.Count -eq 0) { throw "All chunks failed — no CSV produced" }

    if ($Progress) { & $Progress -Percent 95 -Message "Merging $($tempFiles.Count) chunk(s)..." }
    Write-DlLog "Merging $($tempFiles.Count) chunk file(s)..." -Level Info

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
    Write-DlLog "SAVED $([math]::Round($finalSize/1KB,1)) KB (~$totalRows rows)" -Level Success

    if ($failedChunks.Count -gt 0) {
        Write-DlLog "WARNING: $($failedChunks.Count) chunk(s) failed and are NOT in the merged file:" -Level Warn
        foreach ($r in $failedChunks) { Write-DlLog "  - $r" -Level Warn }
    }

    if ($Progress) { & $Progress -Percent 100 -Message "Done" }
    return (Get-Item $OutFile)
}

# ---------------------------------------------------------------------------
# WebEOR — machine mapping CSV
#   $SiteId accepts a numeric id ("19101") or the literal "ALL".
#   With "ALL", WebEOR returns every machine visible to the account —
#   parent plant AND all child sites. With a numeric id, only that site.
#   Handles both observed gstrTableSQL serializations (single literal with
#   \xHH escapes; concatenation with String.fromCharCode).
# ---------------------------------------------------------------------------
function Save-MachineMappingCsv {
    param(
        [Parameter(Mandatory)] $Session,
        [Parameter(Mandatory)] [string] $SiteId,
        [string] $MachineType = "ALL",
        [ValidateSet("Y","N","ALL")]
        [string] $Status = "Y",
        [Parameter(Mandatory)] [string] $OutFile
    )

    if ($SiteId -ne "ALL" -and $SiteId -notmatch '^\d{1,5}$') {
        throw "SiteId must be a numeric site id or 'ALL' (got '$SiteId')"
    }

    $scopeLabel = if ($SiteId -eq "ALL") { "ALL sites" } else { "site $SiteId" }
    Write-DlLog "  [mapping] GET viewer ($scopeLabel, MACHTYPE=$MachineType, Status=$Status)..." -Level Info

    $viewerFields = [ordered]@{
        Help_Loc      = "/WebHelp/WebEOR_Help.htm#Mapping_Machine_Information.htm"
        MACHTYPE      = $MachineType
        SITEID        = $SiteId
        Status        = $Status
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

    try {
        $viewer = Invoke-WebRequest -Uri "$($script:Base)/mapping/GUI_MachineInfo_Viewer.asp?$qs" `
            -WebSession $Session -UseDefaultCredentials -UseBasicParsing `
            -Headers @{ Referer = "$($script:Base)/mapping/GUI_MachineInfo_Viewer.asp" } `
            -TimeoutSec 900
    } catch {
        throw "Mapping viewer GET failed: $($_.Exception.Message)"
    }

    Write-DlLog "  [mapping] viewer returned $($viewer.RawContentLength) bytes" -Level Info
    Write-MappingDebugDump -Name 'viewer' -Content $viewer.Content -Extension 'html'

    # --- JS tokenizer (handles both serialization shapes) -----------------
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

    $sql    = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableSQL'
    $header = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeader'
    $group  = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableHeaderGroup'
    $cols   = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableColumns'
    $src    = Get-JsVarValue -Html $viewer.Content -VarName 'gstrTableSource'

    if ([string]::IsNullOrWhiteSpace($sql))    { throw "Mapping viewer missing gstrTableSQL" }
    if ([string]::IsNullOrWhiteSpace($header)) { throw "Mapping viewer missing gstrTableHeader" }
    if ([string]::IsNullOrWhiteSpace($cols))   { throw "Mapping viewer missing gstrTableColumns" }

    Write-DlLog "  [mapping] decoded SQL: $($sql.Length) chars" -Level Info
    Write-MappingDebugDump -Name 'mappingSql_decoded' -Content $sql

    if ($sql -notmatch '\bSELECT\b' -or $sql -notmatch '\bFROM\b' -or $sql -notmatch '\bORDER BY\b') {
        Write-MappingDebugDump -Name 'mappingSql_suspicious' -Content $sql
        throw "Decoded SQL is not a complete SELECT"
    }
    if ($SiteId -ne "ALL" -and $sql -notmatch [regex]::Escape($SiteId)) {
        Write-DlLog "  WARN decoded SQL missing site id $SiteId" -Level Warn
    }
    if ($SiteId -eq "ALL" -and $sql -match 'AND\s+machine\.site_id\s*=\s*\d+') {
        Write-DlLog "  WARN SITEID=ALL but SQL still contains a site-id filter" -Level Warn
    }
    Write-DlLog "  [mapping] SQL sanity check passed" -Level Success

    $body = @{
        tabCsvGroup   = $group
        tabCsvHeader  = $header
        tabCsvColumns = $cols
        tabCsvSource  = if ([string]::IsNullOrWhiteSpace($src)) { "CLIENT" } else { $src }
        tabCsvSQL     = $sql
    }

    $tmp = "$OutFile.tmp"
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue

    Write-DlLog "  [mapping] csv POST (scope=$scopeLabel, cols=$cols)" -Level Info
    try {
        $null = Invoke-WebRequest -Uri "$($script:Base)/ReportPage/GUI_Report_CSV.asp" `
            -Method POST -Body $body -WebSession $Session `
            -UseDefaultCredentials -UseBasicParsing `
            -Headers @{ Origin = $script:Base; Referer = "$($script:Base)/mapping/GUI_MachineInfo_Viewer.asp" } `
            -OutFile $tmp -TimeoutSec 900
    } catch {
        $partial = if (Test-Path $tmp) { (Get-Item $tmp).Length } else { 0 }
        if ((Test-Path $tmp) -and $script:DebugMapping) {
            $rawDump = Get-Content $tmp -Raw -ErrorAction SilentlyContinue
            Write-MappingDebugDump -Name 'csv_partial_response' -Content $rawDump -Extension 'txt'
        }
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "Mapping CSV POST failed after $partial bytes: $($_.Exception.Message)"
    }

    if (-not (Test-Path $tmp)) { throw "Mapping CSV POST produced no file" }
    $size = (Get-Item $tmp).Length

    if ($script:DebugMapping) {
        $rawDump = Get-Content $tmp -Raw -ErrorAction SilentlyContinue
        Write-MappingDebugDump -Name 'csv_response' -Content $rawDump -Extension 'csv'
    }

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
        throw "Mapping CSV endpoint returned HTML ($size bytes): $($headText.Substring(0,[Math]::Min(160,$headText.Length)))"
    }
    if ($headText -notmatch 'Global ID|Site Name|Machine Type') {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "Mapping CSV header mismatch ($size bytes). Head: $($headText.Substring(0,[Math]::Min(200,$headText.Length)))"
    }

    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    Move-Item $tmp $OutFile
    Write-DlLog "  [mapping] OK $size bytes -> $OutFile" -Level Success
    return (Get-Item $OutFile)
}

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
function Load-Settings {
    $defaults = @{
        SiteId       = $null
        OutputDir    = $script:DefaultDownloads
        FromOffset   = -1
        ToOffset     = -1
        RememberRange = $true
        MappingScope = 1          # 0 = this site only, 1 = ALL sites (parent + children)
        FetchMapping = $false
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
        Write-DlLog "Could not save settings: $($_.Exception.Message)" -Level Debug
    }
}

# ---------------------------------------------------------------------------
# The dialog
# ---------------------------------------------------------------------------
function Show-WebEorDownloaderDialog {

    $settings = Load-Settings

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text          = "WebEOR EOR CSV Downloader"
    $dlg.ClientSize    = New-Object System.Drawing.Size(860, 700)
    $dlg.MinimumSize   = New-Object System.Drawing.Size(760, 560)
    $dlg.StartPosition = "CenterScreen"
    $dlg.Font          = New-Object System.Drawing.Font("Segoe UI", 9)
    $dlg.BackColor     = [System.Drawing.Color]::FromArgb(245,247,250)

    # ---- Site row ----
    $lblSite = New-Object System.Windows.Forms.Label
    $lblSite.Text = "Site:"
    $lblSite.Location = New-Object System.Drawing.Point(15, 18)
    $lblSite.Size = New-Object System.Drawing.Size(45, 22)
    $lblSite.TextAlign = 'MiddleLeft'
    $lblSite.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblSite)

    $cmbSite = New-Object System.Windows.Forms.ComboBox
    $cmbSite.Location = New-Object System.Drawing.Point(65, 15)
    $cmbSite.Size = New-Object System.Drawing.Size(645, 24)
    $cmbSite.DropDownStyle = 'DropDown'
    $cmbSite.AutoCompleteMode   = 'SuggestAppend'
    $cmbSite.AutoCompleteSource = 'ListItems'
    $cmbSite.Sorted = $false
    $cmbSite.Anchor = 'Top,Left,Right'
    $dlg.Controls.Add($cmbSite)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(720, 14)
    $btnRefresh.Size = New-Object System.Drawing.Size(125, 26)
    $btnRefresh.Anchor = 'Top,Right'
    $dlg.Controls.Add($btnRefresh)

    # ---- Presets ----
    $lblPreset = New-Object System.Windows.Forms.Label
    $lblPreset.Text = "Presets:"
    $lblPreset.Location = New-Object System.Drawing.Point(15, 52)
    $lblPreset.Size = New-Object System.Drawing.Size(45, 22)
    $lblPreset.TextAlign = 'MiddleLeft'
    $lblPreset.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblPreset)

    $presetDefs = @(
        @{ Name = "Yesterday"; From = -1;  To = -1 }
        @{ Name = "Today";     From =  0;  To =  0 }
        @{ Name = "Last 7d";   From = -6;  To =  0 }
        @{ Name = "Last 30d";  From = -29; To =  0 }
        @{ Name = "Last 90d";  From = -89; To =  0 }
    )
    $px = 65
    foreach ($p in $presetDefs) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $p.Name
        $b.Location = New-Object System.Drawing.Point($px, 50)
        $b.Size = New-Object System.Drawing.Size(80, 26)
        $b.Tag = $p
        $dlg.Controls.Add($b)
        $b.Add_Click({
            $cfg = $this.Tag
            $today = (Get-Date).Date
            $dtpFrom.Value = $today.AddDays($cfg.From)
            $dtpTo.Value   = $today.AddDays($cfg.To)
        })
        $px += 86
    }

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = "Reset"
    $btnReset.Location = New-Object System.Drawing.Point($px, 50)
    $btnReset.Size = New-Object System.Drawing.Size(80, 26)
    $dlg.Controls.Add($btnReset)

    # ---- Date range ----
    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = "From:"
    $lblFrom.Location = New-Object System.Drawing.Point(15, 91)
    $lblFrom.Size = New-Object System.Drawing.Size(45, 22)
    $lblFrom.TextAlign = 'MiddleLeft'
    $lblFrom.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblFrom)

    $dtpFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtpFrom.Location = New-Object System.Drawing.Point(65, 88)
    $dtpFrom.Size = New-Object System.Drawing.Size(160, 24)
    $dtpFrom.Format = 'Custom'
    $dtpFrom.CustomFormat = 'yyyy-MM-dd'
    $dtpFrom.ShowUpDown = $true
    $dtpFrom.Value = (Get-Date).Date.AddDays($settings.FromOffset)
    $dlg.Controls.Add($dtpFrom)

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = "To:"
    $lblTo.Location = New-Object System.Drawing.Point(240, 91)
    $lblTo.Size = New-Object System.Drawing.Size(30, 22)
    $lblTo.TextAlign = 'MiddleLeft'
    $lblTo.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblTo)

    $dtpTo = New-Object System.Windows.Forms.DateTimePicker
    $dtpTo.Location = New-Object System.Drawing.Point(275, 88)
    $dtpTo.Size = New-Object System.Drawing.Size(160, 24)
    $dtpTo.Format = 'Custom'
    $dtpTo.CustomFormat = 'yyyy-MM-dd'
    $dtpTo.ShowUpDown = $true
    $dtpTo.Value = (Get-Date).Date.AddDays($settings.ToOffset)
    $dlg.Controls.Add($dtpTo)

    $lblEstimate = New-Object System.Windows.Forms.Label
    $lblEstimate.Location = New-Object System.Drawing.Point(450, 91)
    $lblEstimate.Size = New-Object System.Drawing.Size(395, 22)
    $lblEstimate.TextAlign = 'MiddleLeft'
    $lblEstimate.Anchor = 'Top,Left,Right'
    $lblEstimate.ForeColor = [System.Drawing.Color]::DimGray
    $lblEstimate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $dlg.Controls.Add($lblEstimate)

    # ---- Output folder ----
    $lblOut = New-Object System.Windows.Forms.Label
    $lblOut.Text = "Save to:"
    $lblOut.Location = New-Object System.Drawing.Point(15, 129)
    $lblOut.Size = New-Object System.Drawing.Size(55, 22)
    $lblOut.TextAlign = 'MiddleLeft'
    $lblOut.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblOut)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(70, 126)
    $txtOut.Size = New-Object System.Drawing.Size(655, 24)
    $txtOut.Text = if ($settings.OutputDir -and (Test-Path $settings.OutputDir)) {
        $settings.OutputDir
    } else {
        $script:DefaultDownloads
    }
    $txtOut.Anchor = 'Top,Left,Right'
    $dlg.Controls.Add($txtOut)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse…"
    $btnBrowse.Location = New-Object System.Drawing.Point(730, 125)
    $btnBrowse.Size = New-Object System.Drawing.Size(70, 26)
    $btnBrowse.Anchor = 'Top,Right'
    $dlg.Controls.Add($btnBrowse)

    $btnOpenFolder = New-Object System.Windows.Forms.Button
    $btnOpenFolder.Text = "Open"
    $btnOpenFolder.Location = New-Object System.Drawing.Point(805, 125)
    $btnOpenFolder.Size = New-Object System.Drawing.Size(40, 26)
    $btnOpenFolder.Anchor = 'Top,Right'
    $dlg.Controls.Add($btnOpenFolder)

    # ---- Options row 1: overwrite / remember ----
    $chkOverwrite = New-Object System.Windows.Forms.CheckBox
    $chkOverwrite.Text = "Overwrite without asking"
    $chkOverwrite.Location = New-Object System.Drawing.Point(70, 158)
    $chkOverwrite.Size = New-Object System.Drawing.Size(200, 22)
    $chkOverwrite.Checked = $false
    $dlg.Controls.Add($chkOverwrite)

    $chkRemember = New-Object System.Windows.Forms.CheckBox
    $chkRemember.Text = "Remember date range"
    $chkRemember.Location = New-Object System.Drawing.Point(280, 158)
    $chkRemember.Size = New-Object System.Drawing.Size(200, 22)
    $chkRemember.Checked = [bool]$settings.RememberRange
    $dlg.Controls.Add($chkRemember)

    # ---- Options row 2: mapping + scope ----
    $chkMapping = New-Object System.Windows.Forms.CheckBox
    $chkMapping.Text = "Also fetch machine mapping CSV"
    $chkMapping.Location = New-Object System.Drawing.Point(70, 180)
    $chkMapping.Size = New-Object System.Drawing.Size(260, 22)
    $chkMapping.Checked = [bool]$settings.FetchMapping
    $dlg.Controls.Add($chkMapping)

    $lblScope = New-Object System.Windows.Forms.Label
    $lblScope.Text = "scope:"
    $lblScope.Location = New-Object System.Drawing.Point(335, 180)
    $lblScope.Size = New-Object System.Drawing.Size(45, 22)
    $lblScope.TextAlign = 'MiddleLeft'
    $lblScope.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblScope)

    $cmbMapScope = New-Object System.Windows.Forms.ComboBox
    $cmbMapScope.Location = New-Object System.Drawing.Point(380, 178)
    $cmbMapScope.Size = New-Object System.Drawing.Size(300, 24)
    $cmbMapScope.DropDownStyle = 'DropDownList'
    [void]$cmbMapScope.Items.Add("This site only")
    [void]$cmbMapScope.Items.Add("ALL sites (parent + children)")
    $cmbMapScope.SelectedIndex = if ($settings.MappingScope -ne $null -and
                                     $settings.MappingScope -ge 0 -and
                                     $settings.MappingScope -lt 2) {
        [int]$settings.MappingScope
    } else { 1 }
    $dlg.Controls.Add($cmbMapScope)

    # ---- Log panel ----
    $logBox = New-Object System.Windows.Forms.RichTextBox
    $logBox.Location = New-Object System.Drawing.Point(15, 212)
    $logBox.Size = New-Object System.Drawing.Size(830, 398)
    $logBox.Anchor = 'Top,Bottom,Left,Right'
    $logBox.ReadOnly = $true
    $logBox.BackColor = [System.Drawing.Color]::FromArgb(252,252,252)
    $logBox.ForeColor = [System.Drawing.Color]::FromArgb(50,50,50)
    $logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $logBox.WordWrap = $false
    $logBox.ScrollBars = 'Both'
    $logBox.DetectUrls = $false
    $logBox.BorderStyle = 'FixedSingle'
    $dlg.Controls.Add($logBox)
    $script:LogBox = $logBox

    # ---- Progress bar ----
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(15, 622)
    $progressBar.Size = New-Object System.Drawing.Size(830, 18)
    $progressBar.Anchor = 'Bottom,Left,Right'
    $progressBar.Style = 'Continuous'
    $progressBar.Value = 0
    $dlg.Controls.Add($progressBar)

    # ---- Status ----
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(15, 648)
    $lblStatus.Size = New-Object System.Drawing.Size(560, 20)
    $lblStatus.Anchor = 'Bottom,Left,Right'
    $lblStatus.Text = "Ready."
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblStatus)

    # ---- Buttons ----
    $btnDownload = New-Object System.Windows.Forms.Button
    $btnDownload.Text = "Download"
    $btnDownload.Location = New-Object System.Drawing.Point(575, 645)
    $btnDownload.Size = New-Object System.Drawing.Size(130, 40)
    $btnDownload.Anchor = 'Bottom,Right'
    $btnDownload.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnDownload.FlatStyle = 'Flat'
    $btnDownload.FlatAppearance.BorderSize = 0
    $btnDownload.BackColor = [System.Drawing.Color]::FromArgb(52,152,219)
    $btnDownload.ForeColor = [System.Drawing.Color]::White
    $btnDownload.Cursor = 'Hand'
    $dlg.Controls.Add($btnDownload)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(715, 645)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 40)
    $btnCancel.Anchor = 'Bottom,Right'
    $btnCancel.Enabled = $false
    $dlg.Controls.Add($btnCancel)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(805, 645)
    $btnClose.Size = New-Object System.Drawing.Size(40, 40)
    $btnClose.Anchor = 'Bottom,Right'
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnClose)
    $dlg.CancelButton = $btnClose

    # ---- Helpers ----
    $updateEstimate = {
        $from = $dtpFrom.Value.Date
        $to   = $dtpTo.Value.Date
        if ($to -lt $from) {
            $lblEstimate.Text = "invalid range"
            $lblEstimate.ForeColor = [System.Drawing.Color]::Firebrick
            return
        }
        $days   = ($to - $from).Days + 1
        $chunks = [int][Math]::Ceiling($days / $script:ChunkSize)
        if ($days -le $script:ChunkSize) {
            $lblEstimate.Text = "$days day(s) · single request"
        } else {
            $estSec = $chunks * 40
            $mins   = [Math]::Max(1, [int][Math]::Ceiling($estSec / 60))
            $lblEstimate.Text = "$days day(s) · $chunks chunks · ~$mins min"
        }
        $lblEstimate.ForeColor = [System.Drawing.Color]::DimGray
    }

    $doLoadSites = {
        $btnRefresh.Enabled = $false
        $lblStatus.Text = "Loading site list from WebEOR..."
        $lblStatus.ForeColor = [System.Drawing.Color]::Black
        try {
            $sites = Get-WebEorSiteList
            $cmbSite.BeginUpdate()
            try {
                $cmbSite.Items.Clear()
                foreach ($s in $sites) {
                    [void]$cmbSite.Items.Add(("{0}  ({1})" -f $s.Name, $s.Id))
                }
            } finally {
                $cmbSite.EndUpdate()
            }

            $restored = $false
            if ($settings.SiteId) {
                for ($i = 0; $i -lt $cmbSite.Items.Count; $i++) {
                    if ($cmbSite.Items[$i] -match '\((\d{5})\)\s*$' -and
                        $Matches[1] -eq [string]$settings.SiteId) {
                        $cmbSite.SelectedIndex = $i
                        $restored = $true
                        break
                    }
                }
            }
            if (-not $restored -and $cmbSite.Items.Count -gt 0) {
                $cmbSite.SelectedIndex = 0
            }

            $lblStatus.Text = "$($sites.Count) sites loaded (A→Z). Ready."
            $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
        } catch {
            $lblStatus.Text = "Could not load site list — see log."
            $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
            Write-DlLog "Site list failed: $($_.Exception.Message)" -Level Error
            Write-DlLog "Check you are on the USPS intranet and Windows auth works." -Level Warn
        } finally {
            $btnRefresh.Enabled = $true
            & $updateEstimate
        }
    }

    # ---- Wire up ----
    $btnRefresh.Add_Click($doLoadSites)

    $btnBrowse.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Choose output folder"
        $fbd.SelectedPath = $txtOut.Text
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOut.Text = $fbd.SelectedPath
        }
    })

    $btnOpenFolder.Add_Click({
        $folder = $txtOut.Text
        if (Test-Path $folder) { Start-Process explorer.exe -ArgumentList "`"$folder`"" }
        else { [void][System.Windows.Forms.MessageBox]::Show(
            "Folder not found: $folder", "Open Folder", 'OK', 'Warning') }
    })

    $btnReset.Add_Click({
        $today = (Get-Date).Date
        $dtpFrom.Value = $today.AddDays(-1)
        $dtpTo.Value   = $today.AddDays(-1)
    })

    $dtpFrom.Add_ValueChanged($updateEstimate)
    $dtpTo.Add_ValueChanged($updateEstimate)

    $btnCancel.Add_Click({
        $script:CancelRequested = $true
        $btnCancel.Enabled = $false
        $lblStatus.Text = "Cancelling — will stop after the current request..."
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkGoldenrod
        Write-DlLog "Cancel requested (waiting for in-flight HTTP call to finish)" -Level Warn
    })

    $btnDownload.Add_Click({
        # --- Validate site ---
        $siteText = $cmbSite.Text
        if ($siteText -notmatch '^(.+?)\s+\((\d{5})\)\s*$') {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Pick a site from the dropdown first.", "No Site", 'OK', 'Warning')
            return
        }
        $siteName = $Matches[1]
        $sid      = [int]$Matches[2]

        $from = $dtpFrom.Value.Date
        $to   = $dtpTo.Value.Date
        if ($to -lt $from) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "'To' must be on or after 'From'.", "Bad Range", 'OK', 'Warning')
            return
        }

        $outDir = $txtOut.Text.Trim()
        if (-not (Test-Path $outDir)) {
            $mkdir = [System.Windows.Forms.MessageBox]::Show(
                "Output folder does not exist:`r`n$outDir`r`n`r`nCreate it?",
                "Create Folder", 'YesNo', 'Question')
            if ($mkdir -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            try {
                New-Item -ItemType Directory -Force -Path $outDir | Out-Null
            } catch {
                [void][System.Windows.Forms.MessageBox]::Show(
                    "Could not create folder: $_", "Error", 'OK', 'Error')
                return
            }
        }

        $eorFile = Join-Path $outDir ("EOR_{0}_{1}_{2}.csv" -f `
            $sid, $from.ToString('yyyyMMdd'), $to.ToString('yyyyMMdd'))

        if ((Test-Path $eorFile) -and -not $chkOverwrite.Checked) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "File already exists:`r`n$eorFile`r`n`r`nOverwrite?",
                "File Exists", 'YesNo', 'Question')
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        # --- Persist settings ---
        $settings.SiteId        = $sid
        $settings.OutputDir     = $outDir
        $settings.RememberRange = $chkRemember.Checked
        $settings.FetchMapping  = $chkMapping.Checked
        $settings.MappingScope  = $cmbMapScope.SelectedIndex
        if ($chkRemember.Checked) {
            $today = (Get-Date).Date
            $settings.FromOffset = [int]($from - $today).TotalDays
            $settings.ToOffset   = [int]($to   - $today).TotalDays
        }
        Save-Settings $settings

        # --- Prepare UI ---
        $script:CancelRequested = $false
        $btnDownload.Enabled = $false
        $btnRefresh.Enabled  = $false
        $btnClose.Enabled    = $false
        $btnCancel.Enabled   = $true
        $cmbSite.Enabled     = $false
        $dtpFrom.Enabled     = $false
        $dtpTo.Enabled       = $false
        $txtOut.Enabled      = $false
        $btnBrowse.Enabled   = $false
        $btnOpenFolder.Enabled = $false
        $chkMapping.Enabled  = $false
        $cmbMapScope.Enabled = $false

        $days = ($to - $from).Days + 1
        if ($days -le $script:ChunkSize) {
            $progressBar.Style = 'Marquee'
            $progressBar.MarqueeAnimationSpeed = 30
        } else {
            $progressBar.Style = 'Continuous'
            $progressBar.Value = 0
        }

        $lblStatus.Text = "Downloading…"
        $lblStatus.ForeColor = [System.Drawing.Color]::Black

        $progressCb = {
            param($Percent, $Message)
            if ($progressBar.Style -eq 'Continuous') {
                if ($Percent -gt $progressBar.Value) { $progressBar.Value = [Math]::Min(100, $Percent) }
            }
            if ($Message) { $lblStatus.Text = $Message }
            [System.Windows.Forms.Application]::DoEvents()
        }

        try {
            if (-not $script:Session -or $script:SessionSite -ne $sid) {
                $script:Session     = Connect-WebEor -SiteId $sid
                $script:SessionSite = $sid
            }

            $f = Save-EorCsv -Session $script:Session -SiteId $sid `
                -From $from -To $to -OutFile $eorFile -Progress $progressCb

            # ---- Mapping (only if requested and not cancelled) ----
            if ($chkMapping.Checked -and -not $script:CancelRequested) {
                $mapScope = if ($cmbMapScope.SelectedIndex -eq 1) { "ALL" } else { [string]$sid }
                $mapFile  = Join-Path $outDir ("MachineMapping_{0}.csv" -f $mapScope)
                $mapLabel = if ($mapScope -eq "ALL") {
                    "ALL sites (parent + children)"
                } else {
                    "site $mapScope only"
                }

                Write-DlLog "Fetching machine mapping — $mapLabel" -Level Info
                $progressBar.Style = 'Marquee'
                $progressBar.MarqueeAnimationSpeed = 30
                $lblStatus.Text = "Fetching machine mapping ($mapScope)..."
                [System.Windows.Forms.Application]::DoEvents()

                try {
                    $m = Save-MachineMappingCsv -Session $script:Session `
                        -SiteId $mapScope -MachineType "ALL" -Status "Y" -OutFile $mapFile
                    Write-DlLog "Machine mapping saved: $($m.FullName) ($([math]::Round($m.Length/1KB,1)) KB)" -Level Success
                } catch {
                    Write-DlLog "Mapping download failed (EOR still saved): $($_.Exception.Message)" -Level Error
                }
            }

            $progressBar.Style = 'Continuous'
            $progressBar.Value = 100
            $lblStatus.Text = "Done — saved $([math]::Round($f.Length/1KB,1)) KB"
            $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen

            # Precompute — fixes the -f / argument-list precedence bug
            $successMsg = (
                "Downloaded:`r`n{0}`r`n`r`nSize: {1:N1} KB`r`n`r`n" +
                "Yes — open the file`r`nNo — open the containing folder`r`nCancel — do nothing"
            ) -f $f.FullName, ($f.Length/1KB)

            $ans = [System.Windows.Forms.MessageBox]::Show(
                $successMsg, "Download Complete", 'YesNoCancel', 'Information')

            switch ($ans) {
                'Yes'    { Start-Process $f.FullName }
                'No'     { Start-Process explorer.exe -ArgumentList "/select,`"$($f.FullName)`"" }
                default  { }
            }
        } catch {
            $progressBar.Style = 'Continuous'
            $progressBar.Value = 0

            if ($script:CancelRequested) {
                $lblStatus.Text = "Cancelled."
                $lblStatus.ForeColor = [System.Drawing.Color]::DarkGoldenrod
                Write-DlLog "Download cancelled by user" -Level Warn
            }
            elseif ((Test-Path $eorFile) -and (Get-Item $eorFile -ErrorAction SilentlyContinue).Length -gt 100) {
                $lblStatus.Text = "Saved (see log for post-download warning)."
                $lblStatus.ForeColor = [System.Drawing.Color]::DarkGoldenrod
                Write-DlLog "File saved, but a post-download step failed: $($_.Exception.Message)" -Level Warn
                [void][System.Windows.Forms.MessageBox]::Show(
                    "The file was saved successfully to:`r`n$eorFile`r`n`r`n" +
                    "But a post-download step encountered an error:`r`n$($_.Exception.Message)",
                    "Saved with warning", 'OK', 'Warning')
            }
            else {
                $lblStatus.Text = "Download failed — see log."
                $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
                Write-DlLog "Download failed: $($_.Exception.Message)" -Level Error
                [void][System.Windows.Forms.MessageBox]::Show(
                    "Download failed:`r`n`r`n$($_.Exception.Message)`r`n`r`nSee the log pane for details.",
                    "Error", 'OK', 'Error')
            }
        } finally {
            $btnDownload.Enabled = $true
            $btnRefresh.Enabled  = $true
            $btnClose.Enabled    = $true
            $btnCancel.Enabled   = $false
            $cmbSite.Enabled     = $true
            $dtpFrom.Enabled     = $true
            $dtpTo.Enabled       = $true
            $txtOut.Enabled      = $true
            $btnBrowse.Enabled   = $true
            $btnOpenFolder.Enabled = $true
            $chkMapping.Enabled  = $true
            $cmbMapScope.Enabled = $true
        }
    })

    # Ctrl+Enter on the date pickers → download
    $ckHandler = {
        param($sender, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $e.SuppressKeyPress = $true
            if ($btnDownload.Enabled) { $btnDownload.PerformClick() }
        }
    }
    $dtpFrom.Add_KeyDown($ckHandler)
    $dtpTo.Add_KeyDown($ckHandler)

    $dlg.Add_Shown({
        & $updateEstimate
        & $doLoadSites
    })
    $dlg.Add_FormClosed({ $script:LogBox = $null })

    Write-DlLog "Ready. Choose a site and date range, then click Download." -Level Info
    [void]$dlg.ShowDialog()
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
Show-WebEorDownloaderDialog
