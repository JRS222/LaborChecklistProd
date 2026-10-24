#Requires -Version 5.1
<#
.SYNOPSIS
    Build-Presentation.ps1 — converts the 4x10 Data Studio JSON export into a
    real PowerPoint (.pptx) deck using PowerPoint COM automation (Office).

.DESCRIPTION
    Companion to 4x10 Data Studio (index.html).  In the webpage click
    "Export data (JSON)", then run:

        powershell -File Build-Presentation.ps1 -JsonFile 4x10_deck_data_xxx.json

    The .pptx is written next to the JSON file.

.PARAMETER JsonFile
    Path to the JSON exported by the Data Studio webpage.

.PARAMETER OutPptx
    Optional output path.  Defaults to <JsonFile>.pptx.

.NOTES
    Requires PowerPoint (any Office 2016+).  Run under Windows PowerShell 5.1.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$JsonFile,
    [string]$OutPptx
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $JsonFile)) { throw "JSON file not found: $JsonFile" }
$j = Get-Content $JsonFile -Raw | ConvertFrom-Json
if (-not $OutPptx) { $OutPptx = [System.IO.Path]::ChangeExtension($JsonFile, '.pptx') }

function Rgb([int]$r, [int]$g, [int]$b) { return ($r + ($g * 256) + ($b * 65536)) }
$Navy   = Rgb 31 56 100
$Navy2  = Rgb 46 78 140
$Red    = Rgb 200 16 46
$Ink    = Rgb 33 37 43
$Gray   = Rgb 90 96 106
$Light  = Rgb 237 241 248
$White  = Rgb 255 255 255
$Green  = Rgb 30 122 70
$Orange = Rgb 178 106 0
$Palette = @($Navy2, $Red, $Green, $Orange, (Rgb 122 132 148))

$pp = $null; $pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Add($false)   # hidden window
    $pres.PageSetup.SlideWidth  = 960        # 13.333 in
    $pres.PageSetup.SlideHeight = 540        # 7.5 in
    $blank = 12                              # ppLayoutBlank

    $slideNo = 0
    function New-Slide {
        $script:slideNo++
        return $pres.Slides.Add($script:slideNo, $blank)
    }
    function Add-Rect($slide, $l, $t, $w, $h, $color) {
        $sh = $slide.Shapes.AddShape(1, $l, $t, $w, $h)   # msoShapeRectangle
        $sh.Fill.ForeColor.RGB = $color
        $sh.Line.Visible = $false
        $sh.Shadow.Visible = $false
        return $sh
    }
    function Add-Text($slide, $l, $t, $w, $h, $text, $size, $bold, $color, $align) {
        $sh = $slide.Shapes.AddTextbox(1, $l, $t, $w, $h)
        $tf = $sh.TextFrame
        $tf.WordWrap = -1   # msoTrue
        $tf.TextRange.Text = $text
        $tf.TextRange.Font.Size = $size
        $tf.TextRange.Font.Bold = $bold
        $tf.TextRange.Font.Color.RGB = $color
        $tf.TextRange.Font.Name = "Calibri"
        if ($align) { $tf.TextRange.ParagraphFormat.Alignment = $align }
        return $sh
    }
    function Add-Header($slide, $kicker, $title) {
        Add-Rect $slide 0 0 960 64 $Navy | Out-Null
        Add-Rect $slide 0 64 960 4 $Red | Out-Null
        Add-Text $slide 28 8 900 16 $kicker.ToUpper() 11 $true (Rgb 242 184 195) 1 | Out-Null
        Add-Text $slide 28 24 900 34 $title 24 $true $White 1 | Out-Null
    }
    function Add-Footer($slide) {
        $f = "{0}  |  {1}  |  4/10 Modified Work Week proposal  |  generated {2}" -f $j.site, $j.window, $j.generated
        Add-Text $slide 28 516 900 14 $f 10 $false $Gray 1 | Out-Null
    }
    function Add-Bars($slide, $labels, $values, $l, $t, $w, $h, $colorIdx, $caption) {
        if (-not $labels -or $labels.Count -eq 0) {
            Add-Text $slide $l ($t + 20) $w 30 "no data for this chart" 12 $false $Gray 1 | Out-Null
            return
        }
        $n = $labels.Count
        $max = 1
        foreach ($v in $values) { if ($v -gt $max) { $max = $v } }
        if ($caption) { Add-Text $slide $l ($t - 18) $w 16 $caption 11 $true $Gray 1 | Out-Null }
        $chartL = $l + 6; $chartW = $w - 14
        $barAreaW = $chartW / $n
        $barW = [math]::Min(90, $barAreaW * 0.6)
        for ($i = 0; $i -lt $n; $i++) {
            $x = $chartL + ($barAreaW * $i) + (($barAreaW - $barW) / 2)
            $vh = 0
            if ($max -gt 0) { $vh = ($values[$i] / $max) * ($h - 26) }
            $y = $t + ($h - 26 - $vh)
            Add-Rect $slide $x $y $barW $vh $Palette[$colorIdx] | Out-Null
            $lab = [string]$labels[$i]; if ($lab.Length -gt 11) { $lab = $lab.Substring(0, 10) + "…" }
            Add-Text $slide ($x - $barAreaW * 0.2) ($t + $h - 24) ($barAreaW * 1.4) 16 $lab 9 $false $Ink 2 | Out-Null
        }
    }
    function Add-GroupedBars($slide, $labels, $series, $l, $t, $w, $h, $caption) {
        if (-not $labels -or $labels.Count -eq 0) {
            Add-Text $slide $l ($t + 20) $w 30 "no data for this chart" 12 $false $Gray 1 | Out-Null
            return
        }
        $n = $labels.Count; $k = $series.Count
        $max = 1
        foreach ($s in $series) { foreach ($v in $s.data) { if ($v -gt $max) { $max = $v } } }
        if ($caption) { Add-Text $slide $l ($t - 18) $w 16 $caption 11 $true $Gray 1 | Out-Null }
        $chartL = $l + 6; $chartW = $w - 14
        $barAreaW = $chartW / $n
        $slot = [math]::Min(60, $barAreaW * 0.7 / $k)
        for ($i = 0; $i -lt $n; $i++) {
            for ($s2 = 0; $s2 -lt $k; $s2++) {
                $v = $series[$s2].data[$i]
                $x = $chartL + ($barAreaW * $i) + (($barAreaW - ($slot * $k)) / 2) + ($slot * $s2)
                $vh = 0
                if ($max -gt 0) { $vh = ($v / $max) * ($h - 26) }
                $y = $t + ($h - 26 - $vh)
                Add-Rect $slide $x $y $slot $vh $Palette[$s2] | Out-Null
            }
            $lab = [string]$labels[$i]; if ($lab.Length -gt 11) { $lab = $lab.Substring(0, 10) + "…" }
            Add-Text $slide ($chartL + $barAreaW * $i) ($t + $h - 24) $barAreaW 16 $lab 9 $false $Ink 2 | Out-Null
        }
        $leg = ($series | ForEach-Object { $_.name }) -join "  |  "
        Add-Text $slide $chartL ($t + $h - 4) $chartW 12 $leg 9 $true $Gray 1 | Out-Null
    }

    $sc = $j.scorecard
    $ch = $j.charts

    # ---- 1. Title ----
    $s = New-Slide
    Add-Rect $s 0 0 960 540 $Navy | Out-Null
    Add-Rect $s 60 80 5 330 $Red | Out-Null
    Add-Text $s 90 88 840 24 "APWU–USPS NATIONAL AGREEMENT  •  MAINTENANCE CRAFT PROPOSAL" 13 $true (Rgb 242 184 195) 1 | Out-Null
    Add-Text $s 90 120 840 70 "4/10 Modified Work Week" 46 $true $White 1 | Out-Null
    Add-Text $s 90 200 840 34 "Four 10-hour days for full-time regular Maintenance employees" 20 $false (Rgb 220 228 245) 1 | Out-Null
    Add-Text $s 90 260 820 120 ("Site {0}  •  {1}`r`rA win for the Postal Service: same 40 hours, more PMs completed in straight time, less overtime and contractor spend.`rA win for the Union: every protection intact, 52 extra days off a year — authorized by the national Modified Work Week MOU." -f $j.site, $j.window) 15 $false (Rgb 232 237 248) 1 | Out-Null
    Add-Rect $s 0 506 960 34 $Red | Out-Null
    Add-Text $s 28 512 900 24 ("Generated {0} — data: NMARS + WebEOR" -f $j.generated) 11 $false $White 1 | Out-Null

    # ---- 2. Scorecard ----
    $s = New-Slide
    Add-Header $s "SCORECARD" "The numbers, on one slide"
    $kpis = @(
        @("eCBM PM completion", $sc.pmCompletion),
        @("Tasks bypassed", $sc.bypassed),
        @("Contract cost", $sc.contractCost),
        @("Down-time share", $sc.downTimeShare),
        @("Maintenance runs", $sc.maintRuns)
    )
    $kw = 178
    for ($i = 0; $i -lt 5; $i++) {
        $x = 28 + ($i * ($kw + 6))
        Add-Rect $s $x 92 $kw 92 $Light | Out-Null
        Add-Rect $s $x 92 $kw 4 $Navy | Out-Null
        Add-Text $s ($x + 10) 104 ($kw - 16) 30 $kpis[$i][1] 22 $true $Navy 1 | Out-Null
        Add-Text $s ($x + 10) 142 ($kw - 16) 34 $kpis[$i][0] 11 $false $Gray 1 | Out-Null
    }
    $rows = 11; $cols = 2
    $tbl = $s.Shapes.AddTable($rows, $cols, 28, 206, 904, 290)
    $tbl.Table.Columns.Item(1).Width = 620
    $tbl.Table.Columns.Item(2).Width = 284
    $items = @(
        @("eCBM PM completion (site avg)", $sc.pmCompletion),
        @("PM tasks bypassed (period)", $sc.bypassed),
        @("Contract maintenance cost (period)", $sc.contractCost),
        @("Production down-time share", $sc.downTimeShare),
        @("Maintenance runs (windows kept)", $sc.maintRuns),
        @("Avg maintenance window (min)", $sc.maintAvgMin),
        @("Plan-vs-actual run efficiency", $sc.planVsActual),
        @("Jams per 10k pieces fed", $sc.jamsPer10k),
        @("Stock room parts on hand", $sc.stockParts),
        @("Equipment records", $sc.equipment),
        @("Business case", "Same 40 hours = zero added straight-time labor cost")
    )
    for ($r = 0; $r -lt $rows; $r++) {
        $c1 = $tbl.Table.Cell($r + 1, 1).Shape.TextFrame.TextRange
        $c2 = $tbl.Table.Cell($r + 1, 2).Shape.TextFrame.TextRange
        $c1.Text = $items[$r][0]; $c2.Text = $items[$r][1]
        $c1.Font.Size = 12; $c2.Font.Size = 12; $c2.Font.Bold = $true
        $c1.Font.Color.RGB = $Ink; $c2.Font.Color.RGB = $Navy
        if ($r % 2 -eq 0) {
            $tbl.Table.Cell($r + 1, 1).Shape.Fill.ForeColor.RGB = $Light
            $tbl.Table.Cell($r + 1, 2).Shape.Fill.ForeColor.RGB = $Light
        }
    }
    Add-Footer $s

    # ---- 3. PM completion ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Preventive maintenance actually gets done"
    Add-Text $s 28 96 300 40 ("Avg site completion: " + $sc.pmCompletion) 18 $true $Navy 1 | Out-Null
    Add-Text $s 340 96 400 40 ("Rows below 100%: " + $sc.pmBelow100 + "   —   every gap is PM not finished") 13 $false $Gray 1 | Out-Null
    Add-Bars $s $ch.completion.labels $ch.completion.values 40 160 880 300 0 "Completion % by row (top rows)"
    Add-Text $s 40 470 880 34 "A 10-hour shift gives one uninterrupted window: LOTO, setup, and post-dispatch machine access happen once — not twice." 13 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 4. Bypassed ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Bypassed PMs are deferred risk"
    Add-Text $s 28 96 400 40 ("eCBM tasks bypassed: " + $sc.bypassed) 18 $true $Red 1 | Out-Null
    if ($ch.bypassMonthly.labels.Count -gt 0) {
        Add-Bars $s $ch.bypassMonthly.labels $ch.bypassMonthly.values 40 160 880 300 1 "Bypassed tasks by month"
    } elseif ($ch.bypassAcronym.labels.Count -gt 0) {
        Add-Bars $s $ch.bypassAcronym.labels $ch.bypassAcronym.values 40 160 880 300 1 "Bypassed tasks per machine acronym"
    }
    Add-Text $s 40 470 880 34 "Each bypassed task is deferred risk. A 10-hour window finishes jobs that currently can't be completed without taking a machine down mid-shift." 13 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 5. Trends ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "The trend line before the pilot"
    if ($ch.trends) {
        $tlabels = $ch.trends.labels
        $s0 = $ch.trends.series[0]
        Add-Bars $s $tlabels $s0.data 40 170 880 280 0 ("Trend: " + $s0.name)
        $names = ($ch.trends.series | ForEach-Object { $_.name }) -join ", "
        Add-Text $s 40 460 880 24 "Series in report: $names" 11 $false $Gray 1 | Out-Null
    } else {
        Add-Text $s 40 160 880 40 "No trend series detected — check that TRENDS_* files loaded." 14 $false $Gray 1 | Out-Null
    }
    Add-Footer $s

    # ---- 6. Contract cost ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Less overtime, less contractor spend"
    Add-Text $s 28 96 500 40 ("Contract maintenance cost (period): " + $sc.contractCost) 18 $true $Navy 1 | Out-Null
    Add-Bars $s $ch.contract.labels $ch.contract.values 40 170 880 300 2 "Contract cost by scope"
    Add-Text $s 40 480 880 24 "Every contractor call-out absorbed in-house at straight time is hard-dollar savings." 13 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 7. Equipment ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "The machine and parts picture"
    Add-Text $s 28 96 500 40 ("Equipment records: " + $sc.equipment + "   |   Stock room parts: " + $sc.stockParts) 16 $true $Navy 1 | Out-Null
    Add-Bars $s $ch.equipCondition.labels $ch.equipCondition.values 40 170 880 300 3 "Equipment by condition code"
    Add-Text $s 40 480 880 24 "Aging machines (conditions C/D/E) carry the heaviest PM load — the machines a 10-hour crew keeps on schedule." 13 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 8. Production runs ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Production runs: where the machine time goes"
    Add-Text $s 28 96 900 34 ("Down-time share of run time: " + $sc.downTimeShare + "   |   Jams per 10k fed: " + $sc.jamsPer10k + "   (maintenance runs excluded — they do not drive production staffing)") 13 $false $Gray 1 | Out-Null
    if ($ch.prodTour) {
        Add-GroupedBars $s $ch.prodTour.labels $ch.prodTour.series 40 150 880 330 "Minutes by tour (production runs only)"
    }
    Add-Footer $s

    # ---- 9. Staffing & workload ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Staffing & workload: the labor math"
    $st = $j.staffing
    if ($st.session) {
        $mg2 = $st.session.mgmt; $cl2 = $st.session.calc
        $mgt = [double]$mg2.mm7 + [double]$mg2.mpe9 + [double]$mg2.et10
        Add-Text $s 28 96 240 40 ("Machines: " + $st.session.machines) 14 $true $Navy 1 | Out-Null
        Add-Text $s 280 96 250 40 ("eWHEP claimed: " + [math]::Round($mgt)) 14 $true $Navy 1 | Out-Null
        Add-Text $s 540 96 250 40 ("MMO required: " + [math]::Round($cl2.total)) 14 $true $Navy 1 | Out-Null
        Add-Text $s 800 96 160 40 ("Maint hrs/yr: " + [math]::Round($cl2.maint)) 14 $true $Red 1 | Out-Null
    } elseif ($st.csv) {
        Add-Text $s 28 96 240 40 ("Machines: " + $st.csv.machines) 14 $true $Navy 1 | Out-Null
        Add-Text $s 280 96 300 40 ("Total maint hrs/yr: " + [math]::Round($st.csv.total)) 14 $true $Navy 1 | Out-Null
        Add-Text $s 590 96 300 40 ("Operational maint: " + [math]::Round($st.csv.maint)) 14 $true $Red 1 | Out-Null
    } elseif ($st.report) {
        $rp = $st.report
        Add-Text $s 28 96 240 40 ("Machines: " + $rp.machines) 14 $true $Navy 1 | Out-Null
        Add-Text $s 280 96 300 40 ("Staffing hrs: " + [math]::Round([double]$rp.mm7.staff + [double]$rp.mpe9.staff + [double]$rp.et10.staff)) 14 $true $Navy 1 | Out-Null
        Add-Text $s 590 96 300 40 ("Calculated hrs: " + [math]::Round([double]$rp.mm7.calc + [double]$rp.mpe9.calc + [double]$rp.et10.calc)) 14 $true $Red 1 | Out-Null
    } else {
        Add-Text $s 40 120 880 60 "No staffing data in this export - drop a Machine Staffing Report, staffing-table CSVs, or a Staffing Calculator session into the Data Studio." 14 $false $Gray 1 | Out-Null
    }
    if ($ch.staffing) {
        Add-GroupedBars $s $ch.staffing.labels $ch.staffing.series 40 170 880 300 "Claimed vs required by craft (hrs/yr)"
    }
    Add-Text $s 40 480 880 24 "Same total maintenance hours repackaged as 4x10 = same labor cost, 20% fewer commutes, 52 extra off-days - the workload math does not change." 13 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 10. Labor actuals ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Labor actuals: where the hours really go"
    $act = $st.actuals
    if ($act) {
        Add-Text $s 28 96 240 40 ("Actual hours: " + [math]::Round([double]$act.hours)) 14 $true $Navy 1 | Out-Null
        Add-Text $s 280 96 240 40 ("Work orders: " + $act.woCount) 14 $true $Navy 1 | Out-Null
        Add-Text $s 530 96 240 40 ("Active days: " + $act.activeDays) 14 $true $Navy 1 | Out-Null
        Add-Text $s 770 96 200 40 ("Transactions: " + $act.rows) 14 $true $Navy 1 | Out-Null
        if ($ch.laborByAcr) {
            Add-Bars $s $ch.laborByAcr.labels $ch.laborByAcr.values 40 170 880 300 1 "Actual hours by machine acronym (top)"
        }
        if ($st.session) {
            $req  = [double]$st.session.calc.total
            $acth = [double]$act.hours
            $claimed = [double]$st.session.mgmt.mm7 + [double]$st.session.mgmt.mpe9 + [double]$st.session.mgmt.et10
            $ratio = if ($req -gt 0) { [math]::Round($acth / $req * 100) } else { "—" }
            Add-Text $s 40 480 880 24 ("Three-way: eWHEP claimed {0} - MMO required {1} - eMARS actual {2} (~{3}% of required in the exported window)." -f [math]::Round($claimed), [math]::Round($req), [math]::Round($acth), $ratio) 13 $false $Ink 1 | Out-Null
        }
    } else {
        Add-Text $s 40 120 880 60 "No eMARS labor-transaction export in this dataset - export it (Completed Actions -> Print Daily Entries) and reload." 14 $false $Gray 1 | Out-Null
    }
    Add-Footer $s

    # ---- 11. Maintenance windows ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Are maintenance windows being kept?"
    Add-Text $s 28 96 600 34 ("Maintenance runs recorded: " + $sc.maintRuns + "   |   Avg window: " + $sc.maintAvgMin + " min") 16 $true $Green 1 | Out-Null
    if ($ch.maintTour) {
        Add-Bars $s $ch.maintTour.labels $ch.maintTour.values 40 170 880 300 2 "Maintenance windows by tour (minutes)"
    }
    Add-Text $s 40 480 880 24 "Compare actual windows vs the plan and PM due dates — a 10-hour shift matches the windows the plant already gives up." 13 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 10. Plan vs actual ----
    $s = New-Slide
    Add-Header $s "BUSINESS CASE — USPS" "Run plan vs actual runs"
    Add-Text $s 28 96 500 40 ("Plan-vs-actual efficiency: " + $sc.planVsActual) 18 $true $Navy 1 | Out-Null
    if ($ch.planActual) {
        Add-Bars $s $ch.planActual.labels $ch.planActual.values 40 170 880 300 0 "Plan vs actual (summed)"
    }
    if ($j.rpgNote) { Add-Text $s 40 470 880 22 $j.rpgNote 11 $false $Gray 1 | Out-Null }
    Add-Text $s 40 484 880 20 "Join on site, MODS date, machine type, operation, tour. Gaps where actual < plan are the efficiency a 4/10 recovers." 12 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 11. Business narrative ----
    $s = New-Slide
    Add-Header $s "THE PITCH" "Why this makes business sense"
    $bullets = @(
        "Same 40 hours, zero added labor cost.  4x10 = 40. Overtime only after 10 hours in pay status or 40/week; penalty rules and 56/60 limits unchanged; Sunday premium capped at 16 hrs/week (10/4 Guidelines).",
        "More PMs in straight time.  Current completion " + $sc.pmCompletion + " with " + $sc.bypassed + " bypassed tasks shows the gap.",
        "Less OT and contractor spend.  Contract cost this period: " + $sc.contractCost + ".  Production down time: " + $sc.downTimeShare + " of run time.",
        "Coverage and retention.  Staggered Mon-Thu / Tue-Fri crews give 6-day coverage at the same hours; a 3-day weekend recruits and retains maintenance skills at zero payroll cost."
    ) -join "`r`r"
    Add-Text $s 40 110 880 400 $bullets 15 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 12. Union case ----
    $s = New-Slide
    Add-Header $s "UNION CASE — APWU" "A win our members keep"
    $bullets = @(
        "Every protection intact.  88 holiday hours guaranteed; leave charged up to 10 hrs/day; OTDL rights unchanged; NS-day guarantee stays 8 hrs; training 5+ days reverts to 8/5 (CBA pp. 298-301).",
        "Voluntary and reversible.  Either side declines freely - no impasse, no grievance.  30-day written cancellation (CBA p. 297).",
        "No NTFT creep in Maintenance.  4/10 keeps every assignment full-time regular - posted and bid by seniority under Article 38.",
        "52 extra days off a year.  20% fewer commutes.  A real quality-of-life win with nothing traded away."
    ) -join "`r`r"
    Add-Text $s 40 110 880 400 $bullets 15 $false $Ink 1 | Out-Null
    Add-Footer $s

    # ---- 13. RFI ----
    $s = New-Slide
    Add-Header $s "NEXT STEP" "Request for Information — data that completes the case"
    $rfi = @(
        @("TACS overtime + penalty-OT hours and dollars by tour, pay period, occupational group", "The 8-10 window conversion to straight time"),
        @("Staffing complement vs vacancies (PS 4852) by occupational group", "Staffing leverage of 4/10; posting options"),
        @("eCBM/MS-47 PM workload hours due vs completed (detail export)", "PM backlog in hours - the pilot's target"),
        @("Contractor maintenance invoices (vendor, machine, work type)", "Per-machine contractor substitution savings"),
        @("MODS/EDW downtime extract (97/98 transactions)", "Independent cross-check of WebEOR down time"),
        @("Plant operating hours + utility data", "Window alignment for 10-hour tours"),
        @("Safety incident records (OSHA 300/301)", "Fatigue-objection baseline for the pilot"),
        @("Leave/absence usage by day of week (TACS)", "Quality-of-life / commute impact")
    )
    $rows = $rfi.Count + 1
    $tbl2 = $s.Shapes.AddTable($rows, 2, 28, 100, 904, 400)
    $tbl2.Table.Columns.Item(1).Width = 560
    $tbl2.Table.Columns.Item(2).Width = 344
    $h1 = $tbl2.Table.Cell(1, 1).Shape.TextFrame.TextRange
    $h2 = $tbl2.Table.Cell(1, 2).Shape.TextFrame.TextRange
    $h1.Text = "Request via RFI"; $h2.Text = "What it proves"
    $h1.Font.Bold = $true; $h2.Font.Bold = $true
    $h1.Font.Color.RGB = $White; $h2.Font.Color.RGB = $White
    $tbl2.Table.Cell(1, 1).Shape.Fill.ForeColor.RGB = $Navy
    $tbl2.Table.Cell(1, 2).Shape.Fill.ForeColor.RGB = $Navy
    for ($r = 0; $r -lt $rfi.Count; $r++) {
        $a = $tbl2.Table.Cell($r + 2, 1).Shape.TextFrame.TextRange
        $b = $tbl2.Table.Cell($r + 2, 2).Shape.TextFrame.TextRange
        $a.Text = $rfi[$r][0]; $b.Text = $rfi[$r][1]
        $a.Font.Size = 12; $b.Font.Size = 12
        $a.Font.Color.RGB = $Ink; $b.Font.Color.RGB = $Gray
    }
    Add-Footer $s

    $pres.SaveAs($OutPptx, 24)   # 24 = ppSaveAsOpenXMLPresentation
    Write-Host "Saved: $OutPptx ($($pres.Slides.Count) slides)"
}
finally {
    if ($pres) { try { $pres.Close() } catch { } }
    if ($pp)   { try { $pp.Quit() } catch { } }
    if ($pres) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null }
    if ($pp)   { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp)   | Out-Null }
}
