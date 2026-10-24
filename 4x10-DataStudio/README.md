# 4x10 Data Studio

Turns a **NMARS-WebEOR-DataGrabber** run folder into a finished 4/10 Modified Work
Week presentation — for any site, 100% offline (no internet, no installs).

## Network run — do this from the intranet (one press)

The grabber now makes the actual NMARS + WebEOR requests and logs everything:

```powershell
# 1) prove connectivity first (fast, safe)
powershell -File NMARS-WebEOR-DataGrabber.ps1 -NoGui -SiteId 19010 -TestConnection

# 2) full run — every report for one site, diagnostics bundle, studios open after
powershell -File NMARS-WebEOR-DataGrabber.ps1 -NoGui -SiteId 19010 -Reports ALL -Fy 2026 -Quarter 4 -Bundle -OpenStudio
```

(Or double-click the script for the GUI and click Run — same logging.)

Each run folder contains: `run_log.txt` (human), `run_log.jsonl` (structured,
one JSON per event: every request, status, bytes, ms, retries, validation
reasons), `samples\` (raw snippets of anything that failed), `DIAGNOSTICS.md`,
`HOW_TO_REPORT.txt`, and a `*_bundle.zip` when `-Bundle` is used.

**If something fails or parses wrong: send back the zip (or the three files
named in HOW_TO_REPORT.txt)** and the parsers get fixed.

## Workflow (3 steps)

1. **Pull the data** — run `NMARS-WebEOR-DataGrabber.ps1` on the intranet
   (one press: site + FY/AP range).
2. **Open `index.html`** (double-click it — any modern browser).
   - Click **Load report folder** (or drag the `NMARS-WebEOR_yyyyMMdd_HHmmss` folder onto the page).
   - Type the site ID / FY / AP range for the slide headers.
   - Click **Build deck** → a 13-slide presentation appears:
     title, scorecard, PM completion, bypassed tasks, PM trends, contract cost,
     equipment & parts, production runs (down/idle/jam by tour), maintenance
     windows kept, run-plan-vs-actual efficiency, business pitch, union case,
     and the RFI list.
   - **Print / save PDF** → use the browser's "Save as PDF" printer for a
     shareable PDF deck.
3. **(Optional) Real PPTX** — click **Export data (JSON)**, then run
   `powershell -File Build-Presentation.ps1 -JsonFile 4x10_deck_data_xxx.json`
   → a native PowerPoint deck is saved next to the JSON (uses PowerPoint COM;
   works with any Office 2016+).

Also handy: **Copy scorecard** puts the key numbers on the clipboard for pasting
into the earlier PPTX template.

## What it recognizes

NMARS HTML reports (eCBM completion, bypassed tasks, PM trends, contract cost,
equipment inventory, parts issues, stock room, acronym listing), WebEOR CSVs
(production runs, maintenance runs, run-vs-RPG-plan, machine mapping — with the
MODS="M" split and MODS-day convention), and the LaborChecklist toolset:
**Machine Staffing Report HTML** (MM7/MPE9/ET10 staffing vs calculated),
**\*-Staffing-Table.csv / \*-Calculated-Staffing-Table.csv** (workload hrs/yr,
tours/day), **Staffing Calculator session JSON** (staffing-calculator-session-v1:
eWHEP claimed vs MMO required), and the **eMARS labor-transaction export**
(hours, work orders, active days, levels). Files are classified by name and by
content sniffing.

## Honest caveats

- NMARS report layouts vary by plant/version — the parsers match on column
  headers and fall back gracefully. If a metric shows "—" or looks off, open the
  file list tab and eyeball that report; adjust the parsing thresholds in the
  script's CORE block if needed.
- The WebEOR date/tour columns are detected by header name, so keep the default
  CSV column sets from the grabber.
- The in-browser deck is the presentation; `Build-Presentation.ps1` gives you the
  same deck as a real .pptx.

## Coverage Studio (people vs machines)

`Coverage-Studio.html` answers "who was providing coverage, when — and was it enough?"

1. Run `Split-TacsReport.ps1 "C:\TACS\Employee_Everything.csv"` (on the intranet)
   → produces `.Base.csv`, `.Summary.csv`, `.Transactions.csv`.
2. Open `Coverage-Studio.html` and drop:
   - the **\*.Transactions.csv** (clock rings — required),
   - the **\*.Base.csv** (schedules + LDC/level — enables craft grouping),
   - your **EOR run CSVs** (machines + pieces).
3. Click **Build chart**.

You get:
- a **step chart** of people clocked in (lunch subtracted, overnight tours handled)
  overlaid with **machines running** and **pieces per 15 minutes**;
- group-by: Everyone / Maintenance only (LDC 35xx–39xx) / by craft (level 7/9/10 → MM7/MPE9/ET10);
- a **required-vs-actual coverage model** (tune the ratio slider): per bucket,
  required = ⌈machines running × ratio⌉ — red bars show deficit hours;
- a per-hour table with people, machines, pieces, maintenance runs, required and gap.

That model is the prototype of "how many people of which type, and when" — once
you settle the machine-per-person ratio (from MS-47/eCBM workload data), the
numbers it produces are the staffing schedule.

## Files

- `index.html` — the studio (self-contained, offline).
- `Build-Presentation.ps1` — JSON → PPTX (PowerPoint COM).
- `Coverage-Studio.html` — TACS coverage vs machine activity overlay.
- `Split-TacsReport.ps1` — TACS "Employee Everything" → normalized Base/Summary/Transactions CSVs.
- `Sample-Output-DEMO-DATA.pptx` — example deck built from *synthetic* numbers
  (demo only, do not present as real data).
