# EOR & Machine Mapping Analysis — Pittsburgh (site 19010)

**Files:** `EOR_19010_20251021_20261021.csv` (267,299 runs) + `MachineMapping_ALL.csv` (85 rows)
**Period:** 10/21/2025 – 10/21/2026 (full FY2026)
**Method notes:**
- MODS week convention applied: MODS day rolls at 23:00 (Friday 11 PM = Saturday), so day/month =
  calendar date of *Start + 1 hour*. Validated against the `MODS` column's "S MM/DD/YY" dates:
  **92.35% match** on 213,131 sent runs.
- Maintenance runs = `MODS = "M"` (53,377 of 267,299).
- This export is the 12-column set — it has **no down/idle/jam/maintenance-time columns**.
  "Idle" below is a proxy: non-running time inside each machine-day operating window
  (first start → last end), which includes staging and planning.
- One mapping row (an SDUS) is malformed in the source CSV (fields shifted); excluded from counts.

## 1. Scale

| | |
|---|---|
| Total runs (year) | 267,299 |
| Production runs | 213,922 (80.0%) |
| **Maintenance runs** | **53,377 (20.0%) — 1 in every 5 runs is maintenance** |
| Pieces fed (production) | 2,977,285,886 (~3.0 B) |
| Production run time | 210,098 hrs/year (~24 machine-years) |
| Machines with production runs | 70 |
| Machines with maintenance runs | 55 |
| Active machines in mapping | 84 (Pittsburgh P&DC, Irwin PO, Pleasant Hills S&DC) |
| Mapped machines with **no EOR runs all year** | 15 (PSS 2,3,4,10,11,500, SDUS 1&2, F/C 1&2, Manual Parcels, NEC Canceller, NGTC, EAA 2) |

## 2. Maintenance windows: being kept — frequently, but briefly

- **53,377 maintenance runs/year ≈ 146 per day; average window 6.3 minutes.**
- **349 maintenance-days per machine per year** — machines are taken down for maintenance
  nearly every single day.
- Steady all year: 4,100–4,700 maint runs and ~25k–32k maint-minutes per month. No seasonal cliff.
- **By MODS day of week** (runs): Mon 7,577 · Tue 8,109 · Wed 8,231 · Thu 7,788 · Fri 7,969 ·
  Sat 7,867 · **Sun 5,836** (lightest day, −29% vs Wednesday).
- **By clock hour**: heavy 00:00–05:00 (peak 05:00 = 8,290 runs), dip 06:00–08:00,
  moderate 09:00–12:00, evening ramp 22:00–23:00.
- **By tour**: Tour 1 (overnight) = 30,560 maint runs (57%) · Tour 2 = 17,854 · Tour 3 = 4,963.
- **By machine type** (runs / avg window): DBCS 32,039 @ 4.9 min · AFCS200 11,247 @ 8.6 ·
  DIOSS 5,770 @ 4.1 · **AFSM100 2,144 @ 23.3 min (the real PM jobs)** · CIOSS 1,735 @ 3.8 · LCREM 442 @ 5.7.

## 3. Production & the idle-time opportunity

- **By tour**: Tour 3 (evening) 101,177 runs / 1.47 B pieces / 92,548 hrs ·
  Tour 2 65,486 runs / 65,842 hrs · Tour 1 47,259 runs / 51,708 hrs.
- Machine-day operating window averages **17.9 hrs**; machines actually run **9.1 hrs**;
  **49.1% of the operating window is non-running time.**
- Non-running share by tour: **Tour 3 = 59.7%** · Tour 2 = 35.2% · Tour 1 = 29.0%.
- Throughput: DBCS 22.1k pcs/hr (2.495 B pieces — **84% of all volume**) · AFCS200 21.8k ·
  DIOSS 13.5k · CIOSS 12.5k · AFSM100 6.7k · PSS 477 · HSTS 720 · MSWYB 101.
- Production steady: 16.4k–19.3k runs/month (December peak).

## 4. What this means for the 4/10 proposal

1. **Windows are being kept — the problem is depth, not frequency.** ~146 take-downs a day at
   6.3 minutes average (vs 23 min for AFSM100's real PMs) means maintenance happens in short,
   reactive jogs. A 10-hour shift consolidates many short windows into fewer, deeper, planned
   ones — and completes the PMs the current schedule can only start.
2. **Machine access is not the constraint — scheduled labor is.** Machines sit non-running 49%
   of their operating day (Tour 3: 60%). There is an enormous maintenance window available;
   the 4/10 puts a full crew in front of it without overtime.
3. **Tour 1 is already the maintenance tour** (57% of maintenance runs). It is the natural
   4/10 pilot — an overnight compressed schedule fits the existing take-down pattern.
4. **The maintenance load is flat year-round.** Steady 4/10 staffing works; there is no
   peak-season reason to keep 5×8.
5. **1 in 5 runs produces zero pieces.** Maintenance runs are 20% of all machine activity —
   keeping them short, planned, and off the production staffing math (separate files, as the
   Data Studio now does) is itself an efficiency argument.
6. **15 mapped machines have no EOR data** (parcel sorters, SDUS, F/C, canceller). Their PM
   story lives in NMARS eCBM — pull `EQUIP_SITE` + `PM4WKS_SITE` for them, or RFI it.

## 5. Gaps this file can't answer (→ RFI / enriched pull)

- Real machine-reported **down/idle/jam/maintenance minutes** (this 12-column export lacks
  them) → use the grabber's 19-column `EOR_RUNS` pull.
- **TACS overtime hours/cost by tour** — the 8→10-window straight-time conversion dollars.
- **eCBM PM due-vs-completed** for the 15 non-EOR machines and the DBCS fleet.
- **Staffing complement vs vacancies** and **contractor invoices** (see the RFI slide in the deck).
