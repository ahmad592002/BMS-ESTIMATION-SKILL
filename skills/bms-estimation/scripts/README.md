# BMS estimation scripts (proven on Al Moosa University, completed 2026-10-07)

PowerShell 5.1 + Excel COM + MiKTeX poppler (`pdftotext -bbox`, `pdftoppm`). Copy this folder into
the session scratchpad and put the project's own data files next to it.

**Every workbook script reads `$env:BMS_WB`** (full path of the estimation workbook) and throws if it
is not set. Open the workbook in Excel first (`Start-Process $env:BMS_WB`) - the scripts attach with
`BindToMoniker`, which starts a hidden, short-lived instance if Excel is closed. Side files are named
`<project> - ...xlsx`, where `<project>` = `$env:BMS_PROJECT` or the project folder name without its
"07- " prefix. `audit.ps1` allows extra root files listed in `$env:BMS_KEEP` (`;`-separated).

```powershell
$env:BMS_WB = "C:\...\07- <Project>\<Project>-BMS ESTIMATION.xlsm"
& .\run_macro.ps1 -Macros 'DDCSummaryModule.GenerateFullDDCSummary' -Sheet 'DDCFullSummary'
```

**Project data files are NOT kept here** (client tags and quantities): `panels.txt`, `items.txt`,
`io_map.txt`, `eqlist_now.txt`, `points_all.csv`, `sched_all.tsv`, `strip_all.tsv`, `txt\`,
`el_map.txt` (`SN|building|regex||` takeoff -> EquipmentList), `host_map.txt` (`SN|host SN`: place a
qty-0 / BOQ-extra line on the panel of another line), `panel_rename.txt` (`drawn tag|DDCP name`).

## Phase 1-2 (build)

| Step | Script | Input -> output |
|---|---|---|
| Takeoff | `build_takeoff.ps1 -Out <xlsx>` | `panels.txt` + `items.txt` -> per-tag takeoff workbook |
| Schedule tables | `parse_sched.ps1 -Dir <Site> -Out sched_all.tsv` | B-93 "BMS/PMS SCHEDULE" tables -> one row per point |
| Dot strips | `parse_strip.ps1 -Dir <Site> -Out strip_all.tsv [-Only nnn]` | DDC point strips -> one row per mark |
| Merge | `combine.ps1` | table wins over strip; plant points by eye inside the script -> `points_all.csv` |
| Review | `build_io_review.ps1 -Out <xlsx>` | per-equipment point list for the estimator BEFORE generating |
| Templates | `io_templates.txt` + `io_map.txt` + `io_lib.ps1` | `@SHEET:nnn[:group]` lines pull drawn points |
| Dry run | `build_iosummary.ps1 -DryRun` | per-row IO + expected grand totals |
| Generate | `run_all.ps1` | full clear -> EquipmentList button -> `fill_iosummary.ps1` (rows > 120, blocks by component) -> `fix_hl.ps1` |
| Devices | `fill_devices.ps1` | field devices, M = per unit x qty |
| Valves | `apply_valves.ps1 [-DryRun]` | two priced rows per valve (PICV VPF44.65F25+SAX61P03, globe VVF42.65-50+SKB62/F) + Valves direct list |
| Checks | `apply_flags2.ps1` (rules `io_flags.ps1`, `io_flags2.ps1`) | yellow on the doubtful cell + reason in Q |
| Sources | `build_io_source.ps1 -Out <xlsx>` first time; `refresh_sources.ps1` after edits / at the end | IO Summary Sources workbook |
| Cross-project | `scan_valves.ps1` | valve / device models used in the other GTS projects |

## Phase 3-7 (regenerate, price, check)

| Script | What it does |
|---|---|
| `ddc_assign.ps1` + `ddc_load.ps1` | takeoff tag -> EquipmentList row mapping; panel loads |
| `ddc_build.ps1 [-Write]` | DDC List: splits > 250 into -1/-2, network rows floor by floor, BOQ-extra and qty-0 lines placed (`host_map.txt`), PLC/RIO renamed (`panel_rename.txt`); gate Assigned == Total |
| `ddc_sources.ps1` | DDC List Sources workbook (controller loading + tags per controller) |
| `run_macro.ps1 -Macros a,b -Sheet S` | runs template macros safely: activates workbook + sheet, verifies ActiveWorkbook, saves before/after, waits for Ready, retries 0x800A9C68 once |
| `full_chain.ps1` | everything downstream of an upstream edit: DDC List + formatting -> DDCSummary -> controllers -> Full summary -> Workstation -> dampers -> valves / VFDs -> BOQ -> `pf_refill.ps1` |
| `pf_refill.ps1` | Product Finder C17/E17 = every Siemens BOQ line; reports K14, unpriced lines, selling per category |
| `audit.ps1` | read-only re-check of every phase + side files - run before presenting and after estimator edits |
| `revision_facts.ps1` | read-only facts per phase for a final revision (options, counts, devices, panels, BOQ, Breakdown, Cover) |
| `gm_analysis.ps1` | margin by category |
| `review_blocks.ps1` | IOSummary blocks for review (duplicates, empty blocks) |
| `compare_wb.ps1 -Backup <xlsm>` | sheet-by-sheet diff of the live workbook against a backup - "what did the estimator change?" |
| `dump_io.ps1` / `diff_io.ps1` | IOSummary dump / diff |

Not included: `pf_refill` does not touch Product Finder G11 - check it is 0 (LEARNED G).

## Gotchas these scripts already handle (LEARNED H)

Case-insensitive variables (`$K`/`$k`, `$W`/`$w` collide); never index an `[ordered]` dictionary with
an int; `@(List)` throws in PS 5.1; cast every value written to a cell; never `$x.Run` a macro with
arguments; macros act on the ActiveWorkbook; `CutCopyMode` cannot be set on a `New-Object` instance;
wait for `$x.Ready` after a long macro; non-ASCII device names (`0…2000 PPM`) with `[char]0x2026`.

**Per-project sections to replace on the next project:** `combine.ps1` (plant points read by eye),
`io_flags.ps1` `$EqFlag`, `io_flags2.ps1` `$EqQty` / `$EqName` (reasons keyed by EquipmentList SN),
`build_takeoff.ps1` `$iss` / `$cmp`, `build_io_source.ps1` titles for untitled sheets, `io_lib.ps1`
`$FilterNames`, `apply_valves.ps1` `$MODELS` if the estimator picks other models.
`io_templates.txt` is a starting library - extend it per project.
