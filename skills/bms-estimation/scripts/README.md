# BMS estimation scripts (proven on Al Moosa University, 2026-10)

PowerShell 5.1 + Excel COM + MiKTeX poppler (`pdftotext -bbox`, `pdftoppm`). Copy this folder into
the session scratchpad, put the project's own data files next to it, edit the workbook path at the
top of the workbook scripts, then run in the order below. **Project data files are NOT kept here**
(they hold client tags and quantities): `panels.txt`, `items.txt`, `io_map.txt`, `eqlist_now.txt`,
`points_all.csv`, `sched_all.tsv`, `strip_all.tsv`, `txt\` (pdftotext -layout of each sheet).

| Step | Script | Input -> output |
|---|---|---|
| Phase 1 takeoff | `build_takeoff.ps1 -Out <xlsx>` | `panels.txt` (panel/room) + `items.txt` (tags per panel, `#n`/`?n`/`!tag` syntax) -> per-tag takeoff workbook |
| Schedule tables | `parse_sched.ps1 -Dir <Site> -Out sched_all.tsv` | B-93 sheets with "BMS/PMS SCHEDULE" -> one row per point |
| Dot strips | `parse_strip.ps1 -Dir <Site> -Out strip_all.tsv [-Only nnn]` | DDC point strips -> one row per mark (filled dot / ring), x-multipliers, inferred types |
| Merge | `combine.ps1` | table wins over strip; plant sheets entered by hand inside the script -> `points_all.csv` |
| Templates | `io_templates.txt` + `io_map.txt` (one template key per EquipmentList row) + `io_lib.ps1` | `@SHEET:nnn[:group]` lines pull drawn points |
| Dry run | `build_iosummary.ps1 -DryRun` | per-row IO + expected grand totals (no workbook change) |
| Generate | `run_all.ps1` | full clear -> EquipmentList button -> `fill_iosummary.ps1` (append rows > 120, fill blocks by component) -> `fix_hl.ps1` |
| Devices | `fill_devices.ps1` | field devices (drawn instruments, previous-project models), M = per unit x qty |
| Checks | `apply_flags2.ps1` (rules `io_flags.ps1`, `io_flags2.ps1`) | yellow on the doubtful cell only + reason in column Q; EquipmentList B/C + F |
| Sources | `build_io_source.ps1 -Out <xlsx>` | point-by-point source workbook (drawing no., method, basis, device, checks) |

Gotchas these scripts already handle (see `references/LEARNED.md`): PowerShell variables are
case-insensitive (`$W`/`$w`, `$G`/`$g`, `$T`/`$t` collide); `@(List)` throws "Argument types do not
match" in PS 5.1; cast every value written to a cell; never `$x.Run` a macro that takes arguments;
`New-Object` Excel instances are not in the ROT - reach an open workbook with
`[Runtime.InteropServices.Marshal]::BindToMoniker(<path>)`; `CutCopyMode` cannot be set on a
`New-Object` instance; wait for `$x.Ready` after a long macro; write non-ASCII device names
(`0…2000 PPM`) with `[char]0x2026`.

**Per-project sections to replace on the next project:** `combine.ps1` (plant points read by eye),
`io_flags.ps1` `$EqFlag` and `io_flags2.ps1` `$EqQty` / `$EqName` (check reasons keyed by
EquipmentList SN), `build_takeoff.ps1` `$iss` (drawing issues) and `$cmp` (drawing vs BOQ),
`build_io_source.ps1` sheet titles for untitled sheets, `io_lib.ps1` `$FilterNames` (filter order
per sheet). `io_templates.txt` is a starting library - extend it per project.
