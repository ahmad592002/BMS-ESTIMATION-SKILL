# Phase 3 - DDC List

Produces: the panel/equipment assignment matrix.

Row 1 index, row 2 equipment names from `EquipmentList`, row 3 `Total`, row 4 `Assigned`,
then **one row per panel from row 5 down**, panel name in column A.

## Rules

- **Panel count is a tender quantity when the client states one.** If the SOW's bill of quantities
  says "supply N DDC panels", price N - the client is paying per panel and bids are compared on that
  line. Engineering a different count is a commercial decision for the estimator, not a silent fix.
  Where the count is not stated, size panels to roughly 200 points each.
- Name panels by location (`DDC-GF-01`, `DDC-R-02`, `DDC-B1-03`) or by the unit served
  (`DDC-03 (AHU-03)`) when locations are unknown. Where the drawings give panel tags, keep them.
- **GTS uses PXC, never PLC.** Panels drawn as PLC / RIO become `DDCP-<bldg>-<level>-<nn>`.
- An overloaded drawn panel splits into separate panels `<name>-1`, `-2`; IP unitary controllers
  (FCU/VAV) go on network rows split floor by floor - every row <= 250 incl. spare.
- Build with `scripts/ddc_build.ps1 -Write` (placement from the takeoff; `el_map.txt`,
  `host_map.txt`, `panel_rename.txt` are per-project data files) and format with `full_chain.ps1`.
- **Group by physical proximity first** (same plantroom or floor), then by discipline. Without a floor
  plan, group by served unit and say so - real grouping follows plantrooms, and the pairing will shift
  once drawings arrive even if the count does not.
- Software-point equipment (VAVs, thermostats, FCUs) attaches to the panel serving its zone. Without
  a zoning schedule, distribute pro-rata and flag it as an assumption.

## Gates - never present a failing matrix

1. **`Assigned` == `Total` for every column.** Verify each column individually rather than trusting
   the row-4 sum. A mismatch means equipment unassigned or double-counted.
2. **No panel over 250 points** - physical (DI+AI+AO+DO) plus software points, tested **with** the
   `Options` spare factor applied. A panel at 220 real points is already over once 20% is added.
   Split rather than exceed.

## Present

Panel list with per-panel equipment, physical and SP counts, **each panel's loading against 250**, and
explicit confirmation that Assigned == Total across all columns.
