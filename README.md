# GTS BMS Estimation Skill

A [Claude Code](https://claude.com/claude-code) skill that builds a GTS BMS estimation from raw tender
data (equipment schedules, mechanical/electrical specs, SOW, drawings) by filling the
**BMS Template 2026** workbook one reviewed phase at a time.

> **Private repo.** Contains GTS margin conventions, supplier part selection logic and pricing
> mechanics. Keep it private.

## What it does

Runs the estimation as eight gated phases, in the estimator's own order. **One phase per invocation** —
it produces the phase, presents it for review, and stops.

```
0 Setup & Options
1 EquipmentList        <- the single source of truth
2 IOSummary            <- points per equipment type
3 DDC List             <- equipment -> panel assignment   (gate: Assigned == Total)
4 DDCSummary           <- per-panel IO + controller selection
5 DDCFullSummary       <- point-by-point expansion        (gate: 3-way reconciliation)
6 Ancillaries          <- Workstation, dampers, valves, VFD/MCC, FCU
7 BOQ / Breakdown / Cover Page                            (gate: 7-point final audit)
```

Each phase is derived from the one before it, so nothing downstream is ever filled by hand.

## Two things that make it different

**It asks instead of guessing.** Missing or contradictory quantities, equipment it can't map to a
template type, ambiguous scope, unstated protocol — all become numbered questions with proposed
defaults, never silent assumptions.

**It upgrades itself.** When the estimator corrects something, it applies the fix, then asks *why* —
spec requirement, GTS standard, client preference, project condition, its own error, or a commercial
call. The answer sets the rule's scope, and the rule is written to `references/LEARNED.md`, which is
read first on every run and overrides `SKILL.md` on conflict.

## Layout

```
skills/bms-estimation/
├─ SKILL.md                     the workflow and phase gates
└─ references/
   ├─ LEARNED.md                accumulated rules, with reason + scope
   └─ dump_sheets.ps1           dump any .xlsm to plain text without Excel
```

## Install

Copy (or symlink) the skill into your Claude Code skills directory:

```powershell
# user-level - available in every folder you open
Copy-Item -Recurse skills\bms-estimation "$env:USERPROFILE\.claude\skills\"
```

To keep the live skill and the repo in sync automatically, symlink instead (run as Administrator, or
with Developer Mode enabled):

```powershell
New-Item -ItemType SymbolicLink `
  -Path   "$env:USERPROFILE\.claude\skills\bms-estimation" `
  -Target "$PWD\skills\bms-estimation"
```

Then invoke it in any project folder:

```
/bms-estimation
```

## Per-project state

The skill writes `_ESTIMATION_STATE.md` into each **project** folder — phase checklist, decisions,
open questions, key figures. State stays with the job; the skill stays global. Those files are not
tracked here.

## Reading a workbook without Excel

```powershell
powershell -File skills\bms-estimation\references\dump_sheets.ps1 `
  -Xlsm "<file.xlsm>" -OutDir out -MaxRows 200
```

Writing into these workbooks needs Excel/COM — a plain XLSX library strips the VBA project and breaks
all the hidden selection logic.
