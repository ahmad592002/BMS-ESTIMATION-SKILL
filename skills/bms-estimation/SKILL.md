---
name: bms-estimation
description: Build a GTS BMS estimation from raw tender data (equipment schedules, mechanical/electrical specs, SOW, drawings, BOQ requests) one gated phase at a time - EquipmentList -> IOSummary -> DDC List -> DDCSummary -> DDCFullSummary -> Ancillaries -> BOQ/Breakdown/Cover Page - stopping after each phase for the estimator's review, and recording every correction back into this skill. Use whenever a folder contains BMS/HVAC tender material, an equipment or points schedule, a BMS SOW, or a "*.xlsm" estimation template, or when asked to price/estimate a BMS, count IO points, size DDC controllers, or produce a BMS BOQ.
---

# BMS Estimation (GTS) - orchestrator

**This skill runs exactly ONE phase per invocation, then stops and waits for review.**
Never run two phases in a single turn, even when the next looks trivial. A mistake caught at Phase 1
costs a minute; the same mistake found at Phase 7 costs the whole estimation.

This file is the orchestrator only. The work for each phase lives in its own module, loaded when that
phase runs - do not load them all.

## Every invocation

1. **Read `references/LEARNED.md`** - accumulated rules from past runs. These **override** anything in
   this file or the phase modules wherever they conflict.
2. **Find `_ESTIMATION_STATE.md`** in the project folder.
   - Missing -> you are at Phase 0. Create it from `references/state-template.md`.
   - Present -> the current phase is the first not marked `DONE`.
3. **Read the phase module for that phase, and only that one** (table below).
4. **Read `references/workbook-mechanics.md`** before any read or write to the workbook.
5. **Announce** in one line: project, phase number and name, what it will produce.
6. **Do only that phase**, following its module.
7. **Self-check** against the module's exit criteria. Fix failures before presenting. If the change
   touched anything upstream of the BOQ, regenerate every downstream sheet (DDC List -> DDCSummary ->
   DDCFullSummary -> Workstation -> DamperActuators -> ValvesAndActuators -> VFDs/RoomUnits -> BOQ ->
   Product Finder) - see LEARNED.md - and run `scripts/audit.ps1`.
8. **Present for review** - the real rows and counts, not "done" - with numbered assumptions and
   numbered questions, each carrying a proposed default so "yes to all" is a valid reply.
9. **Update `_ESTIMATION_STATE.md`**: mark the phase `AWAITING REVIEW`, log decisions and questions.
10. **STOP.** End with: `Phase N ready for your review. Reply with corrections, or "next" to move to
    Phase N+1.`

Next invocation: corrections -> apply them to the *current* phase, re-present, stay put. Mark `DONE`
only on approval or an explicit "next".

## Phase modules

| Phase | Module | Produces |
|---|---|---|
| 0 | `phases/phase0-setup.md` | Input survey, scope, template copy, `Options` |
| 1 | `phases/phase1-equipment-list.md` | `EquipmentList` - the single source of truth |
| 2 | `phases/phase2-io-summary.md` | `IOSummary` - points and field devices per type |
| 3 | `phases/phase3-ddc-list.md` | `DDC List` - equipment to panel assignment |
| 4 | `phases/phase4-ddc-summary.md` | `DDCSummary` - per-panel IO + controller selection |
| 5 | `phases/phase5-ddc-full-summary.md` | `DDCFullSummary` - point-by-point expansion |
| 6 | `phases/phase6-ancillaries.md` | Workstation, dampers, valves, VFDs, FCUs |
| 7 | `phases/phase7-boq-breakdown-cover.md` | `BOQ`, `Breakdown`, `Cover Page` + final audit |

Each stage derives from the one before it. **Never fill a downstream sheet by hand** when it can be
derived upstream, and never skip forward.

```
raw data -> EquipmentList -> IOSummary -> DDC List -> DDCSummary -> DDCFullSummary
                                                                         |
                            Workstation + dampers + valves + VFDs ->  BOQ -> Breakdown -> Cover Page
```

## Proven way of working (Al Moosa University, completed 2026-10-07 - keep doing it this way)

- **Phase 1 - equipment from the drawings, not the BOQ.** Read the B-92 BMS risers as image tiles;
  build the per-tag takeoff workbook (tag -> DDC/RIO/PLC -> room -> interface, duplicates kept at
  qty 0); EquipmentList = one row per family per building, synced to the takeoff.
- **Phase 2 - points from the B-93 schematics, every sheet.** Parse BMS schedule tables and DDC dot
  strips (`scripts/`), read plant sheets by eye, show the per-unit list and a review workbook
  *before* generating. Generate with the EquipmentList **button**, then complete the rest
  (rows > 120), group points **by component**, name them as drawn, add **field devices** only
  where drawn with **previous-project models**, rewrite H:L, verify totals.
- **Every deliverable gets a "where it came from" workbook** (takeoff; IO Summary Sources).
- **Yellow only the cell in doubt**, reason beside it (IOSummary column Q, EquipmentList column F).
  Assumed software-point counts are always yellow.
- **Project root holds only final files**; backups and superseded files go to `Old Versions`.
- **Phase 3 - PXC panels only** (drawn PLC/RIO -> DDCP), every row <= 250 incl. spare, network rows
  floor by floor; **FCU/VAV = 7 SP integration**; **higher number wins** on every quantity doubt.
- **Valves priced on two rows** (PICV / globe / butterfly models in LEARNED C); sizes yellow.
- **Phase 7 - BOQ -> pf_refill -> Breakdown**; Product Finder G11 = 0 (discount once, in E23).
- **Finish**: regenerate the side files from the final workbook; a final revision is read-only and
  asks one finding at a time; the estimator names the final file - never rename it.
- **Excel automation**: `references/workbook-mechanics.md` + LEARNED H (Start-Process, active
  workbook, `run_macro.ps1`, case-insensitive variables, no int-keyed ordered dictionaries).

Scripts for all of this: `scripts/` (see its README for the run order). They all read the workbook
path from `$env:BMS_WB`.

## Hard gates

These are refusals, not warnings. Never present a phase that fails its gate - fix it first.

| Phase | Gate |
|---|---|
| 3 | `Assigned` == `Total` on **every** column of `DDC List` |
| 3, 4 | **No DDC over 250 points** (physical + software), tested **with** the `Options` spare factor |
| 4 | Every panel's module capacity >= its IO demand including spare |
| 5 | DDCFullSummary == DDCSummary == IOSummary point totals |
| 7 | Product Finder lines == BOQ Siemens lines, PF `G11` = 0, Breakdown B.1 ~ BOQ Siemens cost (customs only) before any margin is quoted |
| 7 | Every BOQ line qty == its source sheet; every cost cell a Pricelist VLOOKUP; no #N/A |
| 7 | Cover Page total == Breakdown Net Value == BOQ total |

## When the data is not good - ask, do not guess

Stop and ask, inside the current phase, whenever: a quantity is missing or contradicts another source;
equipment cannot be mapped to a `TemplateLists` type; scope is ambiguous (supply only / + install /
+ T&C; is MCC ours? is FA ours?); protocol, brand or spare factor is unstated; the client BOQ and the
mechanical schedule disagree; something is drawn but never scheduled.

Ask as a numbered list with a proposed default for each. Never silently invent equipment, and never
silently drop what you cannot classify - list it as an open question.

## Self-upgrade: ask WHY, then record

Every estimator edit is a lesson, but you do not yet know what the lesson is. The same edit can mean
five different things, each updating the skill differently:

| Their reason | Effect |
|---|---|
| **Spec / SOW** requires it | Permanent rule - and check the spec for what else you missed |
| **GTS standard practice** | `Scope: always` |
| **This client** always wants it | `Scope: this client` |
| **Site / project condition** | `Scope: this project` - do NOT generalise |
| **You made an error** | Fix the rule that let the error through |
| **Commercial call** (margin, strategy) | Record the reasoning, never auto-apply the number |

1. **Apply the correction first.** Never hold up the work to interrogate it.
2. **Then ask why - one short question**, grouping related edits, offering the likely reasons so a
   one-word reply works.
3. **Classify and record** - their answer sets `Scope`.
4. If they answer a different question than you asked, follow their answer.
5. **If they don't explain**: record with `Reason: not given`, `Scope: this project`. Ask once, never
   nag, never re-ask in the same project.
6. **Fix causes, not symptoms.** If the edit happened because you misread a schedule, the rule is about
   reading that schedule. If it contradicts a phase module, edit that module too.
7. **Confirm in one line:** `Learned: <rule> (scope: <scope>).`

Skip the why-question for typos and formatting only.

Entry format and the full protocol: `references/LEARNED.md`.

### Then publish as a pull request

Repo: `C:\Users\ahmad\repos\gts-bms-estimation-skill` -> `ahmad592002/BMS-ESTIMATION-SKILL`.
**Every self-upgrade goes up as a PR, never a direct commit to `main`.** The rule applies locally at
once; the PR is the review trail and must never block the estimation.

Once per phase at most, batching that phase's rules:

1. Write the rule to the live skill first (`~/.claude/skills/bms-estimation/`).
2. Copy changed files into the repo (or it is already there if symlinked).
3. `git checkout main && git pull --ff-only`, then `git checkout -b learn/<phase>-<slug>`.
4. Commit with the reason in the message (project, phase, change, reason type, scope).
5. Push, then open the PR with `gh pr create`, or hand over the compare link:
   `https://github.com/ahmad592002/BMS-ESTIMATION-SKILL/compare/main...<branch>?expand=1`
6. Return to `main`.
7. Report: `Learned: <rule> (scope: <scope>) - PR: <url>`

Never push to `main`, never self-merge, never commit client data (`.gitignore` blocks `*.xlsm`,
`*.xlsx`, `*.pdf`, `_ESTIMATION_STATE.md` - do not force past it). If anything fails, say
`rule saved locally; PR pending: <reason>` and carry on with the phase.

## Reference material

| File | Load when |
|---|---|
| `references/LEARNED.md` | **Every invocation, first** |
| `references/workbook-mechanics.md` | Before any workbook read or write |
| `references/state-template.md` | Creating `_ESTIMATION_STATE.md` at Phase 0 |
| `references/dump_sheets.ps1` | Reading a workbook without Excel |
| `scripts/README.md` | Phases 1-2: takeoff, schematic parsing, IOSummary build, devices, checks |

Reference estimations live in `~/OneDrive/Desktop/GTS/BMS_PROJECTS/` (Ajyad Tower, YALJ, Prince
Mansour, RX Premium Hub, Al Moosa University). **YALJ** is the best filled-template example (all 35
sheets); **07- Al Moosa University** is the best full-process example - drawings-only tender (3
buildings, 200 equipment lines, 116 PXC panels), with its takeoff, IO / DDC sources and state file.
**REHAB OASIS RIYADH, KSA** (`~/OneDrive/Desktop/REHAB OASIS RIYADH, KSA/`) is the example of a
consultant-IO-list-only tender continued in the estimator's own workbook (2 buildings, 31 equipment
rows, 21 panels + 3 EMS network rows, 1.28 M SAR).
