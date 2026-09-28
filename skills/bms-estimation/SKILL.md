---
name: bms-estimation
description: Build a GTS BMS estimation from raw tender data (equipment schedules, mechanical/electrical specs, SOW, drawings, BOQ requests) one gated phase at a time - EquipmentList -> IOSummary -> DDC List -> DDCSummary -> DDCFullSummary -> Ancillaries -> BOQ/Breakdown/Cover Page - stopping after each phase for the estimator's review, and recording every correction back into this skill. Use whenever a folder contains BMS/HVAC tender material, an equipment or points schedule, a BMS SOW, or a "*.xlsm" estimation template, or when asked to price/estimate a BMS, count IO points, size DDC controllers, or produce a BMS BOQ.
---

# BMS Estimation (GTS) - one phase per run

**This skill runs exactly ONE phase per invocation, then stops and waits for review.**
Never run two phases in a single turn, even when the next one looks trivial. Each phase is a gate the
estimator opens; a mistake caught at Phase 1 costs a minute, the same mistake found at Phase 7 costs
the whole estimation.

## Every invocation follows this loop

1. **Read `references/LEARNED.md`** - project-specific and hard-won rules from past runs. These
   override the generic guidance in this file wherever they conflict.
2. **Find the state file** `_ESTIMATION_STATE.md` in the project folder.
   - Missing -> you are at Phase 0. Create it (template at the bottom of this file).
   - Present -> read it. The current phase is the first one not marked `DONE`.
3. **Announce** in one line: project, phase number and name, what this phase will produce.
4. **Do only that phase.** Work from the phase's own section below.
5. **Self-check** the phase against its exit criteria. If a check fails, fix it before presenting.
6. **Present the output for review** - the actual rows/counts in the chat, not just "done". Include:
   - what you produced (compact table or counts),
   - **every assumption you made**, numbered,
   - **every question you need answered**, numbered, each with your proposed default so a "yes to all"
     is a valid reply.
7. **Update `_ESTIMATION_STATE.md`**: mark the phase `AWAITING REVIEW`, log decisions and open questions.
8. **STOP.** Do not start the next phase. End with: `Phase N ready for your review. Reply with
   corrections, or "next" to move to Phase N+1.`

On the next invocation, if the user gave corrections: apply them to the *current* phase, re-present,
stay on the same phase. Mark `DONE` only once they approve or say "next".

## When the data is not good - ask, do not guess

Stop and ask, inside the current phase, whenever:

- a quantity is missing, contradictory between schedule and drawings, or clearly a typo;
- the schedule names equipment whose type/function you cannot map to a `TemplateLists` category;
- scope is ambiguous (supply only / supply + install / + T&C; is MCC ours? is FA ours?);
- protocol, brand, or spare factor is not stated anywhere in the spec;
- the client BOQ and the mechanical schedule disagree on counts;
- something is implied by the drawings but never scheduled.

Ask in a numbered list with a proposed default for each. Do not silently invent equipment, and do not
silently drop equipment you cannot classify - list it as an open question instead.

## Self-upgrade: ask WHY, then record

Every time the estimator changes something you produced - a quantity, a Type, a controller choice, a
panel grouping, a price, a wording - **that edit is a lesson, but you do not yet know what the lesson
is.** The same edit can mean five different things, and each one updates the skill differently:

| The estimator's reason | What it means for the skill |
|---|---|
| The **spec / SOW** requires it | Permanent rule - and check the spec for what else you missed |
| **GTS standard practice** | `Scope: always` - this applies to every future project |
| **This client** always wants it | `Scope: this client` |
| **Site / project condition** | `Scope: this project` - do NOT generalise |
| **You made an error** (misread, bad assumption) | Fix the rule that let the error through |
| **Commercial call** (margin, strategy) | Record the reasoning, never auto-apply the number |

So the protocol is:

1. **Apply the correction first.** Never hold up the work to interrogate it.
2. **Then ask why - in one short question, not an interview.** Group several edits into one question
   when they share a likely cause. Offer the likely reasons so a one-word reply works:

   > You changed the FCU count from 6 to 14 and set the kitchen fans to Monitor only.
   > Is that (a) from the spec/SOW, (b) GTS standard, (c) this client's preference,
   > (d) specific to this project, or (e) me misreading the schedule?

3. **Classify and record** using the answer - `Scope` comes directly from their reason.
4. **If they answer a different question than you asked, follow their answer.** They know the job.
5. **If they don't want to explain** ("just do it", no answer, or they move on): record it anyway with
   `Reason: not given` and `Scope: this project`, the safe default. Never nag - ask once, then drop it.
   Do not re-ask the same why later in the same project.
6. **When the reason reveals a deeper cause, fix the cause.** If the edit happened because you misread
   the schedule, the rule to record is about how to read that schedule - not just the corrected number.
   If it contradicts this SKILL.md, **edit SKILL.md too** and note that in the entry.
7. **Confirm in one line:** `Learned: <rule title> (scope: <scope>).`

Skip the why-question only for pure typo fixes and formatting nits - nothing to learn there.

Append to `references/LEARNED.md`, which is read at the start of every invocation:

```markdown
## <short rule title>
- **Phase:** <number/name>   **Date:** <YYYY-MM-DD>   **Project:** <name>
- **Rule:** <the rule, imperative, one or two lines>
- **Change:** <what the estimator actually changed - from -> to>
- **Reason:** <their answer, in their terms> | not given
- **Reason type:** spec | GTS standard | client preference | project condition | my error | commercial
- **Scope:** always | this client | this project type | this project
- **Also changed in SKILL.md:** <section, or "no">
```

One entry per rule. If a rule already exists, **edit it** rather than adding a near-duplicate - and if
a new reason contradicts an old entry, update that entry and note the change; the newer reason wins.

### Then publish the change as a pull request

The skill is version-controlled at `C:\Users\ahmad\repos\gts-bms-estimation-skill`
(`origin` = https://github.com/ahmad592002/BMS-ESTIMATION-SKILL). **Every self-upgrade goes to the repo
as a pull request - never a direct commit to `main`.** The rule takes effect locally straight away; the
PR is the review trail, so it must never block the estimation.

Do this **once per phase at most**, when the phase is presented for review - batch that phase's learned
rules into one PR rather than pushing per edit.

1. **Write the rule to the live skill first** (`~/.claude/skills/bms-estimation/`). The estimation
   continues on the updated rule immediately, PR or no PR.
2. **Get it into the repo working tree.** If `~/.claude/skills/bms-estimation` is a symlink into the
   repo, it is already there. Otherwise copy the changed files across:
   `skills/bms-estimation/references/LEARNED.md` and `SKILL.md` if that changed too.
3. **Branch** off an up-to-date `main`:
   `git -C <repo> checkout main && git -C <repo> pull --ff-only`
   then `git -C <repo> checkout -b learn/<phase>-<short-slug>` (e.g. `learn/p1-fcu-count-from-spec`).
4. **Commit** with the reason in the message, so the PR explains itself:
   ```
   Learn: <rule title>

   Project: <name>   Phase: <n>
   Change: <from -> to>
   Reason (<reason type>): <estimator's words>
   Scope: <scope>
   ```
5. **Push** the branch: `git -C <repo> push -u origin <branch>`
6. **Open the PR.** With `gh` installed:
   `gh pr create --repo ahmad592002/BMS-ESTIMATION-SKILL --base main --head <branch> --title "Learn: <rule title>" --body "<reason, scope, and what it changes>"`
   Without `gh`, push the branch and give the estimator the compare link to click:
   `https://github.com/ahmad592002/BMS-ESTIMATION-SKILL/compare/main...<branch>?expand=1`
7. **Return to `main`** in the repo afterwards so the next branch starts clean.
8. **Report in one line** alongside the phase output:
   `Learned: <rule> (scope: <scope>) - PR: <url>`

Rules for this step:
- **Never push to `main` directly**, and never merge your own PR - the estimator reviews and merges.
- **Never commit client data.** Only `SKILL.md` and `references/*` belong in the PR; the `.gitignore`
  already blocks `*.xlsm`, `*.xlsx`, `*.pdf` and `_ESTIMATION_STATE.md`, so don't force-add past it.
- **If anything fails** - no network, auth prompt, push rejected, permission denied - the rule is
  already saved locally, so say so plainly (`rule saved locally; PR pending: <reason>`) and carry on
  with the phase. Do not retry in a loop, and do not stall the estimation over git.
- If several rules accumulated while offline, open one PR covering them all when the push next works.

---

# The phases

## Phase 0 - Survey and setup

Produce: an inventory of the input files, the `Options` sheet filled, a working copy of the template.

1. List every file in the folder and say what each one is. Typical inputs: equipment schedule (xlsx),
   BMS SOW / scope pdf, spec booklet (often Arabic: `كراسة الشروط والمواصفات`), mechanical + electrical
   drawings, a client BOQ to be priced, contract form.
2. Read the SOW and spec for: scope boundary, protocol, brand restrictions, spare requirement,
   integration list, workstation/server requirement, warranty and T&C obligations.
3. Copy the newest master from `~/OneDrive/Desktop/GTS/BMS_PROJECTS/BMS Template 2026 V03.xlsm` into
   the project folder as `<Project> - GTS offer.xlsm`. **Never edit the master in place.**
   - **V27.1** = `Product_Finder_SI_B_AUT_V27.1` + `Listprice_V27.1_AUT_A_SP` (use this by default).
   - **V26.2** = `Product_FinderV26.2` + `Listprice_V26.2_HVAC_A_SP` (only if the client locked it).
4. Set `Options` - every hidden lookup sheet reads from it, so changing it later re-selects hardware
   underneath finished work:

   | Option | Typical | Note |
   |---|---|---|
   | Spare in DDC | `0.2` | 20% spare IO; raise if spec demands |
   | Protocol | `BACNET/IP` | per spec |
   | Controller Type | `Modular` | Modular vs compact |
   | Brand / FD / DA / VA / VFD Brand | `Siemens` | field devices may differ (e.g. AX) |
   | Style | `IO+DDC` | |
   | Enclosure | `Al Fanar` | |
   | Threaded Threshold | `50` | mm; above = flanged |
   | Spring Return | `Yes` | |
   | Default valve sizes | AHU `65`, BV `200`, FCU `25` | overridden per unit |
   | FCU Communication Protocol | `KNX` | KNX / BIP / MSTP / Modbus / Standalone |
   | Include MCC in BOQ | `No` | `Yes` only when MCC is in scope |
   | VA / VFD / MCC selection by | `2` (KVS / HP / HP) | leave alone |

   Workbook macros: **Ctrl+Shift+I** insert new line, **Ctrl+Shift+G** generate IO points from templates.
   Prefer these over copy-paste - they keep formulas intact.

**Exit criteria:** every input file accounted for; scope stated in one sentence; `Options` values
confirmed with the user (present them as a table for approval - this is the phase's review).

## Phase 1 - EquipmentList

Produce: `EquipmentList` filled. Columns `SN | Qty | Equipment | Function | Type`.

- One row per **equipment family**, not per tag: `Exhaust Fan` qty 23, not EF-01...EF-23.
- `Qty` from the schedule/drawings. It propagates to every later sheet - verify against two sources
  where possible and flag any disagreement as a question.
- `Function` is exactly `Monitor` or `Monitor&Control`. Monitoring-only gets status/alarm DIs and no
  outputs; this one word decides the IO count.
- `Type` must come from the `TemplateLists` picklist: `Chilled Water System`, `Air Handling Unit`,
  `Fan Coil Unit`, `VAV Box`, `Fan`, `Pump`, `Electrical Panel`, `Tank`, `Low Current Systems`,
  `Meters`, `Other`. Wrong Type = wrong IO template = wrong controller.
- Keep duty/standby notation from the schedule (`(2D/1S)`) - it drives DI/DO counts.
- Cover all disciplines, not just HVAC: tanks, water treatment/RO, boilers, diesel, LPG, CO/gas
  detection, lighting, MDB/SMDB/PFC, generator, ATS, UPS, low current, CCTV, fire alarm interface.
- Monitoring-only system interfaces (CCTV, FA, low current) are normally qty `1` per system.

**Exit criteria:** every scheduled item is either a row or a listed open question; no `Type` outside the
picklist; no blank `Function`; totals per discipline presented for review.

**Present:** the full table (SN, Qty, Equipment, Function, Type) plus a discipline subtotal line.
This is the phase the estimator checks hardest - make it easy to scan.

## Phase 2 - IOSummary

Produce: one block per equipment **type**, points counted.

```
r(n)    A=QTY   B=<Type>   C="IO (1 System)"   H="IO (All Systems)"   M="Field Devices"
r(n+1)  B=Location:  C=DI D=AI E=AO F=DO G=SP   H=DI I=AI J=AO K=DO L=SP   M=QTY N=Field Device O=Model P=Comments
rows                 <point name> with a 1 in its column
r(last) B=TOTAL  + column sums
```

- `C:G` = one system. `H:L` = `C:G x QTY` - **formulas, never typed**.
- Generate with **Ctrl+Shift+G** from `IOTemplate` / `Template` rather than typing points.
- `Template` drives parametric systems - Chilled Water System asks Nb. of chillers, pipe size,
  primary/secondary pump counts, DPS/DPT, expansion tank, bypass + valve, valve type/size, supply and
  return temps. Fill from the schedule, let the points generate.
- Sanity patterns: Pump = Run Status, Trip Status (+ Start/Stop DO, Auto/Manual if `Monitor&Control`);
  Tank = High level, Low level, Level Sensor (AI); AHU/FAHU = supply/return temp, filter DP, fan
  status/trip/command, damper AO, valve AO, freeze stat; Electrical panel = breaker status, trip,
  V/I via meter.
- `M:P` = field device per point, `Model` in `[Medium] [Device] [Range]` style
  (`[Water] [Ultrasonic Level Sensor] [0.25 m to 6 m tank]` -> `AX-UL-SEP380-2`). Only points needing
  a physical sensor get a device; statuses/commands wired from MCC do not.

**Exit criteria:** every EquipmentList row has a block; H:L are formulas; each block has a TOTAL row;
grand totals DI/AI/AO/DO/SP presented.

**Present:** per-type point counts and the grand total, plus the field-device list.

## Phase 3 - DDC List

Produce: the panel/equipment assignment matrix. Row 1 index, row 2 equipment names, row 3 `Total`,
row 4 `Assigned`, then one row per panel.

- Name panels by location: `DDC-GF-01`, `DDC-R-02`, `DDC-B1-03`...
- Group by **physical proximity first** (same plantroom/floor), then discipline.
- Leave headroom per the `Options` spare factor; split a panel rather than pack it to 100%.
- Keep each panel's IO within one enclosure's practical module count.
- **Hard limit: no DDC may carry more than 250 points.** Count physical (DI+AI+AO+DO) plus software
  points, and test it **with** the `Options` spare factor applied - a panel at 220 real points is
  already over once 20% spare is added. Split the panel rather than exceed the limit, and report each
  panel's loading against 250 when presenting the phase.

**Exit criteria (hard gate):** `Assigned` == `Total` for **every** column. A mismatch means equipment
unassigned or double-counted - fix before presenting, never present a mismatched matrix.

**Present:** panel list with per-panel equipment count and IO subtotal, plus explicit confirmation
that Assigned == Total across the board.

## Phase 4 - DDCSummary

Produce: per-panel IO totals + controller selection.

```
Q=Controller  R=Quantity  T=UI U=UO V=UIO W=DO X=AO Y=DI Z=RO AA=SP
```

- Panel IO rolls up from the Phase 3 assignment.
- Select from `Controllers` (hidden): automation stations `PXC7.E400S/M/L`, `PXC5.E24`; IO modules
  `TXM1.8D`, `TXM1.16D` (16 DI), `TXM1.8U` (8 universal DI/AI/AO), `TXM1.6R` (6 relay/RO); address keys
  `TXA1.K12` / `TXA1.K24`; power supplies `TXS1.12F10`, `TXS1.EF10`.
- One address key per 12 or 24 modules; one power supply per panel, more for heavy RO/DO loads.

**Exit criteria:** for every panel, module capacity (`T:AA`) >= IO demand **including spare**; station
`SP` capacity >= software points. Show the capacity-vs-demand comparison per panel.

**Present:** per-panel bill of controllers with the capacity check alongside.

## Phase 5 - DDCFullSummary

Produce: the consultant-facing expansion - every individual point with its field device and model,
alongside the controller list, per panel.

**Exit criteria (hard gate):** DDCFullSummary point totals == DDCSummary == IOSummary "All Systems".
If the three disagree, trace back to the source phase - **never patch the last sheet to make it match.**

**Present:** the three-way reconciliation as a table, plus the total point count per panel.

## Phase 6 - Ancillaries

Produce: `Workstation`, `DamperActuators`, `ValvesAndActuators`, `VFDs`/`MCCs`, `FCUs`/`RoomUnits`.

- **Workstation** - physical points (DI/AI/AO/DO -> Total) + software points (MODBUS, BACNET/IP,
  BACNET/MSTP, MBUS, KNX, Other, DXR2/RDF/RDG) roll up to `Total BA`, which sizes the server/licence:
  `CMD.06`, `CCA-CMPXL-BA`, `CCA-1000-BA`, workstation PC, 27" monitor, printers, UPS. Licence tier
  follows total BA points - re-check after any IO change.
- **DamperActuators** - `Unit Name | Actuator Title | Type | Signal | End Switch | Qty | Duct Size (m2)
  | Actuator Description | Part Number | Accessory`. Type `S.R.`/`N.S.R.`, signal `ON/OFF` or
  `Modulating`; torque follows duct area (`Options` defaults: N.S.R 2.5, S.R 2.5, F.S. 1.5 m2).
- **ValvesAndActuators** - sized by KVS (`Options` B51=2); threaded below the 50 mm threshold, flanged
  above; spring return + weather shield per `Options`; BV series e.g. `VKF46`; Globe preferred unless
  the spec says Ball/PICV.
- **VFDs / MCCs** - selected by HP; MCC in the BOQ only when `Options` says `Yes`.
- **FCUs / RoomUnits** - family follows the FCU protocol in `Options`: KNX -> `RDG100KN`/`RDG160KN`/
  `RDF600KN`; BIP -> `DXR2.E09/E10`; MSTP -> `RDB160BN`/`DXR2.M09/M10`; Modbus -> `RDF302`/`RDF300.02`;
  Standalone -> `RDG100`/`RDU340`.

**Exit criteria:** licence tier matches total BA points; every modulating damper and control valve in
the point list has a selected actuator.

## Phase 7 - BOQ, Breakdown, Cover Page

Produce: the priced offer.

**BOQ** - `Item | Category | Model number | DESCRIPTION | QTY | Manufacturer | Origin | unit list |
total list | unit net | margin | total net`. Quantities come **from the summary sheets, never retyped**.
Categories: Controller, Field Device, Valve, Actuator, VFD, Enclosure, Workstation/Software,
Installation material, Services. Margin factor sits in `M2` (e.g. `0.41`) and applies per line; pricing
resolves through `Product_Finder` / `Listprice` / `Pricelist` - **never hardcode a price over a formula**.

**Breakdown** - header (Project Name, Location, Reference, Division `Automation`, Customer, System
`BMS`, Total Value, SOW e.g. `Supply, T&C`, Status, Discount, Net Value) and cost sections (A. Foreign
Materials: Supplier / Selling Cost / Discount / F&C / Currency / Cost Value / Total Cost SR; then local
materials, labour, services). Check **Trading Margin** and **Gross Margin** land in the expected band;
flag to the user if they don't.

**Cover Page** - From `German Technical Services`, To (contact + company), Date, Ref as
`BMS-R01-<Mon><Year>`, subject, the standard 30-day-validity Siemens Desigo/BACnet paragraph, headline
prices (BMS supply/programming/T&C, plus options such as a redundancy server) in **SAR excluding VAT**,
notes and exclusions.

**Before quoting any margin:** `Breakdown!D23` (Siemens cost) reads from the **Product Finder**, not the
BOQ, so it does not follow a BOQ regeneration. Cross-check `BOQ!J112` (total cost) against
`Breakdown!I91`; if they differ, the Product Finder has not been refreshed with the new items and the
margin shown is wrong. Refreshing it is an estimator step - no macro does it. Never rewire D23.

**Exit criteria - the full final audit:**

1. `EquipmentList` qty == `DDC List` row 3 `Total` == row 4 `Assigned`, per column.
2. IOSummary "All Systems" == DDCSummary == DDCFullSummary point totals.
3. Every panel's module capacity >= IO demand including the `Options` spare factor.
4. Workstation licence tier matches total BA points.
5. Every BOQ quantity traces to a summary sheet; no orphan or hand-typed lines.
6. Margins in band; Cover Page total == Breakdown Net Value == BOQ total.
7. Scope statements match the SOW - supply only vs supply + install vs T&C.

Present the audit as a checklist with pass/fail, then the headline price.

---

## Workbook mechanics

- Macro-heavy with many hidden lookup sheets. Edit in Excel (or via COM) so formulas and macros
  survive - a plain XLSX library **strips the VBA project** and breaks all selection logic.
- Read-only inspection without Excel:
  `powershell -File references/dump_sheets.ps1 -Xlsm "<file.xlsm>" -OutDir out -MaxRows 200`
- Reference estimations in `~/OneDrive/Desktop/GTS/BMS_PROJECTS/` (Ajyad Tower, YALJ, Prince Mansour,
  RX Premium Hub). **YALJ is the best full worked example** - all 35 sheets filled.

## `_ESTIMATION_STATE.md` template

Create this in the project folder at Phase 0 and update it at the end of every phase.

```markdown
# BMS Estimation State - <Project Name>

- **Workbook:** <file name>   **Template version:** V27.1 / V26.2
- **Scope:** <one sentence>
- **Last updated:** <YYYY-MM-DD>

## Phases
- [ ] 0 Survey & Options - <status>
- [ ] 1 EquipmentList - <status>
- [ ] 2 IOSummary - <status>
- [ ] 3 DDC List - <status>
- [ ] 4 DDCSummary - <status>
- [ ] 5 DDCFullSummary - <status>
- [ ] 6 Ancillaries - <status>
- [ ] 7 BOQ / Breakdown / Cover Page - <status>

Status values: NOT STARTED | IN PROGRESS | AWAITING REVIEW | DONE

## Decisions
| # | Phase | Question | Answer | Date |
|---|---|---|---|---|

## Open questions
| # | Phase | Question | Proposed default |
|---|---|---|---|

## Key figures (update as they firm up)
| Metric | Value |
|---|---|
| Equipment families | |
| Total DI / AI / AO / DO / SP | |
| DDC panels | |
| Total BA points | |
| BOQ total (SAR excl. VAT) | |
```
