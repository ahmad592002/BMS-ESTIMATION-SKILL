# Learned rules - BMS estimation

Read this file at the start of **every** invocation of the `bms-estimation` skill. Rules here
**override** the generic guidance in `SKILL.md` wherever the two conflict.

When the estimator changes something you produced: **apply it, then ask why** (spec / GTS standard /
client preference / project condition / your error / commercial), and let their answer set the `Scope`.
No answer given -> record with `Reason: not given` and `Scope: this project`, and never ask twice.
See the "Self-upgrade: ask WHY, then record" section of `SKILL.md` for the full protocol.

Edit an existing entry rather than adding a near-duplicate; a newer reason overrides an older one.
Confirm in one line: `Learned: <title> (scope: <scope>).`

Entry format:

```markdown
## <short rule title>
- **Phase:** <number/name>   **Date:** <YYYY-MM-DD>   **Project:** <name>
- **Rule:** <the rule, imperative, one or two lines>
- **Change:** <what was changed - from -> to>
- **Reason:** <their answer, in their terms> | not given
- **Reason type:** spec | GTS standard | client preference | project condition | my error | commercial
- **Scope:** always | this client | this project type | this project
- **Also changed in SKILL.md:** <section, or "no">
```

---

## Estimation runs one phase per invocation
- **Phase:** all   **Date:** 2026-09-28   **Project:** (skill setup)
- **Rule:** Complete one phase, present it for review, then stop. Never chain two phases in one turn,
  and never mark a phase DONE without the estimator's approval or an explicit "next".
- **Change:** skill rewritten from a single end-to-end run -> gated phase-by-phase run
- **Reason:** "i want the skill like every time i use it it finish one phase per time like first get
  all equipement i check if i find thing not good i ask"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** whole structure (phase gates, state file)

## Ask why a correction was made, then scope the rule accordingly
- **Phase:** all   **Date:** 2026-09-28   **Project:** (skill setup)
- **Rule:** After applying any substantive correction, ask once whether it came from the spec, GTS
  standard practice, this client, this project, your error, or a commercial call - and set the recorded
  rule's `Scope` from that answer. Never generalise an unexplained edit.
- **Change:** silent recording of corrections -> ask the reason, then record with a reason type
- **Reason:** "and it ask me if i update why i update and based on my answer update the skill"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** "Self-upgrade: ask WHY, then record"

## Self-upgrades are published as pull requests, never pushed to main
- **Phase:** all   **Date:** 2026-09-28   **Project:** (skill setup)
- **Rule:** After recording a rule in this file, batch that phase's learned rules onto a
  `learn/<phase>-<slug>` branch and open a PR against `main` of
  ahmad592002/BMS-ESTIMATION-SKILL. Never commit to `main` directly and never self-merge - the
  estimator reviews. The rule applies locally at once; a failed push never blocks the estimation.
- **Change:** rules saved only to the local skill folder -> local save plus a PR for review
- **Reason:** "the skills in it update should pull request to the repo"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** "Self-upgrade: ask WHY, then record" -> new subsection
  "Then publish the change as a pull request"

<!-- New entries go below this line -->

## No DDC panel may exceed 250 points
- **Phase:** 3 DDC List / 4 DDCSummary   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Cap every DDC at 250 points - physical (DI+AI+AO+DO) plus software points - and test the
  cap with the `Options` spare factor applied, not on raw counts. Split the panel rather than exceed
  it. Report each panel's loading against 250 whenever Phase 3 or 4 is presented.
- **Change:** panel sizing judged only by module count and enclosure practicality -> explicit 250-point
  ceiling checked including spare
- **Reason:** "you should know every ddc should not have more then 250 point"
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** Phase 3 - DDC List (panel loading rules)

## Standardise field devices on parts already used in other GTS projects
- **Phase:** 2 IOSummary   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Before pricing, compare every field device against the other BMS_PROJECTS workbooks. A
  model appearing in only one project is a candidate for substitution by the house-standard part for
  the same duty. Match by FUNCTION: status/proof points take D.P.S. switches (QBM81-x), measured
  values take transmitters (QBM3020-x, QBE3000-D16). Never swap where the SOW or physics dictates the
  part - raise it as a question instead.
- **Change:** SDA-only devices -> house standard: 28x QBM3020-25 -> QBM81-5 (fan airflow proof),
  44x QBM3020-10 -> QBM81-10 (filter DP), 5x PL-FD113 -> QBE3000-D16; kept 8x QBM3020-5 for duct
  static pressure under VFD control
- **Reason:** "in the filed device in the io summary i think some device are not used in any other
  project then change them to another thing used if is possible"
- **Reason type:** (pending - GTS standard or this tender?)
- **Scope:** always (provisional - confirm)
- **Also changed in SKILL.md:** no

## The model column in IOSummary is a VLOOKUP - edit the description, not the model
- **Phase:** 2 IOSummary   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Column O is `IFNA(VLOOKUP($N<row>, FieldDevices!A:B, 2, FALSE), ...)`. To change a device,
  write the exact `FieldDevices!A` description into column N and let O resolve; FieldDevices quantity
  counts then update themselves. Never overwrite O - it destroys the formula. Changing a point between
  status and measurement also means moving its 1 between column C (DI) and D (AI); H:L are formulas
  (`C:G * $A<block>`) and must never be typed.
- **Change:** n/a - workbook mechanics learned while editing
- **Reason:** discovered when swapping field devices; writing O directly would have broken the lookup
- **Reason type:** my error (avoided)
- **Scope:** always
- **Also changed in SKILL.md:** no

## Write to the workbook via Excel COM, with events disabled
- **Phase:** all   **Date:** 2026-09-28   **Project:** SDA SCITECH Khobar
- **Rule:** Excel COM is available on this machine. Back the workbook up first, attach to the running
  instance if the file is already open (`GetActiveObject("Excel.Application")`), set `EnableEvents =
  $false` so no `Worksheet_Change` macro fires mid-write, edit, `CalculateFullRebuild()`, verify the
  gate, then Save. Never write these workbooks with a plain XLSX library.
- **Change:** content handed over as tables to paste -> written directly into the workbook
- **Reason:** Excel 16.0 present; the workbook was open in a live session during the edit
- **Reason type:** GTS standard
- **Scope:** always
- **Also changed in SKILL.md:** no
