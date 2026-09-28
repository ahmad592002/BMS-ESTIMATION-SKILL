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

<!-- New entries go below this line -->
