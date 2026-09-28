# Phase 0 - Survey and setup

Produces: an inventory of the input files, the scope in one sentence, a working copy of the template,
and the `Options` sheet filled and approved.

## Steps

1. **Read the folder before touching Excel.** Typical inputs: equipment schedule (xlsx), BMS SOW /
   scope pdf, specification booklet (often Arabic: `كراسة الشروط والمواصفات`), mechanical + electrical
   drawings, a client BOQ to be priced, contract form.
2. **Read the SOW and spec** for: scope boundary, protocol, brand restrictions, spare requirement,
   integration list, workstation/server requirement, warranty and T&C obligations.
   - **Read the client's own جدول الكميات (bill of quantities) line by line.** It fixes quantities you
     would otherwise estimate - panel count, sensor counts, valve sizes, VFD ratings, workstation PCs.
     Every later phase should reconcile against it.
   - Search the PDF text for scope words before assuming something is included:
     `pdftotext -enc UTF-8 "<sow>.pdf" - | grep -inE "server|workstation|MDB|generator|UPS|meter"`
3. **Copy the newest master** from `~/OneDrive/Desktop/GTS/BMS_PROJECTS/BMS Template 2026 V03.xlsm`
   into the project folder as `<Project> - GTS offer.xlsm`. **Never edit the master in place.**
   - **V27.1** = `Product_Finder_SI_B_AUT_V27.1` + `Listprice_V27.1_AUT_A_SP` (default).
   - **V26.2** = `Product_FinderV26.2` + `Listprice_V26.2_HVAC_A_SP` (only if the client locked it).
4. **Set `Options` first** - every hidden lookup sheet reads from it, so changing it later re-selects
   hardware underneath finished work.

| Option | Typical | Note |
|---|---|---|
| Spare in DDC | `0.2` | 20% spare IO; raise if the spec demands |
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

**Check `Options` against the devices actually selected.** If `FCU Communication Protocol` says `KNX`
but the room units are `RDF440BN` (BACnet MS/TP), one of the two is wrong - raise it. Changing B13
re-runs the RoomUnits selection, so never flip it silently.

5. **Create `_ESTIMATION_STATE.md`** from `references/state-template.md`.

## Exit criteria

- Every input file accounted for and identified.
- Scope stated in one sentence, traceable to the SOW.
- Client BOQ quantities extracted and recorded in the state file for later reconciliation.
- `Options` presented as a table and **approved by the estimator** - that approval is this phase's
  review gate.
