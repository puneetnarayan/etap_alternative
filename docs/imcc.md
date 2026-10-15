# IMCC design

Page: **IMCC Design** in the sidebar (`/imcc`, its own colour). Code: `src/core/imcc.ts` (design script, rules, feeders), `src/core/imcc-import.ts` (reads the ECR layout), `src/core/imcc-workbook.ts` (Excel), `src/app/imcc/`. Tests: `src/core/__tests__/imcc.test.ts`.

## Stand-alone, with one switch

The IMCC tab has **its own load list** (drop zone on the page, kept for the browser tab). It does not read the sidebar load list and changes nothing in the other tabs, except through the **Automatic PMCC feeders** switch in the sidebar (under the IMCC entry; ON by default; it blinks whether ON or OFF).

- **ON**: each IMCC incomer with a working load becomes an outgoing feeder row of the PMCC load list. Transformer sizing, MCC and BOM see the list with these rows added, so the transformer loading is recalculated and the number of transformers may change (the sizing page shows "+ N IMCC incomer feeders").
- **OFF**: the IMCC design is independent; the PMCC load list is exactly the uploaded file. This is not advisable: the PMCC loading then leaves out the IMCCs.

The switch state and the generated feeders are kept in the browser (localStorage). The uploaded PMCC file is never modified: the rows are added on the fly (`appendFeedersToSheet`, a pure function of the uploaded sheet and the feeders, so there are no duplicates).

## Design script (`designImccs`)

1. Only **working (W / w)** loads are used for the design. `S` (standby) and `Spare` (shown as `SP`) are never counted. Any other status is flagged and the row is not counted or placed.
2. Current per load: the application's `currentFor` (pf DOL 0.7, VFD 0.9, feeder 0.7; DOL efficiency 0.93 provisional; 433 V loads at 433 V). kVA = √3 x V x I / 1000. **SS = soft starter**, calculated like a DOL motor (assumption, OPEN_ITEMS.md #12).
3. Working loads are taken **in list order** and fill bus sections: a new section starts when adding the next load would push the section's **design current** (working current x 1.25) above **1000 A**, the incomer limit. A single load above the limit is flagged and placed alone.
4. Two sections (Incomer 1, Incomer 2) make one IMCC (`IMCC-01`, `IMCC-02`, ...). The number of IMCCs follows from the loads.
5. S and SP rows are placed with their group's section, else with the preceding working load; they carry no working load.
6. Incomer 1 rating: `selectBreaker` (existing ladder) on its section's design current. **Incomer 2 and the bus coupler equal Incomer 1's calculated rating**; a warning is raised if section 2 would need more. These are calculated design ratings, not nameplate or installed-device ratings.
7. `LCE kVA` (working kVA x 0.8, as in the reference workbook) is reported only (OPEN_ITEMS.md #8).

## Operating philosophy

Exactly one incomer ON at a time; the other incomer and the bus coupler OFF (`validateOperatingState`; the operating-state buttons on each IMCC row show the message). Both ON, none ON, or the coupler ON (also with one incomer ON) are errors. Changing the rule is an engineering change and needs a named authoriser (`OperatingRule.authorisedBy`).

## Generated PMCC feeders (`imccFeederRows`)

One row per incomer that has a working load: description `IMCC-01 Incomer 1`, Type Feeder, Amps = the section's working current, 433 V. **Incomer 1 = W, Incomer 2 = S** (alternative sources, never added together as demand) and both carry the IMCC tag as **group**, so the transformer sizing places them on different transformer sections (its existing duty/standby group rule). The feeder is sized from the section's own current, not from the incomer breaker rating (the MCC tab then selects its breaker with the usual rules).

## Excel (`Download Excel (all IMCCs)`)

One workbook, same look as the other workbooks of the application: **Summary** (IMCCs, ratings, checks, the generated PMCC feeders when the switch is ON), one sheet per IMCC in the ECR layout (S. NO., description, kW, current, W/S/SP, method, group, current, kVA, working kVA, working current, LCE kVA, design current, breaker, notes; formulas, section totals, incomer / coupler ratings; spare rows kept and shown as `SP`), and **Notes** (the assumptions the formulas read). Each incomer banner says whether it is supplied from a generated PMCC outgoing feeder (switch ON) or not linked (OFF). The formulas were checked against a LibreOffice recalculation.

## Import (`parseImccSheet`)

Reads the ECR layout: header row with S. NO., KW, W/S/Spare, Method of starting, Group; description in the column after S. NO.; `INCOMER - 1/2` rows are ignored (the script decides the sections). The reference's own current column (415 V, 1.73) is not imported. Rows with no status are flagged, not guessed.

## Running and testing

`npm run dev`, open **IMCC Design**, drop an ECR sheet; toggle the sidebar switch and watch **Transformer sizing**. `npx vitest run`, `npx tsc --noEmit`, `npx next build`.
