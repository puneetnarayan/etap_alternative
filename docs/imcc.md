# IMCC design

Page: **IMCC Design** in the sidebar (`/imcc`, its own colour). Code: `src/core/imcc.ts` (design script, rules, feeders), `src/core/imcc-import.ts` (reads the IMCC list), `src/core/intake.ts` (PMCC / IMCC by file name), `src/core/imcc-workbook.ts` (Excel), `src/app/imcc/`. Tests: `src/core/__tests__/imcc.test.ts`.

## Inputs: two files, identified by name

Both inputs have the same layout (first sheet; header row with Description, W/S, Type (RDOL/DOL/VFD/SS/Feeder), Amps rating (A), kW rating (kW), Voltage level (V), optional Duty/standby group). Samples: `public/PMCC_sample_load_list.xlsx` and `public/IMCC_sample_load_list.xlsx` (linked from the pages).

- The **file name must start with PMCC or IMCC** (any case) so the data is identified: a PMCC file goes to the PMCC design (transformer sizing, MCC, BOM), an IMCC file to the IMCC design. Any other name is refused with a message.
- Either drop zone (sidebar, or the IMCC page) takes one or both files. If **both** are dropped together, a pop-up asks for confirmation; then each file is processed by its own logic. Two files of one kind at once are refused. Cancelling changes nothing.
- Entered amps are used in preference to kW (as everywhere in the app). W, w = working; S = standby; Spare = spare feeder (SP); a blank or other status is flagged and the row is not counted or placed; a number that cannot be read (e.g. "-") is flagged; a row at another voltage than the IMCC's (433 V) is not placed.

## Sidebar

Grouped tabs: **Input files** (drop zone showing the PMCC and IMCC file names, the **Use sample files** button and links to both samples), **General** (Home), **PMCC design** (Transformer sizing, MCC Load List / Breaker Selection, BOM; blue), **IMCC design** (IMCC Design and the Automatic PMCC feeders switch; teal) and **Analysis** (Load flow). **Use sample files** loads `PMCC_sample_load_list.xlsx` and `IMCC_sample_load_list.xlsx` (no pop-up: the click is the confirmation) and opens the Transformer sizing output; the IMCC tab shows the IMCC output. Field boxes, selects and buttons use pastel colours (`src/app/globals.css`).

## Stand-alone, with one switch

The IMCC design is separate from the other tabs and is linked to them only by the **Automatic PMCC feeders** switch in the sidebar (under the IMCC entry; ON by default; it blinks whether ON or OFF).

- **ON**: each IMCC incomer with a working load becomes an outgoing feeder row of the PMCC load list. Transformer sizing, MCC and BOM see the list with these rows added, so the transformer loading is recalculated and the number of transformers may change (the sizing page shows "+ N IMCC incomer feeders"). With the two sample files the sizing goes from 11 to 12 x 2000 kVA.
- **OFF**: the IMCC design is independent; the PMCC load list is exactly the uploaded file. This is not advisable: the PMCC loading then leaves out the IMCCs.

The switch state is kept in the browser (localStorage), the two lists in the browser tab (sessionStorage). The uploaded PMCC file is never modified: the rows are added on the fly (`appendFeedersToSheet`, a pure function of the uploaded sheet and the feeders, so there are no duplicates).

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

## Output design (display and Excel)

Both show each IMCC in **bus order, as the transformer report does**: Incomer 1 (cyan) | its loads | Bus coupler (light yellow) | the loads of section 2 | Incomer 2 (cyan). Colours are those of the other reports: light-orange headers, W lavender, S pink, spare (SP) gold, DOL / SS, VFD and Feeder pastels, loading green when within and red when over 100 % of the rating.

- **Loading block** at the top of each IMCC: Incomer 1, Incomer 2, bus coupler with normal state (ON / OFF), working kVA and A, design A, calculated rating and **loading % of the rating** (Incomer 2 is loaded against Incomer 1's rating, which it follows).
- **Bus layout** below: the incomer rows carry the section totals; the loads show method, W/S/SP, kW, current, kVA, working kVA / A, design A and breaker.
- **Excel** (`Download Excel (all IMCCs)`): Summary (IMCCs, ratings, loading %, checks, the generated PMCC feeders when the switch is ON), one sheet per IMCC (loading block, colour legend, bus layout, formulas for currents, kVA and totals, conditional red when over 100 %), and Notes (assumptions the formulas read). Spare rows are kept and shown as `SP`. The formulas were checked against a LibreOffice recalculation.
- On screen the loading rows can be collapsed; the operating-state buttons check a state against the rule.

## Running and testing

`npm run dev`, open **IMCC Design**, drop `IMCC_sample_load_list.xlsx` (and `PMCC_sample_load_list.xlsx`); toggle the sidebar switch and watch **Transformer sizing**. `npx vitest run`, `npx tsc --noEmit`, `npx next build`.
