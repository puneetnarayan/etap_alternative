# IMCC design (integrated with PMCC)

Page: **IMCC Design** in the sidebar (`/imcc`). Code: `src/core/project.ts` (model, rules, calculations), `src/core/imcc-import.ts` (reads the reference layout), `src/core/imcc-workbook.ts` (Excel), `src/app/imcc/` and `src/app/projectStore.ts` (UI, storage). Tests: `src/core/__tests__/imcc.test.ts`.

## Data model (one source of truth)

`DesignProject { version, name, pmccs[], imccs[], generated[] }`, kept in the browser (localStorage) and as a downloadable `.json`. Projects without IMCCs (older data) open with defaults (`migrateProject`).

- `Pmcc { id, tag, fileName, sheet }`: the PMCC load list as uploaded. Its rows are the **manual** outgoing feeders and are never changed by IMCC work.
- `Imcc { id, tag, description, voltageV, incomer1 {pmccId, feederRef}, incomer2 {pmccId, feederRef}, normallyOn, rule, loads[] }`.
- `ImccLoad { id, section 1|2, description, kw, status W|S|SP, method DOL|RDOL|VFD|PWR FDR, group, note }`.
- `GeneratedFeeder` (id `imcc:<imccId>:<1|2>`): the PMCC outgoing feeder of an IMCC incomer. **Derived** from the IMCCs by `syncLinks`, so it cannot be duplicated or contradict the IMCC. Fields: PMCC, IMCC, incomer, description (`<IMCC tag> Incomer n`), destination, feeder reference, status, working current / kVA of that bus section, design current, calculated PMCC feeder rating.

All links use ids, never descriptions. Every operation (`saveImcc`, `duplicateImcc`, `deleteImcc`, `addPmcc`, `deletePmcc`, `renamePmcc`) is a pure function returning a whole new project or a list of errors; the UI stores the result only on success, so a failure leaves the previous consistent project.

## Operating philosophy

Default rule: exactly one incomer ON, the other incomer OFF, bus coupler OFF. `validateOperatingState` rejects both incomers ON, no incomer ON, and the coupler ON (also with one incomer ON: one incomer is never assumed to supply both sections). The rule is configurable, but a rule different from the default is refused unless a named authoriser is entered (an explicit engineering change).

## Calculation flow

1. Status: `W` / `w` working, `S` standby, `Spare` full spare (`SP`). Anything else is flagged and not counted. Only W counts in the normal working load; S and SP do not.
2. Load current: the application's existing `currentFor` (MCC tab assumptions: pf DOL 0.7, VFD 0.9, feeder 0.7; DOL efficiency 0.93 provisional; 433 V loads at 433 V). kVA = √3 x V x I / 1000.
3. Section (bus section 1 / 2) working current and working kVA = sum over W loads of that section. Loads without a section are flagged, never allocated.
4. Incomer 1 design rating: section 1 working current x design multiplier (1.25) → the existing `selectBreaker` ladder. **Incomer 2 and the bus coupler = Incomer 1's calculated rating**, always (a warning is raised if section 2 would need more).
5. PMCC outgoing feeder for incomer n: current = section n working current, design = x 1.25, rating from the same `selectBreaker` on its own current (not copied from the incomer rating). Status W for the incomer normally ON, S for the other, so the two alternative sources are never added together as normal demand.
6. PMCC summary: working kVA counts W only; standby kVA and spare counts are separate.
7. `LCE kVA` (working kVA x 0.8, reference workbook) is reported only; its meaning is open (OPEN_ITEMS.md #8).

## Excel

- **IMCC workbook** (`Excel` on an IMCC): a sheet in the reference layout (S. NO., description, kW, current, W/S/SP, method, group, current, kVA, working kVA, working current, LCE kVA, design current, breaker, notes), formulas for the derived columns and section totals, banners naming the upstream PMCC and feeder, incomer / coupler ratings, overall totals, issues; a `Notes` sheet holds the assumptions the formulas read. Spare rows are kept and shown as `SP`.
- **PMCC workbook** (`Excel` on a PMCC): the existing MCC Load List workbook, built from the PMCC list with one extra row per linked IMCC incomer (`Source` = Auto from IMCC, `Destination`), plus an `IMCC links` sheet. A PMCC without IMCC links is exported exactly as before.
- Both exports come from the same project data. The formulas were checked against a LibreOffice recalculation (section totals equal the engine to 12 digits).

## Importing the reference layout

`parseImccSheet` reads the header row (S. NO., KW, W/S/Spare, Method of starting, Group), the description (column after S. NO.), and `INCOMER - 1/2` rows as section starts. The reference's own current column (415 V, 1.73) is not imported. Rows with no status are imported and flagged (the reference sheets have several).

## Running and testing

`npm run dev`, open **IMCC Design**; add a load list in the sidebar, add two PMCCs, create an IMCC, import an ECR sheet, Save. `npx vitest run` (IMCC tests A to J are in `imcc.test.ts`), `npx tsc --noEmit`, `npx next build`.
