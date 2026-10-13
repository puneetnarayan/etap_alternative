# MCC Load List / Breaker Selection

Sidebar entry **MCC Load List / Breaker Selection** (route `/mcc-load-list`). The sidebar has a hamburger button and starts **expanded** on every fresh page load (its state is not saved).

## Behaviour
- Clicking the entry opens the page **and** builds the workbook from the last load list used in this browser tab (a list dropped on the Transformer sizing page, or on the MCC page, is kept in `sessionStorage`). Only the click triggers it; re-renders do not.
- No list yet: the page says so and offers a drop zone (dropping a file builds and downloads the workbook straight away) and the sample file link. Nothing is invented.
- It is a web app, so "open in Excel" is a **download** (`MCC_Load_List_Breaker_Selection_<date-time>.xlsx`); browsers cannot launch desktop Excel themselves, and the page says so. A loading state and errors are shown in the sidebar and on the page.
- Built with ExcelJS (MIT), because write-excel-file has no filters, cell notes or drop-downs. It is loaded only when the workbook is generated.

## Results on the page
The page shows the same figures as a table (colours as on the Transformer sizing page: orange headers, W/S and type pastels; breaker cell green = preliminary selection, yellow = ACB to verify, light red = engineering review) with a summary line and the assumptions, for the last load list used in the tab. The rules live in `mccRows` (`src/core/mcc-breaker.ts`), shared with the tests; the workbook has its own formulas, checked against the same rules.

## Workbook
**Load list**: your original columns first, unchanged, then added columns (blue = input data, green = calculated, yellow = preliminary selection). Header row frozen and filtered; formulas stay formulas (they follow the yellow assumption cells on the Notes sheet); header cells carry notes; the two breaker-type columns have a drop-down (MCCB / EDO ACB / Engineering review required). **Notes**: assumptions, formulas, rules and the final checks.

Added columns: load type, voltage, kW and rated current as read; calculation voltage; power factor; efficiency; entered current; calculated current; current used and its source; design current; feeder protection type, indicative rating, selected breaker; overload relay setting; incomer type, rating; CT ratio and use; cable design current; cable size / runs and length (linked to source columns of that name if the list has them, otherwise blank inputs); engineering notes.

## Rules
- `I = kW × 1000 / (√3 × V × PF)`; DOL motors in kW also ÷ efficiency (kW taken as motor output). PF: general feeder 0.7, DOL 0.7, VFD 0.9. Efficiency is a **provisional 93 %**, editable on the Notes sheet and per row. The 433 V level is calculated at 420 V, as in the transformer tool (editable).
- A valid rated current (the Amps column) is used in preference to a calculation. Missing kW / voltage / DOL efficiency leaves the calculated cells blank and the notes say what is missing; no zeros.
- `Design Current = Current used × 1.25` (editable screening factor, not a final device rating).
- Breaker: up to 400 A → next of 32, 63, 100, 160, 250, 400 A MCCB; above 400 A and below 500 A → **engineering review** (no undersized 400 A MCCB); 500 A to 630 A → 630 A EDO ACB; above 630 A → 800 / 1000 / 1250 A EDO ACB flagged for verification; above 1250 A → review. A 630 A MCCB is never selected.
- Incomer columns are filled only for rows whose description contains "incomer". Where such a row is an 800 A rating, it is standardised to a **1000 A EDO ACB** (only there; other 800 A values are not converted). Incomers are not sized from the sum of feeder ratings.
- CT ratio only on incomer rows: smallest standard primary not below the incomer rating, `/5 A`, indicative. Not assigned to motor feeders.
- Everything is preliminary and needs engineering approval; the Notes sheet lists the final checks.

## Interpretations to confirm
- The efficiency is applied to **DOL** motors only (the brief names DOL for it). The transformer tool applies 95 % to VFD motors too, so VFD currents here are not comparable with its VFD kVA.
- Overload relay setting = the motor current used (DOL only); VFD blank (set in the drive).
- Cable design current = design current.
- Between 400 A and 500 A is review for feeders and incomers alike. A 630 A–800 A requirement picks an 800 A ACB (the 800 → 1000 A standardisation is applied only to an existing 800 A ACB incomer, as the brief says).
- Verified: formulas recalculated in LibreOffice match the TypeScript rules for the sample; the file has not been opened in desktop Excel here.
