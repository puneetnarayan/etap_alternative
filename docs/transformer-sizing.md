# Transformer sizing – rules as implemented

Source: the user's load-distribution brief (Rev 0) plus follow-up answers. Code: `src/core/transformer-sizing.ts`.

## Transformer-wise sheet and table

The Excel report has a **Transformer-wise** sheet (second, after Summary) and the page shows the same table under the result (`src/core/transformer-wise.ts`):

1. **Per transformer** (normal running): number of loads, kW (input), kVA, current on the secondary (kVA x 1000 / (√3 x V)), rating, rated current, **% loading** (green within the normal limit, red over) and the margin to the limit; then a **Total** row (sums of loads, kW, kvar, kVA and current; installed kVA; average loading = total load / installed).
2. **Why this number of transformers is not an excess** (Excel: on the separate **Justification** sheet, which also gives the basis used - limits, standby rule, factors - and the two tables below; on the page, under the table): utilisation of the installed capacity; what the number taken can carry within both limits and the margin over the load (also as a share of what one transformer adds); the next smaller number tried and by how much it is short; smaller ratings at this number and why they do not fit; and whether this combination has the smallest installed capacity of all that fit.
3. **Numbers of transformers tried**: each count with the smallest rating that fits, installed kVA, total load, average normal and one-out loading, capacity within the limits, margin (negative = short), result; the number taken is marked.
4. **Every number and rating tried**, smallest installed capacity first (Excel only).

## Ratings, factors and the editable feeder block (latest)

- **4000 kVA is for the 690 V system only.** At 433 V the ratings are 2000 and 2500 kVA (`ratingsForVoltage`); the 4000 kVA radio is disabled unless 690 V is the only level selected, and the core never tries 4000 kVA at 433 V.
- **Diversity factor (industrial loads)**, editable, default **1**: the load considered is the sum of the loads of the list x this factor (0.8 = 80 % of the sum). It applies to the loads of the list, not to the fixed feeder block. Confirmed: it is a multiplier (a value below 1 reduces the load), not the textbook ratio above 1 that divides.
- **Motor load factor**, editable, default **100 %**: the share of the rated motor power absorbed; motors only (kW rows and DOL / VFD rows given in amps), not feeders.
- **Standby (S)** rows are never counted, also inside a group that has working members: the working members are sized, the standby ones are listed as excluded.
- **PMCC/IMCC feeder block**: a table on the page (feeder rating, installed, counted in the load), default 6 x 63 A (4 counted), 6 x 32 A (4), 6 x 16 A (4), 10 x 6 A (6); rows can be edited, added, removed and restored to the defaults. Counted is limited to installed; a blank counted means all.

## Inputs
Excel `.xlsx`, first sheet, header row required (column order does not matter):
Description · Type (DOL / VFD / Feeder) · Amps rating · kW rating · Voltage level (433 or 690 V).
Optional: **Working/Standby (W/S)** column (any header containing "working", "standby" or "duty", or "W/S"; values W / Working / Duty or S / Standby), shown as W or S in the results and report. **Only working loads are counted: standby (S) rows are left out of the calculation** (they would add transformers that are not needed) and are listed as excluded ("Standby (S): not counted in the calculation") in the Issues list and sheet. The MCC and BOM tabs still select a breaker for every row.
Optional: **Duty/standby group** (any header containing "group" or "pair"): loads with the same group name are placed on different transformers' bus sections.
Sample: `public/PMCC_sample_load_list.xlsx` is the file users download from the page (supplied by the project owner; replace the file to change it). The tests for it are in `src/core/__tests__/public-sample.test.ts`. The earlier hand-built small plant is kept only as a test fixture, `src/core/__tests__/fixtures/small_plant_load_list.xlsx`.

## Voltage selection
Radio buttons choose the voltage level: **433 V**, **690 V**, and **Both** (only when the file has loads at both levels; a level the file does not have is not offered). Before a file is loaded both levels are shown. With Both, each level gets its own result section and its own Excel report button; a new load list starts on 433 V (or 690 V if that is all the file has). The assumptions (efficiency, power factors, voltage used for amps at 433 V) are always visible and editable.

## Load classification and kVA
| Given | Treated as | kVA |
|---|---|---|
| kW | Motor (DOL pf 0.7, VFD pf 0.9; other/blank type → DOL, with a warning) | input kW = kW ÷ 0.95; kVA = input kW ÷ pf |
| Amps only, Type DOL/VFD | Motor | √3 × V × A ÷ 1000 (V = 433 for the 433 V level) |
| Amps only, anything else | Outgoing feeder (pf 0.7) | √3 × V × A ÷ 1000 (V = 433 for the 433 V level) |
| Neither | Error – row not counted | – |

**Voltage for amps-only loads:** loads at the 433 V level are converted at **433 V** (kVA = √3 × 433 × A ÷ 1000), loads at 690 V at 690 V. This was 420 V earlier; calculations now use the nominal voltage even where a project document says 415 V or 420 V. The voltage used at the 433 V level is an editable assumption on the page (and `calculationVoltageV` in the options: a level not listed uses its own value).
kvar = input kW × tan(arccos pf), so kVA² = kW² + kvar². For amps-only loads kW = kVA × pf.
If kW is given, any amps value on the same row is ignored. Working loads count at 100 %; standby (S) rows are not counted.
Rows whose voltage differs from the selected transformer are excluded (listed on screen); a blank voltage is assumed to match, with a warning.

## Transformer selection
Transformers: 6600 / 433 V or 6600 / 690 V (chosen in the UI). The rating is chosen with radio buttons (2000, 2500 or 4000 kVA; default 2000). All transformers in a group are identical.

**The number of transformers is not fixed.** The tool starts with a board of **3 transformers** (3 incomers + 2 bus couplers; "start with 2" is an option for small loads) and adds transformers whenever the loads cannot be balanced within the limits:
- More transformers are built only from **independent sets of 3, with a set of 2 at the end for the remainder**: 3, then 5 (3+2), 6 (3+3), 8 (3+3+2), 9 (3+3+3), 11, 12, … up to **20**. A set of 2 has 2 incomers + 1 bus coupler, a set of 3 has 3 incomers + 2 bus couplers (one coupler between neighbouring incomers) and there is **no coupler between sets**; couplers = transformers − sets. 4 (2+2), 7 (3+2+2), 10 … carry no more than one transformer fewer, so they are skipped.
- **Limits:** normal running ≤ **80 %** of the rating on every transformer's section; with one transformer out, only the other one or two in the same set carry that set's load and each stays ≤ **90 %**. For rating R a section carries 0.8R, a set of 2 carries min(0.8·2·R, 0.9·1·R) = 0.9R and a set of 3 carries min(0.8·3·R, 0.9·2·R) = 1.8R (2000 kVA: 1800 and 3600).
- **"Fits" means the loads, placed whole, can really be balanced.** It is not enough that the total is below the capacity: the loads are placed on the transformers' sections (below) and every section and every set must be within its limit. If not ("loads cannot be balanced within the limits"), the next number of transformers is tried. For example 5 loads of 1000 kVA on 2000 kVA transformers: 5 transformers (3+2) carry 5400 kVA in total, but only 4 such loads can be placed (3 on the set of 3, 1 on the set of 2), so 6 (3+3) is used.
- **A load that is too big for any section** (above 80 % of the largest rating considered, e.g. a 1700 kVA feeder on 2000 kVA transformers) cannot be helped by more transformers; no number is added, the load is named and the page suggests a rating that would work.
- For each number of transformers the smallest rating that fits is chosen (if none fits, the largest rating is shown, marked "does not fit"). With one rating chosen the recommendation is the first number of transformers, in the order above, that fits.
- Total kVA: arithmetic sum of each load's kVA (confirmed by the user). It is slightly above √((Σ kW)² + (Σ kvar)²).
- The 80 % / 90 % limits are fixed (not editable). Power factors and efficiency are editable in the UI.
- Only the number started with and the final (fitting) number get a full card / Excel sheet; every number tried is listed in a "numbers of transformers tried" table. Sections are named T1, T2 … for one set and S1-T1, S1-T2, S2-T1 … for several.

## Current
Each load row has a **Current (A)** column right after kVA: I = kVA × 1000 ÷ (√3 × V), with V = 433 for the 433 V level and 690 for the 690 V level, and the kVA from the rules above (so the power factors DOL 0.7, VFD 0.9, other 0.7 and the 95 % efficiency are already in it). A load given in amps shows those amps back. It is shown on the page, on the Loads sheet and in the bus layout (also for each incomer, from the section's kVA); it is never summed. PMCC/IMCC feeders show their rating when counted and 0 when not counted; spares and couplers have none.

## Load distribution
Whole loads are never split. They are first placed largest-first into the set that is emptiest relative to what it can carry (a set of 2 carries less than a set of 3), then onto that set's least-loaded section, never putting two loads of the same duty/standby group on one section (when the group is no bigger than the number of sections; otherwise a note is shown).
If that first placement leaves a section or set over its limit, the placement is repaired: loads are moved or swapped between sections while that reduces the excess, and if that gets stuck the best placement is perturbed and improved again (seeded, so results are repeatable; limited work budget). This is a heuristic: "cannot be balanced" means no valid placement was found, not a proof that none exists.
Not modelled: electrical sharing between sections through the couplers (equal sharing within a set is assumed).

## PMCC / IMCC feeders
Every set of transformers (every PMCC / IMCC board) gets a fixed block of outgoing feeders, **listed**: 6 × 63 A, 6 × 32 A, 6 × 16 A and 10 × 6 A (28 installed per board) of which **4, 4, 4 and 6 outputs are counted in the load calculation** (480 A, about 360 kVA at 433 V per board); the other 10 are listed in the layout with zero load and "(not counted)" in the description. They are converted like any other amps-only feeder (√3 × 420 V × A for the 433 V level, power factor 0.7). Half of each kind sits on each side of the middle transformer of a set of 3 (the two end sections; the middle section gets none), or on each side of the bus coupler of a set of 2. They are fixed: the balancing never moves them. Each combination tried has its own total (load list + feeders of all its sets), shown in the Summary table and on the combination sheet; the "Total kVA of the load list" in the Summary excludes them. They are listed in the bus layout without a row number. A checkbox on the page (on by default) switches them on or off; in the code `DEFAULT_OPTIONS.mccFeeders` is empty and `STANDARD_MCC_FEEDERS` holds the block above. No spare feeders are added for these ratings.

## Excel report
The load list is added once, at the top of the sidebar, and is shared by all calculations. The report is created with **Download Excel report** under the result (file name `<input name>_transformer_sizing_<voltage>V.xlsx`). A switch at the top of the page, **Include kW, kvar and power factor** (on by default), controls those columns. The button re-creates it for the current voltage / rating / number of transformers.

Sheets: **Summary** (assumptions, totals, recommendation, every number of transformers tried), **Original load list** (the uploaded sheet exactly as read), **Loads** (classified loads with kW / kvar / kVA), one sheet per transformer combination shown ("2 Transformers", "3 Transformers", and the final added number, e.g. "6 Transformers": ratings checked, distribution and one-out table per set, and the bus layout) and **Issues** (excluded / error / warning rows, only when there are any).

**Bus layout** (on each combination sheet, one table per set): the board listed in order along the bus, with the **incomers in cyan** (`#80DEEA`) and the **bus couplers in light yellow** (`#FFF9C4`) and the outgoing loads between them. Each section's loads sit between its incomer and a coupler: set of 2 = Incomer T1 · loads · Bus coupler BC1 · loads · Incomer T2; set of 3 = Incomer T1 · loads · BC1 · Incomer T2 · loads · BC2 · loads · Incomer T3. An incomer row shows the section's total load and loading %. Couplers are named BC1, BC2 … (S1-BC1 … with several sets). Read top to bottom = left to right along the board.

**Spare feeders** (Excel only; not on the page): for each type (DOL / VFD / Feeder) and rating found in the load list, **one** spare outgoing feeder is added to the output and only its **SP cell** (W/S column) has the gold background (`#FFD966`, a stronger shade than the pastels so it stands out from all ten others). A rating is the kW for motors and the amps for loads given in amps, so "DOL 180 A" and "DOL 185 kW" are different ratings; a feeder given with a kW rating counts as a DOL motor. Spares appear (a) on the **Loads** sheet, in a block below the Total row, and (b) in the **bus layout** of every combination sheet, right after the first load of the same type and rating (one spare in total, not one per section). A spare row shows **SP** in the W/S column (alongside W and S), its description ends with **(Spare)**, and it shows only its type and rating: no kW, kvar or kVA. **Spares are not part of the calculation**: they are never in the loads, so no total, section, set, loading or number of transformers includes them (tests check this). The Summary shows how many were added, and the legend above the bus layout has a "Spare feeder" entry. The sample file gives 13 spares.

Formatting:
- **Colours:** table headers light orange (`#FFE0B2`); cells within their limit light green (`#C6EFCE`); cells over a limit light red (`#FFC7CE`, added so the two states are never confused); incomers cyan, bus couplers light yellow. Load kinds have their own pastels, chosen to be clearly different from those shades and from each other: **W** lavender `#DEB1FB`, **S** pink `#FBB1DE`, **DOL** periwinkle blue `#BACEF3`, **VFD** lime `#CBFBB1`, **Feeder** sand `#D8D4CA` (applied to the W/S and "Treated as" cells, to the same values in the original sheet, and listed in the legend above the bus layout). Limits are inclusive (exactly 90 % is within).
- **Numbers:** only percentages have a decimal place (`0.0`); every other number is shown as a whole number (`#,##0`; the value underneath is not rounded). The one exception is **power factor**, which stays at two decimals (`0.00`) because it is a ratio below 1 and would otherwise show as 1 or 0. Numbers in the "Original load list" sheet are shown exactly as uploaded.
- **Alignment:** all numbers and percentages, and the W/S columns, are centred; text is left-aligned.
- Thin borders, column widths from the content, frozen header rows on the list sheets, landscape printing.
- **The page uses the same colours and number formats as the report** (orange headers, thin borders, green / light red against the limits, the five pastels for W, S, DOL, VFD and Feeder, centred numbers, whole numbers except percentages, power factor with two decimals). The shared rules live in `src/core/format.ts` (number display) and `REPORT_COLORS` / `kindColor` in `src/core/report.ts`; messages generated by the calculation follow them too (e.g. "carries 4,524 kVA, above the 80.0 % normal limit of 1,600 kVA"). The report's incomer / bus-coupler colours apply only to the Excel bus layout; the page has no such table.
- **No author metadata:** the generated file contains no document properties at all (no creator / last-modified-by). The sample file served for download has them removed too.

## Interpretations to confirm
- PMCC / IMCC feeders: one board per set of transformers and three-phase feeders (√3 × V × A) are confirmed. Odd counts, if the block is ever changed, would put the extra feeder on the first side.
- "With one transformer out, 90 %" is applied as a 90 % limit in the one-out case (confirmed).
- Duty/standby pairs are identified by the optional group column (confirmed).
- "Number of transformers can be a combination of 2 or 3 only, one bus coupler between the incomers" is read as independent sets of 2 or 3 with couplers only inside a set.
- "Start with 3 transformer combinations, in the end balance can be two trafo combinations once options for 3 trafo get exhausted" is read as: start with a set of 3, add sets of 3 while the loads cannot be balanced, and use a set of 2 only at the end for the remainder (3, 5 = 3+2, 6, 8 = 3+3+2, 9, …). The automatic start is therefore 3 (not 2), even for a small load that a set of 2 could carry; "start with 2" is available. (Confirmed by the user.)
- "Increase the number of transformers if the load balancing is not happening" is read as: fit is decided on the real placement of whole loads, not only on the total.
