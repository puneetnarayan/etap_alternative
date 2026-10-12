# Transformer sizing – rules as implemented

Source: the user's load-distribution brief (Rev 0) plus follow-up answers. Code: `src/core/transformer-sizing.ts`.

## Inputs
Excel `.xlsx`, first sheet, header row required (column order does not matter):
Description · Type (DOL / VFD / Feeder) · Amps rating · kW rating · Voltage level (433 or 690 V).
Optional: **Working/Standby (W/S)** column (any header containing "working", "standby" or "duty", or "W/S"; values W / Working / Duty or S / Standby), shown as W or S in the results and report. Standby loads count at 100 % like working loads.
Optional: **Duty/standby group** (any header containing "group" or "pair"): loads with the same group name are placed on different transformers' bus sections.
Sample: `public/sample_load_list.xlsx` is the file users download from the page (supplied by the project owner; replace the file to change it). The tests for it are in `src/core/__tests__/public-sample.test.ts`. The earlier hand-built small plant is kept only as a test fixture, `src/core/__tests__/fixtures/small_plant_load_list.xlsx`.

## Load classification and kVA
| Given | Treated as | kVA |
|---|---|---|
| kW | Motor (DOL pf 0.7, VFD pf 0.9; other/blank type → DOL, with a warning) | input kW = kW ÷ 0.95; kVA = input kW ÷ pf |
| Amps only, Type DOL/VFD | Motor | √3 × V × A ÷ 1000 |
| Amps only, anything else | Outgoing feeder (pf 0.7) | √3 × V × A ÷ 1000 |
| Neither | Error – row not counted | – |

kvar = input kW × tan(arccos pf), so kVA² = kW² + kvar². For amps-only loads kW = kVA × pf.
If kW is given, any amps value on the same row is ignored. All loads count at 100 % (standby included).
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

## Load distribution
Whole loads are never split. They are first placed largest-first into the set that is emptiest relative to what it can carry (a set of 2 carries less than a set of 3), then onto that set's least-loaded section, never putting two loads of the same duty/standby group on one section (when the group is no bigger than the number of sections; otherwise a note is shown).
If that first placement leaves a section or set over its limit, the placement is repaired: loads are moved or swapped between sections while that reduces the excess, and if that gets stuck the best placement is perturbed and improved again (seeded, so results are repeatable; limited work budget). This is a heuristic: "cannot be balanced" means no valid placement was found, not a proof that none exists.
Not modelled: electrical sharing between sections through the couplers (equal sharing within a set is assumed).

## Excel report
The report is **created and downloaded automatically as soon as the load list is dropped** (file name `<input name>_transformer_sizing_<voltage>V.xlsx`). A switch next to the drop area, **Include kW, kvar and power factor** (on by default), controls those columns. "Download Excel report again" re-creates it for the current voltage / rating / number of transformers.

Sheets: **Summary** (assumptions, totals, recommendation, every number of transformers tried), **Original load list** (the uploaded sheet exactly as read), **Loads** (classified loads with kW / kvar / kVA), one sheet per transformer combination shown ("2 Transformers", "3 Transformers", and the final added number, e.g. "6 Transformers": ratings checked, distribution and one-out table per set, and the bus layout) and **Issues** (excluded / error / warning rows, only when there are any).

**Bus layout** (on each combination sheet, one table per set): the board listed in order along the bus, with the **incomers in cyan** (`#80DEEA`) and the **bus couplers in light yellow** (`#FFF9C4`) and the outgoing loads between them. Each section's loads sit between its incomer and a coupler: set of 2 = Incomer T1 · loads · Bus coupler BC1 · loads · Incomer T2; set of 3 = Incomer T1 · loads · BC1 · Incomer T2 · loads · BC2 · loads · Incomer T3. An incomer row shows the section's total load and loading %. Couplers are named BC1, BC2 … (S1-BC1 … with several sets). Read top to bottom = left to right along the board.

Formatting: table headers light orange (`#FFE0B2`); cells within their limit light green (`#C6EFCE`); cells over a limit light red (`#FFC7CE`, added so the two states are never confused); thin borders, column widths from the content, frozen header rows on the list sheets, landscape printing. Limits are inclusive (exactly 90 % is within).

## Interpretations to confirm
- "With one transformer out, 90 %" is applied as a 90 % limit in the one-out case (confirmed).
- Duty/standby pairs are identified by the optional group column (confirmed).
- "Number of transformers can be a combination of 2 or 3 only, one bus coupler between the incomers" is read as independent sets of 2 or 3 with couplers only inside a set.
- "Start with 3 transformer combinations, in the end balance can be two trafo combinations once options for 3 trafo get exhausted" is read as: start with a set of 3, add sets of 3 while the loads cannot be balanced, and use a set of 2 only at the end for the remainder (3, 5 = 3+2, 6, 8 = 3+3+2, 9, …). The automatic start is therefore 3 (not 2), even for a small load that a set of 2 could carry; "start with 2" is available. (Confirmed by the user.)
- "Increase the number of transformers if the load balancing is not happening" is read as: fit is decided on the real placement of whole loads, not only on the total.
