# Transformer sizing – rules as implemented

Source: the user's load-distribution brief (Rev 0) plus follow-up answers. Code: `src/core/transformer-sizing.ts`.

## Inputs
Excel `.xlsx`, first sheet, header row required (column order does not matter):
Description · Type (DOL / VFD / Feeder) · Amps rating · kW rating · Voltage level (433 or 690 V).
Optional: **Working/Standby (W/S)** column (any header containing "working", "standby" or "duty", or "W/S"; values W / Working / Duty or S / Standby), shown as W or S in the results and report. Standby loads count at 100 % like working loads.
Optional: **Duty/standby group** (any header containing "group" or "pair"): loads with the same group name are placed on different transformers' bus sections.
Sample: `public/sample_load_list.xlsx` (regenerate with `scripts/make_sample_xlsx.py`).

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
Transformers: 6600 / 433 V or 6600 / 690 V (chosen in the UI). The rating is chosen with radio buttons (2000, 2500 or 4000 kVA; default 2000). If the chosen rating cannot carry the load, the page says which standard rating would. All transformers in a group are identical.
- **2 transformers:** 2 incomers + 1 bus coupler. **3 transformers:** 3 incomers + 2 bus couplers (bus sections in a line, T1 – T2 – T3). The UI can show both (automatic) or force 2 or 3.
- Load shared equally in normal running: total ÷ (n × rating) ≤ **80 %**.
- One transformer out, the remaining n − 1 sharing the whole load equally through the closed bus couplers: total ÷ ((n − 1) × rating) ≤ **90 %**. Each outage is listed per transformer.
- **Adding transformers:** if none of the selected numbers (2, 3, or both in automatic mode) fits, more are added one at a time (4, 5, 6 …, up to **20**) until a combination fits. More than 3 transformers are built only from **independent sets of 2 or 3**: a set of 2 has 2 incomers and 1 bus coupler, a set of 3 has 3 incomers and 2 bus couplers (one coupler between neighbouring incomers) and there is **no coupler between sets**. The split with the fewest sets is used (4 = 2+2, 5 = 3+2, 6 = 3+3, 7 = 3+2+2, 8 = 3+3+2, 9 = 3+3+3, …); couplers = transformers − sets.
- **Capacity:** with one transformer out, only the other one or two in the same set carry that set's load. For rating R a set of 2 carries min(0.8·2·R, 0.9·1·R) = 0.9R and a set of 3 carries min(0.8·3·R, 0.9·2·R) = 1.8R; the total is the sum over the sets (2000 kVA: 1800 / 3600 / 3600 / 5400 / 7200 / 7200 / 9000 / 10800 for 2 … 9 transformers, 23400 for 20). The one-out loading shown is total ÷ ((transformers − sets) × R). Note that 4 (2+2), 7 (3+2+2), 10 … carry no more than one transformer fewer, so they never become the recommendation.
- Loads are placed whole: largest first into the set that is emptiest relative to what it can carry, then onto that set's least-loaded section. If an indivisible load makes a set exceed its limit, a note says so. Sections are named T1, T2 … for one set and S1-T1, S1-T2, S2-T1 … for several.
- Only the selected numbers and the final (fitting) number get a full card / Excel sheet; every number tried is listed in a "numbers of transformers tried" table.
- For each number of transformers the smallest rating that passes both limits is chosen (if none passes, the largest rating is shown, marked "does not fit").
- With a single rating chosen, the recommendation is the first of 2 × rating, 3 × rating that passes.
- Total kVA: arithmetic sum of each load's kVA (confirmed by the user). It is slightly above √((Σ kW)² + (Σ kvar)²).
- The 80 % / 90 % limits are fixed (not editable). Power factors and efficiency are editable in the UI.

## Load distribution (draft)
Whole loads are placed on the transformer bus sections largest-first onto the least-loaded section, never putting two loads of the same duty/standby group on one section (when the group is no bigger than the number of sections; otherwise a note is shown). If the total fits but the balanced distribution still pushes a section over the normal limit, a note says so.
Not yet modelled: electrical sharing between sections through the couplers (equal sharing within a set is assumed).

## Excel report
The report is **created and downloaded automatically as soon as the load list is dropped** (file name `<input name>_transformer_sizing_<voltage>V.xlsx`). A switch next to the drop area, **Include kW, kvar and power factor** (on by default), controls those columns. "Download Excel report again" re-creates it for the current voltage / rating / number of transformers.

Sheets: **Summary** (assumptions, totals, recommendation, every number of transformers tried), **Original load list** (the uploaded sheet exactly as read), **Loads** (classified loads with kW / kvar / kVA), one sheet per transformer combination shown ("2 Transformers", "3 Transformers", and the final added number, e.g. "6 Transformers": ratings checked, distribution and one-out table per set, loads by transformer) and **Issues** (excluded / error / warning rows, only when there are any).

Formatting: table headers light orange (`#FFE0B2`); cells within their limit light green (`#C6EFCE`); cells over a limit light red (`#FFC7CE`, added so the two states are never confused); thin borders, column widths from the content, frozen header rows on the list sheets, landscape printing. Limits are inclusive (exactly 90 % is within).

## Interpretations to confirm
- "With one transformer out, 90 %" is applied as a 90 % limit in the one-out case (confirmed).
- Duty/standby pairs are identified by the optional group column (confirmed).
- "Number of transformers can be a combination of 2 or 3 only, one bus coupler between the incomers" is read as independent sets of 2 or 3 with couplers only inside a set. The choice of split for a given number (fewest sets) is mine – to be confirmed.
