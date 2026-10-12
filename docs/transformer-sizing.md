# Transformer sizing – rules as implemented

Source: the user's load-distribution brief (Rev 0) plus follow-up answers. Code: `src/core/transformer-sizing.ts`.

## Inputs
Excel `.xlsx`, first sheet, header row required (column order does not matter):
Description · Type (DOL / VFD / Feeder) · Amps rating · kW rating · Voltage level (433 or 690 V).
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
Transformers: 6600 / 433 V or 6600 / 690 V (chosen in the UI). Ratings are typed in the UI (default 2000, 2500, 4000 kVA; often just one). All transformers in a group are identical.
- **2 transformers:** 2 incomers + 1 bus coupler. **3 transformers:** 3 incomers + 2 bus couplers (bus sections in a line, T1 – T2 – T3). The UI can show both (automatic) or force 2 or 3.
- Load shared equally in normal running: total ÷ (n × rating) ≤ **80 %**.
- One transformer out, the remaining n − 1 sharing the whole load equally through the closed bus couplers: total ÷ ((n − 1) × rating) ≤ **90 %**. Each outage is listed per transformer.
- For each number of transformers the smallest rating that passes both limits is chosen (if none passes, the largest rating is shown, marked "does not fit").
- The overall recommendation tries candidates in order of lowest total installed kVA (2×2000, 2×2500, 3×2000, 3×2500, 2×4000, 3×4000) and takes the first that passes; "add a 3rd transformer before increasing rating" is selectable.
- Total kVA: sum of each load's kVA (conservative, default) or √((Σ kW)² + (Σ kvar)²) — selectable; both are shown. **Which one to use is still to be confirmed.**
- Limits, power factors and efficiency are editable in the UI.

## Load distribution (draft)
Whole loads are placed on the transformer bus sections largest-first onto the least-loaded section, never putting two loads of the same duty/standby group on one section (when the group is no bigger than the number of sections; otherwise a note is shown). If the total fits but the balanced distribution still pushes a section over the normal limit, a note says so.
Not yet modelled: electrical sharing between sections through the couplers (equal sharing is assumed), and per-section checks of the outage cases.

## Excel report
"Download Excel report…" opens a panel with a switch **Include kW, kvar and power factor** (on by default). The workbook has: Summary (assumptions, totals, recommendation, per-combination summary), Loads, one sheet per transformer combination ("2 Transformers", "3 Transformers": ratings checked, distribution, one-out table, loads by transformer) and Issues (excluded / error / warning rows).

## Interpretations to confirm
- "With one transformer out, 90 %" is applied as a 90 % limit in the one-out case (confirmed).
- Total kVA basis (arithmetic vs vector) — see above.
- Duty/standby pairs are identified by the optional group column.
