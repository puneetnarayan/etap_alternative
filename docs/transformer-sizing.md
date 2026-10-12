# Transformer sizing – rules as implemented

Source: the user's load-distribution brief (Rev 0) plus follow-up answers. Code: `src/core/transformer-sizing.ts`.

## Inputs
Excel `.xlsx`, first sheet, header row required (column order does not matter):
Description · Type (DOL / VFD / Feeder) · Amps rating · kW rating · Voltage level (433 or 690 V).
Sample: `public/sample_load_list.xlsx` (regenerate with `scripts/make_sample_xlsx.py`).

## Load classification and kVA
| Given | Treated as | kVA |
|---|---|---|
| kW | Motor (DOL pf 0.7, VFD pf 0.9; other/blank type → DOL, with a warning) | kW ÷ (pf × 0.95) |
| Amps only, Type DOL/VFD | Motor | √3 × V × A ÷ 1000 |
| Amps only, anything else | Outgoing feeder (pf 0.7) | √3 × V × A ÷ 1000 |
| Neither | Error – row not counted | – |

If kW is given, any amps value on the same row is ignored. All loads count at 100 % (standby included).
Rows whose voltage differs from the selected transformer are excluded (listed on screen); a blank voltage is assumed to match, with a warning.

## Transformer selection
Transformers: 6600 / 433 V or 6600 / 690 V (chosen in the UI). Ratings 2000, 2500, 4000 kVA. All transformers in a group are identical.
- **2 transformers:** 2 incomers + 1 bus coupler. **3 transformers:** 3 incomers + 2 bus couplers.
- Load shared equally in normal running: total ÷ (n × rating) ≤ **80 %**.
- One transformer out, load shared equally by the rest: total ÷ ((n − 1) × rating) ≤ **90 %**.
- Candidates are tried in order of lowest total installed kVA (2×2000, 2×2500, 3×2000, 3×2500, 2×4000, 3×4000); the first that passes both limits is recommended. An alternative order ("add a 3rd transformer before increasing rating") is selectable. Limits, power factors and efficiency are editable in the UI.

## Load distribution (draft – to be refined with the user)
Whole loads are placed on the transformer bus sections largest-first onto the least-loaded section. Not yet modelled: keeping duty/standby pairs on different sections, bus-section topology for the 3-transformer case, and per-section checks with one transformer out.

## Interpretations to confirm
- "With one transformer out, then also 90 %" was read as a 90 % limit in the one-out case.
- "2600 kVA" in the brief was corrected by the user to 2500 kVA.
