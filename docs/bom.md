# BOM tab

Sidebar entry **BOM** (route `/bom`). It builds one unit BOM for each unique item (type + rating) of the load list.

## Flow
1. The load list is the one last used in the browser tab (Transformer sizing, MCC Load List or this page), or dropped on this page. Dropping a file builds and downloads the workbook at once.
2. Each load goes through the preliminary breaker selection of the MCC Load List tab (so its editable assumptions, including "630 A as MCCB", apply).
3. Loads with the same **type + breaker rating (+ motor kW for motors)** are one item, counted.
4. Each item is matched to a unit-BOM sheet of the template `public/bom/unit_bom_template.xlsx` and cloned into the output with its layout, fonts, column widths and colours unchanged; only rating text is set (breaker rating, CT primary with a 1 A secondary, ELR rating, sheet title).
5. The output workbook has an **Index** (SL, type, breaker / rating, **Class**, motor kW, quantity, load rows, sheet, action) and one sheet per item. The output has no author metadata.

## Matching rules (all preliminary)
| Item | Sheet |
|---|---|
| Incomer (description contains "incomer") | Incomer 1000A / 1250A IMCC, 1600A, 2000A, 2500A, 3200A, 4000A |
| Feeder on an ACB | OG 630A ACB Feeder (new), OG 1000A ACB, OG 1250A ACB; above 1250 A: no sheet (red action) |
| Feeder on an MCCB of 250 A and above | OG 250 to 630A drawout Fdr MCCB |
| Feeder on an MCCB of 16 to 200 A | OG <250A power FDR - MCCB |
| Plain MPCB / MCB feeder (6, 10 A) | none: **not required** |
| DOL motor, MCCB | DOL < 75KW (75 kW and above: red action to confirm) |
| DOL motor above 200 kW (soft starter, BOM index note) | on an ACB: both OG 315KW and OG 225KW soft-starter sheets are kept with a **red action** (rule to be reviewed: 225 kW or 315 kW); on an MCCB: OG 75KW Soft starter |
| VFD motor, MCCB | 55kW VFD |
| RDOL (Type text) | RDOL < 5KW |
| No breaker selected (engineering review) | no sheet, red action |

The hidden 630 A, 800 A ACB and 3200 A sheets are in the template, kept for later use.

## Template
Built from the supplied SGB PMCC / IMCC BOM (converted from `.xlsb`), with the CT secondary changed to 1 A everywhere (4000 A incomer and bus coupler CTs 4000/1 A), the index and tools sheets removed, and new sheets cloned from existing ones: OG 630A ACB Feeder, Incomer / Buscoupler 1600A, 2000A, 2500A (vendor comments cleared; busbar sizes left "to be confirmed by vendor"). The template still contains the client and vendor text of the source BOM, kept as it is for the time being. Items found in review that were not decided (typos, ELR ratings, sheet labels) are not corrected in the template.

## Open points (reminders)
- **Class** column (breaker device class): empty until the logic is given.
- Soft-starter ACB rule (225 or 315 kW).
- Bus couplers (not in a load list) are not generated; they will follow the transformer sets.
- Incomer rating from the transformer (kVA, voltage) is not linked yet.
- DOL motors of 75 kW and above have no sheet of their own ("< 75KW" template reused).
