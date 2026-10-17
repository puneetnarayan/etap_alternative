# Open engineering items

Not resolved by guessing. Accepted calculation conventions are unchanged unless explicitly authorised. Remind at each design milestone.

| # | Item | State |
|---|------|-------|
| 1 | Coordination of IMCC incomer ratings and PMCC outgoing feeder ratings. Both PMCC feeders carry the whole IMCC working current; the incomer ACB size is the user's choice (630 / 800 / 1000 / 1250 A). | Open |
| 2 | Applicable breaker-sizing standards and design margins (1.25 multiplier, ladder, 630 A / ACB threshold). | Open |
| 3 | Panel interpretation: one bus, all loads on it; Incomer 1 normal, Incomer 2 and the bus coupler (side by side at the other end) take over the whole IMCC when Incomer 1 is unhealthy; all three have the same rating. Loads are filled into IMCCs in list order. | Confirmed |
| 4 | Validation of the incomer sizing algorithm: IMCC design current = sum of working currents x 1.25; limit = ACB size x maximum loading (90 %, adjustable) on that design current (confirmed). The 1.25 multiplier and the electrical validation remain open. | Open |
| 5 | Duty/standby group validation (only a warning for a group with standby but no working load). | Open |
| 6 | Final Excel output format (one workbook: Summary, one sheet per IMCC, Notes). The PMCC workbook is the existing one; with the switch ON the list it is built from has the incomer rows added. | Open |
| 7 | The IMCC sample data: 16 of 51 rows have no W/S/Spare status (flagged, not counted); one row has "-" for amps and kW; the amps column is used as given (calculated at which voltage and power factor?). The ECR-9 / ECR-10 sheets were only an example. | Open |
| 8 | Meaning and applicability of the LCE calculation (working kVA x 0.8). Reported, not used. | Open |
| 9 | Electrical calculation and regression testing against an independent calculation. | Open |
| 10 | Standby (S) rows are not counted in the transformer sizing (PMCC or IMCC): confirmed, they would add transformers that are not needed. They are listed as excluded; the MCC and BOM tabs still select a breaker for them. PMCC parsing of `Spare` / `SP` is unchanged (not recognised there). | Confirmed |
| 11 | Carried over: BOM Class column logic; soft-starter rule (225 kW or 315 kW); which "yellow cells" need technical review. | Open |
| 12 | Soft starter (SS) is calculated like a DOL motor (pf 0.7, efficiency 0.93). | Confirmed |
| 13 | The incomer limit applies to design current (working x 1.25), at most the selected ACB x 90 % (adjustable). | Confirmed (90 % on design current) |
