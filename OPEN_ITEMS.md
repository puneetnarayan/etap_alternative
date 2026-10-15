# Open engineering items

Not resolved by guessing. Accepted calculation conventions are unchanged unless explicitly authorised. Remind at each design milestone.

| # | Item | State |
|---|------|-------|
| 1 | Coordination of IMCC incomer ratings and PMCC outgoing feeder ratings. The PMCC feeder is sized from its own section current; Incomer 2 and the coupler follow Incomer 1 (rule). They can differ; the app warns. | Open |
| 2 | Applicable breaker-sizing standards and design margins (1.25 multiplier, ladder, 630 A / ACB threshold). | Open |
| 3 | Bus-section mapping and load allocation. The script fills sections in list order up to 1000 A (agreed rule); the real allocation (and what a source carries if the other is lost) is not defined. | Open |
| 4 | Validation of the incomer sizing algorithm: incomer design current = sum of the section's working currents (arithmetic) x 1.25. There was no incomer-from-loads algorithm in the app before. | Open |
| 5 | Duty/standby group validation (only a warning for a group with standby but no working load). | Open |
| 6 | Final Excel output format (one workbook: Summary, one sheet per IMCC, Notes). The PMCC workbook is the existing one; with the switch ON the list it is built from has the incomer rows added. | Open |
| 7 | The IMCC sample data: 16 of 51 rows have no W/S/Spare status (flagged, not counted); one row has "-" for amps and kW; the amps column is used as given (calculated at which voltage and power factor?). The ECR-9 / ECR-10 sheets were only an example. | Open |
| 8 | Meaning and applicability of the LCE calculation (working kVA x 0.8). Reported, not used. | Open |
| 9 | Electrical calculation and regression testing against an independent calculation. | Open |
| 10 | The standby incomer (Incomer 2) feeder is counted at 100 % in the transformer sizing, for sizing only, as every standby load is. PMCC parsing of `Spare` / `SP` is unchanged (not recognised there). | Confirmed (sizing only) |
| 11 | Carried over: BOM Class column logic; soft-starter rule (225 kW or 315 kW); which "yellow cells" need technical review. | Open |
| 12 | Soft starter (SS) is calculated like a DOL motor (pf 0.7, efficiency 0.93). | Confirmed |
| 13 | The 1000 A incomer limit applies to design current (working current x 1.25); 1000 A is the maximum ACB size for an IMCC incomer. | Confirmed |
