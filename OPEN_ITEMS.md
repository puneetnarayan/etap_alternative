# Open engineering items

Not resolved by guessing. Accepted calculation conventions are unchanged unless explicitly authorised. Remind at each design milestone.

| # | Item | State |
|---|------|-------|
| 1 | Coordination of IMCC incomer ratings and PMCC outgoing feeder ratings. The PMCC feeder is sized from its own section current; Incomer 2 and the coupler follow Incomer 1 (rule). They can differ; the app warns. | Open |
| 2 | Applicable breaker-sizing standards and design margins (1.25 multiplier, ladder, 630 A / ACB threshold). | Open |
| 3 | Bus-section mapping and load allocation. The script fills sections in list order up to 1000 A (agreed rule); the real allocation (and what a source carries if the other is lost) is not defined. | Open |
| 4 | Validation of the incomer sizing algorithm: incomer design current = sum of the section's working currents (arithmetic) x 1.25. There was no incomer-from-loads algorithm in the app before. | Open |
| 5 | Duty/standby group validation (only a warning for a group with standby but no working load). | Open |
| 6 | Final Excel output format (one workbook: Summary, one sheet per IMCC in the ECR layout, Notes). The PMCC workbook is the existing one; with the switch ON the list it is built from has the incomer rows added. | Open |
| 7 | Verification of the ECR-9 / ECR-10 reference data: many rows have no W/S/Spare status (ECR-10 rows 7 to 17 and 40 to 44; ECR-9 I/C, B/C rows and the 1-ph feeders). The reference uses 415 V and 1.73; the app uses 433 V and √3. | Open |
| 8 | Meaning and applicability of the LCE calculation (working kVA x 0.8). Reported, not used. | Open |
| 9 | Electrical calculation and regression testing against an independent calculation. | Open |
| 10 | With the switch ON the standby incomer (Incomer 2) feeder is counted at 100 % in the transformer sizing, as every standby load is. PMCC parsing of `Spare` / `SP` is unchanged (not recognised there). | Open |
| 11 | Carried over: BOM Class column logic; soft-starter rule (225 kW or 315 kW); which "yellow cells" need technical review. | Open |
| 12 | Soft starter (SS) is calculated like a DOL motor (pf 0.7, efficiency 0.93). Confirm, and whether a soft starter needs its own pf, efficiency or starting current. | Open |
| 13 | 1000 A incomer limit is applied to design current (working current x 1.25), so the incomer breaker never exceeds 1000 A. Confirm. | Open |
