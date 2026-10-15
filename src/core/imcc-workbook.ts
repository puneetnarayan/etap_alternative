/**
 * Excel output for the IMCC design (docs/imcc.md). ExcelJS, loaded only when a workbook is generated.
 *
 * One workbook: a Summary sheet, one sheet per IMCC (layout of the reference workbook: S. NO., description, kW, current,
 * W/S/SP, method, group, current, kVA, working kVA, working current, LCE kVA, design current, breaker, notes; formulas,
 * section totals, incomer and bus-coupler ratings) and a Notes sheet with the assumptions the formulas read.
 * Same look as the other workbooks of the application: Calibri 11, light-orange headers, blue input / green calculated.
 */
import ExcelJS from "exceljs";
import { DEFAULT_MCC_ASSUMPTIONS, breakerLabel, type MccAssumptions } from "./mcc-breaker";
import { stripAuthor } from "./mcc-workbook";
import { LCE_FACTOR, imccFeederRows, type ImccDesign, type ImccDesignSet, type LoadResult, type SectionResult } from "./imcc";

const THIN = { style: "thin" as const, color: { argb: "FFBFBFBF" } };
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const fill = (argb: string) => ({ type: "pattern" as const, pattern: "solid" as const, fgColor: { argb } });
const HEADER = "FFFCE4D6"; // light orange, as the other reports
const INPUT = "FFDDEBF7";
const CALC = "FFE2EFDA";
const SPARE = "FFFFD966";
const BANNER = "FFFFF2CC";
const F = (extra: Partial<ExcelJS.Font> = {}): Partial<ExcelJS.Font> => ({ name: "Calibri", size: 11, ...extra });

// Notes sheet cells the formulas read.
const N = { pfGen: "Notes!$B$4", pfDol: "Notes!$B$5", pfVfd: "Notes!$B$6", eff: "Notes!$B$7", mult: "Notes!$B$8", volt: "Notes!$B$9", lce: "Notes!$B$10" };

export interface ImccWorkbookOptions {
  sourceName?: string;
  /** "Automatic PMCC feeders" switch state: names the PMCC outgoing feeder of each incomer, or says it is not linked. */
  linked: boolean;
}

export function buildImccWorkbook(set: ImccDesignSet, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS, o: ImccWorkbookOptions = { linked: false }): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.creator = ""; wb.lastModifiedBy = "";
  wb.calcProperties.fullCalcOnLoad = true;
  buildSummary(wb.addWorksheet("Summary", { views: [{ showGridLines: false }] }), set, o);
  for (const i of set.imccs) buildImccSheet(wb.addWorksheet(i.tag, { views: [{ state: "frozen", ySplit: 12, showGridLines: false }] }), i, set, o);
  buildNotes(wb.addWorksheet("Notes", { views: [{ showGridLines: false }] }), set, a);
  return wb;
}

function buildSummary(ws: ExcelJS.Worksheet, set: ImccDesignSet, o: ImccWorkbookOptions) {
  [6, 12, 14, 14, 14, 18, 14, 14, 14, 18, 60].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  ws.getCell("A1").value = "IMCC design – summary"; ws.getCell("A1").font = F({ bold: true, size: 16 });
  ws.getCell("A2").value = "Preliminary engineering aid, not a substitute for approved design. Only working (W / w) loads are used; the number of IMCCs follows from the loads."; ws.getCell("A2").font = F({ italic: true });
  const info: [string, string][] = [
    ["Source", o.sourceName ?? "IMCC load list"],
    ["Incomer limit", `${set.maxIncomerA} A design current per incomer (working current x design multiplier)`],
    ["IMCCs", String(set.imccs.length)],
    ["Automatic PMCC feeders", o.linked ? "ON: each incomer is an outgoing feeder of the PMCC list (see below)" : "OFF: the IMCC design is independent; nothing is added to the PMCC"],
    ["Operating philosophy", "Exactly one incomer ON at a time; the other incomer and the bus coupler stay OFF."],
  ];
  info.forEach(([k, v], i) => {
    ws.getCell(4 + i, 1).value = k; ws.getCell(4 + i, 1).font = F({ bold: true }); ws.mergeCells(4 + i, 1, 4 + i, 3);
    ws.getCell(4 + i, 4).value = v; ws.getCell(4 + i, 4).font = F(); ws.mergeCells(4 + i, 4, 4 + i, 11);
  });
  const H = 10;
  ["", "IMCC", "Loads (W / S / SP)", "Incomer 1 working A", "Incomer 1 design A", "Incomer 1 rating", "Incomer 2 working A", "Incomer 2 design A", "Incomer 2 rating (= Incomer 1)", "Bus coupler (= Incomer 1)", "Checks"].forEach((t, i) => {
    const c = ws.getCell(H, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(H).height = 34;
  set.imccs.forEach((i, n) => {
    const [s1, s2] = i.sections;
    const v: ExcelJS.CellValue[] = [n + 1, i.tag, `${s1.working + s2.working} / ${s1.standby + s2.standby} / ${s1.spare + s2.spare}`, s1.workingA, s1.designA,
      s1.workingA > 0 ? breakerLabel(i.incomer1) : "", s2.workingA, s2.designA, s2.workingA > 0 ? breakerLabel(i.incomer2) : "", breakerLabel(i.coupler),
      i.issues.map((x) => x.message).join(" ")];
    v.forEach((x, c) => { const cell = ws.getCell(H + 1 + n, c + 1); cell.value = x; cell.border = BORDER; cell.font = F(); cell.alignment = { vertical: "top", wrapText: c === 10 }; if (typeof x === "number" && c > 2) cell.numFmt = "#,##0"; });
  });
  let r = H + set.imccs.length + 3;
  if (o.linked) {
    ws.getCell(r, 1).value = "PMCC outgoing feeders generated for the incomers (switch ON)"; ws.getCell(r, 1).font = F({ bold: true, size: 13 }); r++;
    ["", "Outgoing feeder", "Status", "Current (A)", "kVA", "Group"].forEach((t, i) => { const c = ws.getCell(r, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; });
    r++;
    for (const f of imccFeederRows(set)) {
      [r, f.description, f.status, f.amps, f.kva, f.group].forEach((x, i) => { if (i === 0) return; const c = ws.getCell(r, i + 1); c.value = x as ExcelJS.CellValue; c.border = BORDER; c.font = F(); if (typeof x === "number") c.numFmt = "#,##0"; });
      r++;
    }
    ws.getCell(r + 1, 1).value = "The two incomers are alternative sources: Incomer 1 is W, Incomer 2 is S (standby); they share a group so they land on different transformer sections.";
    ws.getCell(r + 1, 1).font = F({ italic: true });
    r += 3;
  }
  for (const x of set.issues) { ws.getCell(r, 2).value = `${x.severity === "error" ? "Error" : "Check"}: ${x.message}`; ws.getCell(r, 2).font = F({ color: { argb: x.severity === "error" ? "FFC00000" : "FF7F6000" } }); r++; }
}

function buildImccSheet(ws: ExcelJS.Worksheet, imcc: ImccDesign, set: ImccDesignSet, o: ImccWorkbookOptions) {
  [7, 46, 9, 12, 9, 11, 30, 11, 10, 11, 12, 10, 12, 20, 60, 12].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  ws.getCell("A1").value = `${imcc.tag} – IMCC load list`; ws.getCell("A1").font = F({ bold: true, size: 16 });
  ws.getCell("A2").value = "Preliminary engineering aid, not a substitute for approved design. Blue = input, green = calculated. Assumptions: Notes sheet."; ws.getCell("A2").font = F({ italic: true });
  const feeder = (n: 1 | 2) => (o.linked ? (imcc.sections[n - 1].workingA > 0 ? `PMCC outgoing feeder "${imcc.tag} Incomer ${n}"` : "no PMCC feeder (no working load)") : "PMCC feeder not generated (Automatic PMCC feeders switch OFF)");
  const info: [string, string][] = [
    ["IMCC tag", imcc.tag],
    ["Design basis", `Working (W / w) loads only; incomer limit ${set.maxIncomerA} A design current. Voltage ${set.voltageV} V.`],
    ["Incomer 1 supplied from", feeder(1)],
    ["Incomer 2 supplied from", feeder(2)],
    ["Operating philosophy", "Exactly one incomer ON at a time (normally Incomer 1); the other incomer and the bus coupler stay OFF."],
  ];
  info.forEach(([k, v], i) => {
    const r = 4 + i;
    ws.getCell(r, 1).value = k; ws.getCell(r, 1).font = F({ bold: true }); ws.mergeCells(r, 1, r, 2);
    ws.getCell(r, 3).value = v; ws.getCell(r, 3).font = F(); ws.mergeCells(r, 3, r, 15);
  });

  const HEAD = 12;
  ["S. NO.", "Load description", "kW", "Current entered (A)", "W / S / SP", "Method of starting", "Group", "Current (A)", "kVA",
    "Working kVA", "Working current (A)", `LCE kVA (x ${LCE_FACTOR})`, "Design current (A)", "Calculated breaker", "Notes / feeder description", "Status as entered"].forEach((t, i) => {
    const c = ws.getCell(HEAD, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(HEAD).height = 34;

  let row = HEAD + 1;
  const totals: number[] = [];
  for (const sec of imcc.sections) {
    ws.mergeCells(row, 1, row, 15);
    const b = ws.getCell(row, 1);
    b.value = `INCOMER ${sec.section} – supplied from ${feeder(sec.section)}`;
    b.font = F({ bold: true }); b.fill = fill(BANNER); b.border = BORDER;
    row++;
    const first = row;
    for (const x of sec.loads) { writeLoad(ws, row, x); row++; }
    const last = row - 1;
    ws.getCell(row, 2).value = `Total – Incomer ${sec.section} section`;
    for (const [col, ci] of [["J", 10], ["K", 11], ["L", 12]] as const) {
      const c = ws.getCell(row, ci); c.value = (last >= first ? { formula: `SUM(${col}${first}:${col}${last})` } : 0) as ExcelJS.CellValue; c.numFmt = "#,##0";
    }
    ws.getCell(row, 13).value = { formula: `K${row}*${N.mult}` } as ExcelJS.CellValue; ws.getCell(row, 13).numFmt = "#,##0";
    ws.getCell(row, 14).value = sec.workingA > 0 ? breakerLabel(sec.breaker) : "";
    ws.getCell(row, 15).value = "Design current = working current x design multiplier; breaker from the application's standard ratings.";
    for (let c = 1; c <= 16; c++) { const cell = ws.getCell(row, c); cell.font = F({ bold: true }); cell.border = BORDER; cell.fill = fill(CALC); }
    totals.push(row);
    row += 2;
  }

  ws.mergeCells(row, 1, row, 15);
  const sh = ws.getCell(row, 1);
  sh.value = "INCOMER, BUS COUPLER AND OVERALL (calculated design ratings, not nameplate or installed-device ratings)";
  sh.font = F({ bold: true }); sh.fill = fill(BANNER); sh.border = BORDER;
  row++;
  const rate = (s: { ratingA?: number } & Parameters<typeof breakerLabel>[0]) => (s.ratingA ? breakerLabel(s) : "Not calculated");
  const summary: [string, string, string][] = [
    ["Incomer 1", rate(imcc.incomer1), `Calculated from the Incomer 1 section's working load (limit ${set.maxIncomerA} A).`],
    ["Incomer 2", rate(imcc.incomer2), "Design rating equals Incomer 1 (rule)."],
    ["Bus coupler", rate(imcc.coupler), "Design rating equals Incomer 1 (rule). Stays OFF in normal operation."],
  ];
  for (const [k, v, note] of summary) {
    ws.getCell(row, 2).value = k; ws.getCell(row, 2).font = F({ bold: true });
    ws.getCell(row, 14).value = v; ws.getCell(row, 14).font = F({ bold: true });
    ws.getCell(row, 15).value = note; ws.getCell(row, 15).font = F();
    for (let c = 1; c <= 15; c++) ws.getCell(row, c).border = BORDER;
    row++;
  }
  ws.getCell(row, 2).value = "Both sections, summed";
  for (const [col, ci] of [["J", 10], ["K", 11], ["L", 12]] as const) {
    const c = ws.getCell(row, ci); c.value = { formula: totals.map((t) => `${col}${t}`).join("+") } as ExcelJS.CellValue; c.numFmt = "#,##0";
  }
  ws.getCell(row, 15).value = "Information only: under the operating philosophy the two sections are not supplied together, so this is not one source's demand.";
  for (let c = 1; c <= 15; c++) { ws.getCell(row, c).border = BORDER; ws.getCell(row, c).font = F({ bold: c !== 15 }); }
  row += 2;
  for (const i of imcc.issues) { ws.getCell(row, 2).value = `Check: ${i.message}`; ws.getCell(row, 2).font = F({ color: { argb: "FF7F6000" } }); row++; }
}

function writeLoad(ws: ExcelJS.Worksheet, r: number, x: LoadResult) {
  const l = x.load;
  const set = (c: number, v: ExcelJS.CellValue, fmt?: string, bg?: string) => {
    const cell = ws.getCell(r, c);
    cell.value = v; cell.border = BORDER; cell.font = F();
    cell.alignment = { vertical: "middle", wrapText: c === 2 || c === 15 };
    if (fmt) cell.numFmt = fmt;
    if (bg) cell.fill = fill(bg);
  };
  const num0 = "#,##0";
  set(1, l.sNo ?? "", undefined, INPUT);
  set(2, l.description, undefined, INPUT);
  set(3, l.kw ?? null, "#,##0.0", INPUT);
  set(4, l.ratedA ?? null, num0, INPUT);
  set(5, l.status ?? "", undefined, l.status === "SP" ? SPARE : INPUT);
  ws.getCell(r, 5).alignment = { horizontal: "center", vertical: "middle" };
  set(6, l.method ?? "", undefined, INPUT);
  set(7, l.group ?? "", undefined, INPUT);
  // Same conventions as the application: pf by type, efficiency for DOL / RDOL / SS motors given in kW, voltage per Notes.
  const motor = `OR(F${r}="DOL",F${r}="RDOL",F${r}="SS")`;
  const pf = `IF(F${r}="VFD",${N.pfVfd},IF(${motor},${N.pfDol},${N.pfGen}))`;
  const eff = `IF(${motor},${N.eff},1)`;
  set(8, { formula: `IF(ISNUMBER(D${r}),D${r},IF(ISNUMBER(C${r}),C${r}*1000/(SQRT(3)*${N.volt}*${pf}*${eff}),""))` } as ExcelJS.CellValue, num0, CALC);
  set(9, { formula: `IF(ISNUMBER(H${r}),SQRT(3)*${N.volt}*H${r}/1000,"")` } as ExcelJS.CellValue, num0, CALC);
  set(10, { formula: `IF(AND(E${r}="W",ISNUMBER(I${r})),I${r},0)` } as ExcelJS.CellValue, num0, CALC);
  set(11, { formula: `IF(AND(E${r}="W",ISNUMBER(H${r})),H${r},0)` } as ExcelJS.CellValue, num0, CALC);
  set(12, { formula: `J${r}*${N.lce}` } as ExcelJS.CellValue, num0, CALC);
  set(13, { formula: `IF(ISNUMBER(H${r}),H${r}*${N.mult},"")` } as ExcelJS.CellValue, num0, CALC);
  set(14, x.breaker ? breakerLabel(x.breaker) : "", undefined, CALC);
  set(15, l.note ?? "", undefined, INPUT);
  set(16, l.statusText ?? "", undefined, INPUT);
}

function buildNotes(ws: ExcelJS.Worksheet, set: ImccDesignSet, a: MccAssumptions) {
  ws.getColumn(1).width = 52; ws.getColumn(2).width = 14; ws.getColumn(3).width = 100;
  ws.getCell("A1").value = "Notes: IMCC design"; ws.getCell("A1").font = F({ bold: true, size: 16 });
  ws.getCell("A2").value = "Edit the assumption values; the IMCC sheets' formulas follow them. The breaker columns are values from the application's selection and do not change with these cells."; ws.getCell("A2").font = F({ italic: true });
  ["Assumption", "Value", "Used for"].forEach((t, i) => { const c = ws.getCell(3, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; });
  const calcVolt = Math.round(set.voltageV) === 433 ? a.calcVoltage433 : set.voltageV;
  const rows: [string, number, string, string?][] = [
    ["Power factor, general / power feeder", a.pfGeneral, "Current of PWR FDR rows given in kW."],
    ["Power factor, DOL / RDOL / SS", a.pfDol, "Current of DOL, RDOL and soft-starter (SS) motors given in kW."],
    ["Power factor, VFD", a.pfVfd, "Current of VFD motors given in kW."],
    ["Motor efficiency, DOL / RDOL / SS (provisional)", a.efficiency, "Applied to DOL, RDOL and SS motors given in kW only.", "0%"],
    ["Design current multiplier", a.designMultiplier, "Design current = working current x multiplier; the incomer limit applies to it."],
    ["Voltage used in the current formula (V)", calcVolt, "433 V loads are calculated at the value set in the application (433 V by default)."],
    ["LCE factor", LCE_FACTOR, "Reference workbook: LCE kVA = working kVA x 0.8. Meaning not confirmed (OPEN_ITEMS.md); reported only, never used for sizing."],
  ];
  rows.forEach(([k, v, u, fmt], i) => {
    const r = 4 + i;
    ws.getCell(r, 1).value = k; ws.getCell(r, 2).value = v; ws.getCell(r, 3).value = u;
    ws.getCell(r, 2).fill = fill("FFFFF2CC"); if (fmt) ws.getCell(r, 2).numFmt = fmt;
    for (let c = 1; c <= 3; c++) { ws.getCell(r, c).border = BORDER; ws.getCell(r, c).font = F(); }
  });
  const text: [string, string][] = [
    ["Design script", `Working loads in list order fill bus sections until a section's design current would exceed ${set.maxIncomerA} A; the next load starts a new section. Two sections make one IMCC, so the number of IMCCs follows from the loads.`],
    ["Status", "W and w = working (counted, used for the design). S = standby and Spare (SP) are kept with their group's section but not counted."],
    ["Operating philosophy", "Exactly one IMCC incomer is ON at a time; the other incomer and the bus coupler stay OFF. One incomer is never assumed to supply both sections through the coupler."],
    ["Incomer ratings", "Incomer 1 = calculated from its own section's working load. Incomer 2 and the bus coupler = Incomer 1's calculated design rating (calculated, not nameplate or installed-device ratings)."],
    ["Automatic PMCC feeders", "When the switch is ON, each incomer is an outgoing feeder of the PMCC load list (Incomer 1 = W, Incomer 2 = S, same group). When OFF the IMCC design is independent."],
    ["Soft starter (SS)", "Calculated like a DOL motor (assumption, OPEN_ITEMS.md)."],
    ["Reference workbook", "The layout follows the supplied ECR-9 / ECR-10 sheets. Their current formula (415 V, 1.73) is not used: this application's conventions are."],
    ["Open items", "See OPEN_ITEMS.md."],
  ];
  let r = 13;
  for (const [k, v] of text) { ws.getCell(r, 1).value = k; ws.getCell(r, 1).font = F({ bold: true }); ws.getCell(r, 3).value = v; ws.getCell(r, 3).font = F(); ws.getCell(r, 3).alignment = { wrapText: true, vertical: "top" }; r++; }
}

export async function imccWorkbookBuffer(set: ImccDesignSet, a?: MccAssumptions, o?: ImccWorkbookOptions): Promise<ArrayBuffer> {
  const wb = buildImccWorkbook(set, a, o);
  wb.creator = ""; wb.lastModifiedBy = "";
  const clean = stripAuthor(new Uint8Array(await wb.xlsx.writeBuffer()));
  return clean.buffer.slice(clean.byteOffset, clean.byteOffset + clean.byteLength) as ArrayBuffer;
}
