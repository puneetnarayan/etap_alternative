/**
 * Excel output for IMCC and PMCC designs (docs/imcc.md). ExcelJS, loaded only when a workbook is generated.
 *
 * IMCC: one load-list sheet in the layout of the reference workbook (S. NO., description, kW, current, W/S/SP,
 * method, group, kVA, working kVA, working current, LCE kVA, notes) with formulas, section totals and the incomer /
 * bus-coupler ratings, plus a Notes sheet that holds the assumptions the formulas use.
 *
 * PMCC: the existing MCC Load List workbook (mcc-workbook.ts), built from the PMCC load list with the generated IMCC
 * incomer feeders added as rows (Source = "Auto from IMCC"), plus an "IMCC links" sheet. A PMCC with no IMCC is
 * exported exactly as before.
 */
import ExcelJS from "exceljs";
import { DEFAULT_MCC_ASSUMPTIONS, breakerLabel, type MccAssumptions } from "./mcc-breaker";
import { buildMccWorkbook, stripAuthor } from "./mcc-workbook";
import { LCE_FACTOR, computeImcc, pmccExportSheet, type Imcc, type LoadResult, type DesignProject, type SectionResult } from "./project";

const THIN = { style: "thin" as const, color: { argb: "FFBFBFBF" } };
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const fill = (argb: string) => ({ type: "pattern" as const, pattern: "solid" as const, fgColor: { argb } });
const HEADER = "FFFCE4D6"; // light orange, as the other reports
const INPUT = "FFDDEBF7";
const CALC = "FFE2EFDA";
const SPARE = "FFFFD966";
const BANNER = "FFFFF2CC";
const F = (extra: Partial<ExcelJS.Font> = {}): Partial<ExcelJS.Font> => ({ name: "Calibri", size: 11, ...extra });

const toBuffer = async (wb: ExcelJS.Workbook): Promise<ArrayBuffer> => {
  wb.creator = ""; wb.lastModifiedBy = "";
  const clean = stripAuthor(new Uint8Array(await wb.xlsx.writeBuffer()));
  return clean.buffer.slice(clean.byteOffset, clean.byteOffset + clean.byteLength) as ArrayBuffer;
};

/* ---------------------------------- IMCC ---------------------------------- */

// Notes sheet cells the formulas read.
const N = { pfGen: "Notes!$B$4", pfDol: "Notes!$B$5", pfVfd: "Notes!$B$6", eff: "Notes!$B$7", mult: "Notes!$B$8", volt: "Notes!$B$9", lce: "Notes!$B$10" };

export function buildImccWorkbook(project: DesignProject, imccId: string, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): ExcelJS.Workbook {
  const imcc = project.imccs.find((i) => i.id === imccId);
  if (!imcc) throw new Error("The IMCC does not exist.");
  const res = computeImcc(imcc, a);
  const wb = new ExcelJS.Workbook();
  wb.creator = ""; wb.lastModifiedBy = "";
  wb.calcProperties.fullCalcOnLoad = true;
  const ws = wb.addWorksheet(imcc.tag.slice(0, 28) || "IMCC", { views: [{ state: "frozen", ySplit: 12, showGridLines: false }] });
  const notes = wb.addWorksheet("Notes", { views: [{ showGridLines: false }] });
  const widths = [7, 46, 9, 12, 9, 11, 30, 11, 10, 11, 12, 10, 12, 20, 60, 12];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const pmccTag = (id: string) => project.pmccs.find((p) => p.id === id)?.tag ?? "(missing PMCC)";

  ws.getCell("A1").value = `${imcc.tag} – IMCC load list`;
  ws.getCell("A1").font = F({ bold: true, size: 16 });
  ws.getCell("A2").value = "Preliminary engineering aid, not a substitute for approved design. Blue = input, green = calculated. Assumptions: Notes sheet.";
  ws.getCell("A2").font = F({ italic: true });
  const info: [string, string][] = [
    ["IMCC tag", imcc.tag],
    ["Description", imcc.description],
    ["Incomer 1 supplied from", `${pmccTag(imcc.incomer1.pmccId)} – feeder ${imcc.incomer1.feederRef}`],
    ["Incomer 2 supplied from", `${pmccTag(imcc.incomer2.pmccId)} – feeder ${imcc.incomer2.feederRef}`],
    ["Operating philosophy", `Exactly one incomer ON at a time (normally Incomer ${imcc.normallyOn}); the other incomer and the bus coupler stay OFF. Voltage ${imcc.voltageV} V.`],
  ];
  info.forEach(([k, v], i) => {
    const r = 4 + i;
    ws.getCell(r, 1).value = k; ws.getCell(r, 1).font = F({ bold: true });
    ws.mergeCells(r, 1, r, 2);
    ws.getCell(r, 3).value = v; ws.getCell(r, 3).font = F();
    ws.mergeCells(r, 3, r, 15);
  });
  ws.getCell("A10").value = `Operating state: ${res.operating.ok ? "normal state valid" : res.operating.errors.join(" ")}`;
  ws.getCell("A10").font = F({ italic: true });

  const HEAD = 12;
  const heads = ["S. NO.", "Load description", "kW", "Current entered (A)", "W / S / SP", "Method of starting", "Group", "Current (A)", "kVA",
    "Working kVA", "Working current (A)", `LCE kVA (x ${LCE_FACTOR})`, "Design current (A)", "Calculated breaker", "Notes / feeder description", "Status as entered"];
  heads.forEach((t, i) => {
    const c = ws.getCell(HEAD, i + 1);
    c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(HEAD).height = 34;

  let row = HEAD + 1;
  const totals: { section: SectionResult; row: number }[] = [];
  for (const sec of res.sections) {
    const src = sec.section === 1 ? imcc.incomer1 : imcc.incomer2;
    ws.mergeCells(row, 1, row, 15);
    const b = ws.getCell(row, 1);
    b.value = `INCOMER ${sec.section} – supplied from ${pmccTag(src.pmccId)}, feeder ${src.feederRef}`;
    b.font = F({ bold: true }); b.fill = fill(BANNER); b.border = BORDER;
    row++;
    const first = row;
    for (const x of sec.loads) { writeLoad(ws, row, x); row++; }
    const last = row - 1;
    const sumF = (col: string) => (last >= first ? { formula: `SUM(${col}${first}:${col}${last})` } : 0);
    ws.getCell(row, 2).value = `Total – Incomer ${sec.section} section`;
    for (const [col, ci] of [["J", 10], ["K", 11], ["L", 12]] as const) { const c = ws.getCell(row, ci); c.value = sumF(col) as ExcelJS.CellValue; c.numFmt = "#,##0"; }
    ws.getCell(row, 13).value = { formula: `K${row}*${N.mult}` } as ExcelJS.CellValue; ws.getCell(row, 13).numFmt = "#,##0";
    ws.getCell(row, 14).value = sec.workingA > 0 ? breakerLabel(sec.breaker) : "";
    ws.getCell(row, 15).value = "Design current = working current x design multiplier; breaker from the application's standard ratings.";
    for (let c = 1; c <= 16; c++) { const cell = ws.getCell(row, c); cell.font = F({ bold: true }); cell.border = BORDER; cell.fill = fill(CALC); }
    totals.push({ section: sec, row });
    row += 2;
  }

  // incomer / coupler ratings and overall totals
  ws.mergeCells(row, 1, row, 15);
  const sh = ws.getCell(row, 1);
  sh.value = "INCOMER, BUS COUPLER AND OVERALL (design ratings are calculated, not nameplate or installed-device ratings)";
  sh.font = F({ bold: true }); sh.fill = fill(BANNER); sh.border = BORDER;
  row++;
  const summary: [string, string, string][] = [
    ["Incomer 1", res.incomer1.ratingA ? breakerLabel(res.incomer1) : "Not calculated", "Calculated from the Incomer 1 section's working load."],
    ["Incomer 2", res.incomer2.ratingA ? breakerLabel(res.incomer2) : "Not calculated", "Design rating equals Incomer 1 (rule)."],
    ["Bus coupler", res.coupler.ratingA ? breakerLabel(res.coupler) : "Not calculated", "Design rating equals Incomer 1 (rule). Stays OFF in normal operation."],
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
    const c = ws.getCell(row, ci);
    c.value = { formula: totals.map((t) => `${col}${t.row}`).join("+") } as ExcelJS.CellValue; c.numFmt = "#,##0";
  }
  ws.getCell(row, 15).value = "Information only: under the operating philosophy the two sections are not supplied together, so this is not one source's demand.";
  for (let c = 1; c <= 15; c++) { ws.getCell(row, c).border = BORDER; ws.getCell(row, c).font = F({ bold: c !== 15 }); }
  row += 2;
  for (const i of res.issues) { ws.getCell(row, 2).value = `${i.severity === "error" ? "Error" : "Check"}: ${i.message}`; ws.getCell(row, 2).font = F({ color: { argb: i.severity === "error" ? "FFC00000" : "FF7F6000" } }); row++; }

  buildImccNotes(notes, imcc, a);
  return wb;
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
  // Same conventions as the application: pf by type, efficiency for DOL motors given in kW, voltage per Notes.
  const pf = `IF(F${r}="VFD",${N.pfVfd},IF(OR(F${r}="DOL",F${r}="RDOL"),${N.pfDol},${N.pfGen}))`;
  const eff = `IF(OR(F${r}="DOL",F${r}="RDOL"),${N.eff},1)`;
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

function buildImccNotes(ws: ExcelJS.Worksheet, imcc: Imcc, a: MccAssumptions) {
  ws.getColumn(1).width = 52; ws.getColumn(2).width = 14; ws.getColumn(3).width = 100;
  ws.getCell("A1").value = `Notes: ${imcc.tag}`; ws.getCell("A1").font = F({ bold: true, size: 16 });
  ws.getCell("A2").value = "Edit the assumption values; the Load list formulas follow them. The breaker columns are values from the application's selection and do not change with these cells."; ws.getCell("A2").font = F({ italic: true });
  ["Assumption", "Value", "Used for"].forEach((t, i) => { const c = ws.getCell(3, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; });
  const rows: [string, number, string, string?][] = [
    ["Power factor, general / power feeder", a.pfGeneral, "Current of PWR FDR rows given in kW."],
    ["Power factor, DOL / RDOL", a.pfDol, "Current of DOL and RDOL motors given in kW."],
    ["Power factor, VFD", a.pfVfd, "Current of VFD motors given in kW."],
    ["Motor efficiency, DOL / RDOL (provisional)", a.efficiency, "Applied to DOL and RDOL motors given in kW only.", "0%"],
    ["Design current multiplier", a.designMultiplier, "Design current = working current x multiplier."],
    ["Voltage used in the current formula (V)", Math.round(imcc.voltageV) === 433 ? a.calcVoltage433 : imcc.voltageV, "433 V loads are calculated at the value set in the application (433 V by default)."],
    ["LCE factor", LCE_FACTOR, "Reference workbook: LCE kVA = working kVA x 0.8. Meaning not confirmed (OPEN_ITEMS.md); reported only, never used for sizing."],
  ];
  rows.forEach(([k, v, u, fmt], i) => {
    const r = 4 + i;
    ws.getCell(r, 1).value = k; ws.getCell(r, 2).value = v; ws.getCell(r, 3).value = u;
    ws.getCell(r, 2).fill = fill("FFFFF2CC"); if (fmt) ws.getCell(r, 2).numFmt = fmt;
    for (let c = 1; c <= 3; c++) { ws.getCell(r, c).border = BORDER; ws.getCell(r, c).font = F(); }
  });
  const text: [string, string][] = [
    ["Status", "W and w = working (counted). S = standby (not counted in the normal working load). Spare = full spare feeder, shown as SP, not counted. Anything else is flagged."],
    ["Operating philosophy", "Exactly one IMCC incomer is ON at a time; the other incomer and the bus coupler stay OFF. One incomer is never assumed to supply both sections through the coupler."],
    ["Incomer ratings", "Incomer 1 = calculated from its own section's working load. Incomer 2 and the bus coupler = Incomer 1's calculated design rating. These are calculated design ratings, not nameplate or installed-device ratings."],
    ["PMCC link", "Each incomer is an outgoing feeder of its upstream PMCC (generated, linked record). The PMCC feeder is sized from its own section's current with the same rules, not copied from the incomer rating."],
    ["Reference workbook", "The layout follows the supplied ECR-9 / ECR-10 sheets. Their current formula (415 V, 1.73) is not used: this application's conventions are."],
    ["Open items", "See OPEN_ITEMS.md."],
  ];
  let r = 13;
  for (const [k, v] of text) { ws.getCell(r, 1).value = k; ws.getCell(r, 1).font = F({ bold: true }); ws.getCell(r, 3).value = v; ws.getCell(r, 3).font = F(); ws.getCell(r, 3).alignment = { wrapText: true, vertical: "top" }; r++; }
}

export async function imccWorkbookBuffer(project: DesignProject, imccId: string, a?: MccAssumptions): Promise<ArrayBuffer> {
  return toBuffer(buildImccWorkbook(project, imccId, a));
}

/* ---------------------------------- PMCC ---------------------------------- */

export async function pmccWorkbookBuffer(project: DesignProject, pmccId: string, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): Promise<ArrayBuffer> {
  const pmcc = project.pmccs.find((p) => p.id === pmccId);
  if (!pmcc) throw new Error("The PMCC does not exist.");
  const sheet = pmccExportSheet(project, pmccId, a);
  if (!sheet) throw new Error(`${pmcc.tag} has no load list yet.`);
  const { workbook: wb } = buildMccWorkbook(sheet as never, pmcc.fileName ?? pmcc.tag, a);
  const links = project.generated.filter((g) => g.pmccId === pmccId);
  if (links.length) {
    const ws = wb.addWorksheet("IMCC links", { views: [{ showGridLines: false }] });
    ws.getCell("A1").value = `${pmcc.tag}: outgoing feeders to IMCCs`; ws.getCell("A1").font = F({ bold: true, size: 16 });
    ws.getCell("A2").value = "Generated from the IMCC designs (linked records, not typed rows). Change the IMCC and this list follows."; ws.getCell("A2").font = F({ italic: true });
    const h = ["Feeder reference", "Outgoing feeder description", "Destination IMCC", "Incomer", "Status", "Working current of that section (A)", "Working kVA of that section", "Feeder design current (A)", "PMCC feeder rating (calculated)", "IMCC incomer design rating", "Source"];
    h.forEach((t, i) => { const c = ws.getCell(4, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; });
    ws.getRow(4).height = 48;
    [18, 30, 18, 9, 9, 18, 16, 16, 20, 20, 16].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    links.forEach((g, i) => {
      const imcc = project.imccs.find((x) => x.id === g.imccId)!;
      const r = computeImcc(imcc, a);
      const inc = g.incomer === 1 ? r.incomer1 : r.incomer2;
      const v: ExcelJS.CellValue[] = [g.feederRef, g.description, g.destination, g.incomer, g.status, g.ratedA, g.kva, g.designA ?? null, g.rating?.ratingA ? breakerLabel(g.rating) : "", inc.ratingA ? breakerLabel(inc) : "", "Auto from IMCC"];
      v.forEach((x, c) => { const cell = ws.getCell(5 + i, c + 1); cell.value = x; cell.border = BORDER; cell.font = F(); if (typeof x === "number" && c >= 5) cell.numFmt = "#,##0"; });
    });
    ws.getCell(6 + links.length, 1).value = "Alternative sources: the two incomers of an IMCC are not added together as normal demand. The one normally ON is W, the other S (standby).";
    ws.getCell(6 + links.length, 1).font = F({ italic: true });
  }
  return toBuffer(wb);
}
