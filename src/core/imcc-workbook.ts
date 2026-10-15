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
import { LCE_FACTOR, imccFeederRows, loadTypeOf, type ImccDesign, type ImccDesignSet, type LoadResult } from "./imcc";
import { REPORT_COLORS, kindColor } from "./report";

const THIN = { style: "thin" as const, color: { argb: "FFBFBFBF" } };
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const fill = (argb: string) => ({ type: "pattern" as const, pattern: "solid" as const, fgColor: { argb } });
// The colours of the other reports of the application (report.ts): light-orange headers, cyan incomers, light-yellow
// bus couplers, W lavender, S pink, spare gold, DOL / VFD / Feeder pastels, green within / red over a limit.
const argb = (hex: string) => `FF${hex.replace("#", "").toUpperCase()}`;
const HEADER = argb(REPORT_COLORS.header);
const INCOMER = argb(REPORT_COLORS.incomer);
const COUPLER = argb(REPORT_COLORS.coupler);
const WITHIN = argb(REPORT_COLORS.within);
const OVER = argb(REPORT_COLORS.over);
const SPARE = argb(REPORT_COLORS.spare);
const kind = (v: string) => { const c = kindColor(v); return c ? argb(c) : undefined; };
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
  [6, 12, 14, 14, 14, 18, 12, 14, 14, 18, 12, 18, 60].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
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
    ws.getCell(4 + i, 4).value = v; ws.getCell(4 + i, 4).font = F(); ws.mergeCells(4 + i, 4, 4 + i, 13);
  });
  const H = 10;
  ["", "IMCC", "Loads (W / S / SP)", "Incomer 1 working A", "Incomer 1 design A", "Incomer 1 rating", "Incomer 1 loading", "Incomer 2 working A", "Incomer 2 design A", "Incomer 2 rating (= Incomer 1)", "Incomer 2 loading", "Bus coupler (= Incomer 1)", "Checks"].forEach((t, i) => {
    const c = ws.getCell(H, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getColumn(12).width = 18; ws.getColumn(13).width = 60;
  ws.getRow(H).height = 34;
  set.imccs.forEach((i, n) => {
    const [s1, s2] = i.sections;
    const load = (sec: typeof s1) => (i.incomer1.ratingA ? sec.designA / i.incomer1.ratingA : undefined);
    const v: ExcelJS.CellValue[] = [n + 1, i.tag, `${s1.working + s2.working} / ${s1.standby + s2.standby} / ${s1.spare + s2.spare}`, s1.workingA, s1.designA,
      s1.workingA > 0 ? breakerLabel(i.incomer1) : "", load(s1) ?? "", s2.workingA, s2.designA, s2.workingA > 0 ? breakerLabel(i.incomer2) : "", s2.workingA > 0 ? load(s2) ?? "" : "", breakerLabel(i.coupler),
      i.issues.map((x) => x.message).join(" ")];
    v.forEach((x, c) => {
      const cell = ws.getCell(H + 1 + n, c + 1); cell.value = x; cell.border = BORDER; cell.font = F(); cell.alignment = { vertical: "top", wrapText: c === 12 };
      if (typeof x === "number" && c > 2) cell.numFmt = c === 6 || c === 10 ? "0.0%" : "#,##0";
      if ((c === 6 || c === 10) && typeof x === "number") cell.fill = fill(x > 1 + 1e-9 ? OVER : WITHIN);
      if (c === 1) cell.fill = fill(INCOMER);
      if (c === 11) cell.fill = fill(COUPLER);
    });
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

const COLS = 18; // A..R

function buildImccSheet(ws: ExcelJS.Worksheet, imcc: ImccDesign, set: ImccDesignSet, o: ImccWorkbookOptions) {
  [9, 46, 9, 12, 10, 12, 30, 11, 10, 11, 12, 10, 12, 20, 60, 12, 10, 11].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  ws.getCell("A1").value = `${imcc.tag} – IMCC load list and loading`; ws.getCell("A1").font = F({ bold: true, size: 16 });
  ws.getCell("A2").value = "Preliminary engineering aid, not a substitute for approved design. Assumptions: Notes sheet."; ws.getCell("A2").font = F({ italic: true });
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

  const head = (r: number, labels: string[]) => labels.forEach((t, i) => {
    const c = ws.getCell(r, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  const bar = (r: number, color: string, bold = true) => { for (let c = 1; c <= COLS; c++) { const cell = ws.getCell(r, c); cell.fill = fill(color); cell.border = BORDER; cell.font = F({ bold }); cell.alignment = { vertical: "middle", wrapText: c === 2 || c === 15 }; } };

  // ---- row positions of the bus layout (known first, the loading block refers to them) ----
  const SUMMARY = 11; // header row of the loading block
  const LEGEND = 16;
  const HEAD = 19;
  const [s1, s2] = imcc.sections;
  const inc1 = HEAD + 1;
  const first1 = inc1 + 1, last1 = first1 + s1.loads.length - 1;
  const coup = last1 + 1;
  const first2 = coup + 1, last2 = first2 + s2.loads.length - 1;
  const inc2 = last2 + 1;
  const mult = N.mult;

  // ---- loading block ----
  ws.getCell(SUMMARY - 1, 1).value = "Loading"; ws.getCell(SUMMARY - 1, 1).font = F({ bold: true, size: 13 });
  head(SUMMARY, ["", "Element", "", "", "Normal state", "", "", "", "", "Working kVA", "Working current (A)", "", "Design current (A)", "Calculated rating", "Note", "", "Rating (A)", "Loading (% of rating)"]);
  ws.getRow(SUMMARY).height = 34;
  const lr = (n: number, label: string, color: string, state: string, srcRow: number | null, rating: { ratingA?: number } & Parameters<typeof breakerLabel>[0], note: string, loading: boolean) => {
    const r = SUMMARY + n;
    bar(r, color);
    ws.getCell(r, 2).value = label; ws.getCell(r, 5).value = state; ws.getCell(r, 5).alignment = { horizontal: "center" };
    if (srcRow !== null) {
      for (const [col, ci] of [["J", 10], ["K", 11], ["M", 13]] as const) { const c = ws.getCell(r, ci); c.value = { formula: `${col}${srcRow}` } as ExcelJS.CellValue; c.numFmt = "#,##0"; }
    }
    ws.getCell(r, 14).value = rating.ratingA ? breakerLabel(rating) : "Not calculated";
    ws.getCell(r, 15).value = note;
    ws.getCell(r, 17).value = rating.ratingA ?? null; ws.getCell(r, 17).numFmt = "#,##0";
    if (loading && srcRow !== null) {
      const c = ws.getCell(r, 18);
      c.value = { formula: `IF(N(Q${r})>0,M${r}/Q${r},"")` } as ExcelJS.CellValue; c.numFmt = "0.0%";
      c.fill = fill(WITHIN); // green; red when over 100 % (conditional format below)
    }
  };
  lr(1, "Incomer 1", INCOMER, "ON", inc1, imcc.incomer1, `From the Incomer 1 section's working load (limit ${set.maxIncomerA} A).`, true);
  lr(2, "Incomer 2", INCOMER, "OFF", inc2, imcc.incomer2, "Design rating equals Incomer 1 (rule); loading is against that rating.", true);
  lr(3, "Bus coupler", COUPLER, "OFF", null, imcc.coupler, "Design rating equals Incomer 1 (rule). Stays OFF in normal operation.", false);
  ws.addConditionalFormatting({
    ref: `R${SUMMARY + 1}:R${SUMMARY + 2}`,
    rules: [{ type: "cellIs", operator: "greaterThan", formulae: [1], priority: 1, style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: OVER } } } }],
  });
  ws.getCell(SUMMARY + 4, 2).value = "Both sections summed (information only: the two sections are not supplied together)";
  for (const [col, ci] of [["J", 10], ["K", 11]] as const) { const c = ws.getCell(SUMMARY + 4, ci); c.value = { formula: `${col}${inc1}+${col}${inc2}` } as ExcelJS.CellValue; c.numFmt = "#,##0"; c.font = F({ bold: true }); }
  ws.getCell(SUMMARY + 4, 2).font = F({ italic: true });

  // ---- legend ----
  ws.getCell(LEGEND, 1).value = "Colours"; ws.getCell(LEGEND, 1).font = F({ bold: true });
  ([["Incomer", INCOMER], ["Bus coupler", COUPLER], ["W", kind("W")!], ["S", kind("S")!], ["SP (spare)", SPARE], ["DOL / SS", kind("DOL")!], ["VFD", kind("VFD")!], ["Feeder", kind("Feeder")!], ["Within limit", WITHIN], ["Over limit", OVER]] as [string, string][]).forEach(([t, color], i) => {
    const c = ws.getCell(LEGEND, 2 + i); c.value = t; c.fill = fill(color); c.border = BORDER; c.font = F({ bold: true }); c.alignment = { horizontal: "center" };
  });

  // ---- bus layout: incomer 1 | loads 1 | bus coupler | loads 2 | incomer 2 ----
  head(HEAD, ["S. NO.", "Load description", "kW", "Current entered (A)", "W / S / SP", "Method of starting", "Group", "Current (A)", "kVA",
    "Working kVA", "Working current (A)", `LCE kVA (x ${LCE_FACTOR})`, "Design current (A)", "Calculated breaker", "Notes / feeder description", "Status as entered", "Rating (A)", "Loading (% of rating)"]);
  ws.getRow(HEAD).height = 34;

  const incomerRow = (r: number, n: 1 | 2, sec: typeof s1, first: number, last: number, rating: typeof imcc.incomer1) => {
    bar(r, INCOMER);
    ws.getCell(r, 1).value = `INC ${n}`;
    ws.getCell(r, 2).value = `Incomer ${n} – supplied from ${feeder(n)}`;
    ws.getCell(r, 5).value = n === 1 ? "ON" : "OFF"; ws.getCell(r, 5).alignment = { horizontal: "center" };
    for (const [col, ci] of [["J", 10], ["K", 11], ["L", 12]] as const) {
      const c = ws.getCell(r, ci); c.value = (last >= first ? { formula: `SUM(${col}${first}:${col}${last})` } : 0) as ExcelJS.CellValue; c.numFmt = "#,##0";
    }
    ws.getCell(r, 13).value = { formula: `K${r}*${mult}` } as ExcelJS.CellValue; ws.getCell(r, 13).numFmt = "#,##0";
    ws.getCell(r, 14).value = sec.workingA > 0 || n === 2 ? (rating.ratingA ? breakerLabel(rating) : "") : "";
    ws.getCell(r, 15).value = n === 1 ? "Totals of the loads below, calculated rating from the working load." : "Totals of the loads above; design rating equals Incomer 1.";
    ws.getCell(r, 17).value = rating.ratingA ?? null; ws.getCell(r, 17).numFmt = "#,##0";
    const c = ws.getCell(r, 18); c.value = { formula: `IF(N(Q${r})>0,M${r}/Q${r},"")` } as ExcelJS.CellValue; c.numFmt = "0.0%";
  };
  incomerRow(inc1, 1, s1, first1, last1, imcc.incomer1);
  s1.loads.forEach((x, n) => writeLoad(ws, first1 + n, x));
  bar(coup, COUPLER);
  ws.getCell(coup, 1).value = "BC";
  ws.getCell(coup, 2).value = "Bus coupler between Incomer 1 and Incomer 2 sections";
  ws.getCell(coup, 5).value = "OFF"; ws.getCell(coup, 5).alignment = { horizontal: "center" };
  ws.getCell(coup, 14).value = imcc.coupler.ratingA ? breakerLabel(imcc.coupler) : "";
  ws.getCell(coup, 15).value = "Design rating equals Incomer 1 (rule). Stays OFF in normal operation.";
  ws.getCell(coup, 17).value = imcc.coupler.ratingA ?? null; ws.getCell(coup, 17).numFmt = "#,##0";
  s2.loads.forEach((x, n) => writeLoad(ws, first2 + n, x));
  incomerRow(inc2, 2, s2, first2, last2, imcc.incomer2);
  ws.addConditionalFormatting({
    ref: `R${inc1}:R${inc2}`,
    rules: [{ type: "cellIs", operator: "greaterThan", formulae: [1], priority: 2, style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: OVER } } } }],
  });
  for (const r of [inc1, inc2]) ws.getCell(r, 18).fill = fill(WITHIN);

  let row = inc2 + 2;
  for (const i of imcc.issues) { ws.getCell(row, 2).value = `Check: ${i.message}`; ws.getCell(row, 2).font = F({ color: { argb: "FF7F6000" } }); row++; }
  ws.views = [{ showGridLines: false }];
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
  const typeLabel = loadTypeOf(l.method) === "FEEDER" ? "Feeder" : loadTypeOf(l.method);
  set(1, l.sNo ?? "");
  set(2, l.description);
  set(3, l.kw ?? null, "#,##0.0");
  set(4, l.ratedA ?? null, num0);
  set(5, l.status ?? "", undefined, l.status === "SP" ? SPARE : l.status ? kind(l.status) : OVER);
  ws.getCell(r, 5).alignment = { horizontal: "center", vertical: "middle" };
  set(6, l.method ?? "", undefined, l.method ? kind(typeLabel) : undefined);
  set(7, l.group ?? "");
  // Same conventions as the application: pf by type, efficiency for DOL / RDOL / SS motors given in kW, voltage per Notes.
  const motor = `OR(F${r}="DOL",F${r}="RDOL",F${r}="SS")`;
  const pf = `IF(F${r}="VFD",${N.pfVfd},IF(${motor},${N.pfDol},${N.pfGen}))`;
  const eff = `IF(${motor},${N.eff},1)`;
  set(8, { formula: `IF(ISNUMBER(D${r}),D${r},IF(ISNUMBER(C${r}),C${r}*1000/(SQRT(3)*${N.volt}*${pf}*${eff}),""))` } as ExcelJS.CellValue, num0);
  set(9, { formula: `IF(ISNUMBER(H${r}),SQRT(3)*${N.volt}*H${r}/1000,"")` } as ExcelJS.CellValue, num0);
  set(10, { formula: `IF(AND(E${r}="W",ISNUMBER(I${r})),I${r},0)` } as ExcelJS.CellValue, num0);
  set(11, { formula: `IF(AND(E${r}="W",ISNUMBER(H${r})),H${r},0)` } as ExcelJS.CellValue, num0);
  set(12, { formula: `J${r}*${N.lce}` } as ExcelJS.CellValue, num0);
  set(13, { formula: `IF(ISNUMBER(H${r}),H${r}*${N.mult},"")` } as ExcelJS.CellValue, num0);
  set(14, x.breaker ? breakerLabel(x.breaker) : "");
  set(15, l.note ?? "");
  set(16, l.statusText ?? "");
  set(17, null); set(18, null);
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
