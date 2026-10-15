/**
 * Excel output for the IMCC design (docs/imcc.md). ExcelJS, loaded only when a workbook is generated.
 *
 * One workbook: a Summary sheet, one sheet per IMCC (panel order: Incomer 1 | loads | Incomer 2 | Bus coupler; loading block,
 * formulas, ON green / OFF light red) and a Notes sheet with the assumptions the formulas read.
 * Same look as the other workbooks of the application: Calibri 11, light-orange headers, blue input / green calculated.
 */
import ExcelJS from "exceljs";
import { DEFAULT_MCC_ASSUMPTIONS, breakerLabel, type MccAssumptions } from "./mcc-breaker";
import { stripAuthor } from "./mcc-workbook";
import { FAILOVER_STATE, LCE_FACTOR, NORMAL_STATE, imccFeederRows, loadTypeOf, type ImccDesign, type ImccDesignSet, type LoadResult } from "./imcc";
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
  for (const i of set.imccs) buildImccSheet(wb.addWorksheet(i.tag, { views: [{ showGridLines: false }] }), i, set, o);
  buildNotes(wb.addWorksheet("Notes", { views: [{ showGridLines: false }] }), set, a);
  return wb;
}

/** ON = light green, OFF = light red. */
const stateFill = (v: string) => fill(v === "ON" ? WITHIN : OVER);
const pctText = (v: number | undefined) => (v === undefined ? "" : v);

function buildSummary(ws: ExcelJS.Worksheet, set: ImccDesignSet, o: ImccWorkbookOptions) {
  [6, 12, 16, 14, 14, 18, 14, 18, 60].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  ws.getCell("A1").value = "IMCC design – summary"; ws.getCell("A1").font = F({ bold: true, size: 16 });
  ws.getCell("A2").value = "Preliminary engineering aid, not a substitute for approved design. Only working (W / w) loads are used; the number of IMCCs follows from the loads."; ws.getCell("A2").font = F({ italic: true });
  const info: [string, string][] = [
    ["Source", o.sourceName ?? "IMCC load list"],
    ["Limit per IMCC", `${set.maxDesignA.toFixed(0)} A design current (working current x design multiplier): ${set.maxIncomerA} A ACB at ${(set.acbMaxLoading * 100).toFixed(0)} % maximum loading`],
    ["IMCCs", String(set.imccs.length)],
    ["Automatic PMCC feeders", o.linked ? "ON: each incomer is an outgoing feeder of the PMCC list (see below)" : "OFF: the IMCC design is independent; nothing is added to the PMCC"],
    ["Operating philosophy", "Normal: Incomer 1 ON, Incomer 2 and bus coupler OFF. Incomer 1 unhealthy: Incomer 2 and bus coupler both ON, with an interlock that switches Incomer 1 OFF first."],
  ];
  info.forEach(([k, v], i) => {
    ws.getCell(4 + i, 1).value = k; ws.getCell(4 + i, 1).font = F({ bold: true }); ws.mergeCells(4 + i, 1, 4 + i, 3);
    ws.getCell(4 + i, 4).value = v; ws.getCell(4 + i, 4).font = F(); ws.mergeCells(4 + i, 4, 4 + i, 9);
  });
  const H = 10;
  ["", "IMCC", "Loads (W / S / SP)", "Working current (A)", "Design current (A)", "Incomer 1, Incomer 2 and bus coupler rating", "Loading of the rating", "Loading limit", "Checks"].forEach((t, i) => {
    const c = ws.getCell(H, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(H).height = 48;
  set.imccs.forEach((i, n) => {
    const b = i.bus;
    const v: ExcelJS.CellValue[] = [n + 1, i.tag, `${b.working} / ${b.standby} / ${b.spare}`, b.workingA, b.designA, i.incomer1.ratingA ? breakerLabel(i.incomer1) : "Not calculated",
      pctText(i.loading), i.loadingLimit, i.issues.map((x) => x.message).join(" ")];
    v.forEach((x, c) => {
      const cell = ws.getCell(H + 1 + n, c + 1); cell.value = x; cell.border = BORDER; cell.font = F(); cell.alignment = { vertical: "top", wrapText: c === 8 };
      if (typeof x === "number" && c > 2) cell.numFmt = c >= 6 ? "0%" : "#,##0";
      if (c === 6 && typeof x === "number") cell.fill = fill(x > i.loadingLimit + 1e-9 ? OVER : WITHIN);
      if (c === 1) cell.fill = fill(INCOMER);
    });
  });
  let r = H + set.imccs.length + 3;
  if (o.linked) {
    ws.getCell(r, 1).value = "PMCC outgoing feeders generated for the incomers (switch ON)"; ws.getCell(r, 1).font = F({ bold: true, size: 13 }); r++;
    ["", "Outgoing feeder", "Status", "Current (A)", "kVA", "Group"].forEach((t, i) => { const c = ws.getCell(r, i + 1); c.value = t; c.font = F({ bold: true }); c.fill = fill(HEADER); c.border = BORDER; });
    r++;
    for (const f of imccFeederRows(set)) {
      [r, f.description, f.status, f.amps, f.kva, f.group].forEach((x, i) => {
        if (i === 0) return;
        const c = ws.getCell(r, i + 1); c.value = x as ExcelJS.CellValue; c.border = BORDER; c.font = F(); if (typeof x === "number") c.numFmt = "#,##0";
        if (i === 2) c.fill = fill(kind(String(x))!);
      });
      r++;
    }
    ws.getCell(r + 1, 1).value = "Only one incomer carries the IMCC at a time: Incomer 1 is W, Incomer 2 is S (it takes over if Incomer 1 is unhealthy). They share a group so they land on different transformer sections.";
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
  const feeder = (n: 1 | 2) => (o.linked ? (imcc.bus.workingA > 0 ? `PMCC outgoing feeder "${imcc.tag} Incomer ${n}"` : "no PMCC feeder (no working load)") : "PMCC feeder not generated (Automatic PMCC feeders switch OFF)");
  const info: [string, string][] = [
    ["IMCC tag", imcc.tag],
    ["Design basis", `Working (W / w) loads only. Limit ${set.maxDesignA.toFixed(0)} A design current: ${set.maxIncomerA} A ACB at ${(set.acbMaxLoading * 100).toFixed(0)} % maximum loading. Voltage ${set.voltageV} V.`],
    ["Incomer 1 supplied from", feeder(1)],
    ["Incomer 2 supplied from", feeder(2)],
    ["Panel order", "Incomer 1 at the beginning; Incomer 2 and the bus coupler side by side at the other end."],
    ["Operating philosophy", "Normal: Incomer 1 ON, Incomer 2 and bus coupler OFF. Incomer 1 unhealthy: Incomer 2 and bus coupler both ON, interlocked so Incomer 1 is switched OFF first."],
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

  // ---- row positions: INC 1 | loads | INC 2 | BC (known first: the loading block refers to them) ----
  const SUMMARY = 12; // header row of the loading block
  const LEGEND = 17;
  const HEAD = 20;
  const bus = imcc.bus;
  const inc1 = HEAD + 1;
  const first = inc1 + 1, last = first + bus.loads.length - 1;
  const inc2 = last + 1;
  const coup = inc2 + 1;
  const limit = imcc.loadingLimit;

  // ---- loading block ----
  ws.getCell(SUMMARY - 1, 1).value = "Loading"; ws.getCell(SUMMARY - 1, 1).font = F({ bold: true, size: 13 });
  head(SUMMARY, ["", "Element", "", "", "Normal state", "Incomer 1 unhealthy", "", "", "", "Working kVA", "Working current (A)", "", "Design current (A)", "Calculated rating", "Note", "", "Rating (A)", "Loading (% of rating)"]);
  ws.getRow(SUMMARY).height = 34;
  const lr = (n: number, label: string, color: string, normal: SwitchLike, fail: SwitchLike, srcRow: number, rating: Parameters<typeof breakerLabel>[0], note: string, loading: boolean) => {
    const r = SUMMARY + n;
    bar(r, color);
    ws.getCell(r, 2).value = label;
    for (const [ci, st] of [[5, normal], [6, fail]] as const) { const c = ws.getCell(r, ci); c.value = st; c.fill = stateFill(st); c.alignment = { horizontal: "center" }; }
    for (const [col, ci] of [["J", 10], ["K", 11], ["M", 13]] as const) { const c = ws.getCell(r, ci); c.value = { formula: `${col}${srcRow}` } as ExcelJS.CellValue; c.numFmt = "#,##0"; }
    ws.getCell(r, 14).value = rating.ratingA ? breakerLabel(rating) : "Not calculated";
    ws.getCell(r, 15).value = note;
    ws.getCell(r, 17).value = rating.ratingA ?? null; ws.getCell(r, 17).numFmt = "#,##0";
    if (loading) { const c = ws.getCell(r, 18); c.value = { formula: `IF(N(Q${r})>0,M${r}/Q${r},"")` } as ExcelJS.CellValue; c.numFmt = "0.0%"; c.fill = fill(WITHIN); }
  };
  lr(1, "Incomer 1", INCOMER, "ON", "OFF", inc1, imcc.incomer1, "Carries the whole IMCC in the normal state.", true);
  lr(2, "Incomer 2", INCOMER, "OFF", "ON", inc1, imcc.incomer2, "Takes over the whole IMCC (with the bus coupler) if Incomer 1 is unhealthy: same rating as Incomer 1.", true);
  lr(3, "Bus coupler", COUPLER, "OFF", "ON", inc1, imcc.coupler, "Switched ON together with Incomer 2: same rating as Incomer 1.", true);
  const cf = (ref: string, priority: number) => ws.addConditionalFormatting({
    ref, rules: [{ type: "cellIs", operator: "greaterThan", formulae: [limit], priority, style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: OVER } } } }],
  });
  cf(`R${SUMMARY + 1}:R${SUMMARY + 3}`, 1);
  ws.getCell(SUMMARY + 4, 2).value = `Loading limit: ${(limit * 100).toFixed(0)} % of the rating (${imcc.incomer1.kind === "EDO ACB" ? "ACB maximum loading" : "MCCB: 100 %"}). Red = above the limit.`;
  ws.getCell(SUMMARY + 4, 2).font = F({ italic: true });

  // ---- legend ----
  ws.getCell(LEGEND, 1).value = "Colours"; ws.getCell(LEGEND, 1).font = F({ bold: true });
  ([["Incomer", INCOMER], ["Bus coupler", COUPLER], ["ON", WITHIN], ["OFF", OVER], ["W", kind("W")!], ["S", kind("S")!], ["SP (spare)", SPARE], ["DOL / SS", kind("DOL")!], ["VFD", kind("VFD")!], ["Feeder", kind("Feeder")!]] as [string, string][]).forEach(([t, color], i) => {
    const c = ws.getCell(LEGEND, 2 + i); c.value = t; c.fill = fill(color); c.border = BORDER; c.font = F({ bold: true }); c.alignment = { horizontal: "center" };
  });

  // ---- panel: INC 1 | loads | INC 2 | BC ----
  head(HEAD, ["S. NO.", "Load description", "kW", "Current entered (A)", "W / S / SP", "Method of starting", "Group", "Current (A)", "kVA",
    "Working kVA", "Working current (A)", `LCE kVA (x ${LCE_FACTOR})`, "Design current (A)", "Calculated breaker", "Notes / feeder description", "Status as entered", "Rating (A)", "Loading (% of rating)"]);
  ws.getRow(HEAD).height = 34;
  const totalsRow = (r: number, label: string, tag: string, st: SwitchLike, rating: Parameters<typeof breakerLabel>[0], color: string, note: string, sums: boolean) => {
    bar(r, color);
    ws.getCell(r, 1).value = tag;
    ws.getCell(r, 2).value = label;
    ws.getCell(r, 5).value = st; ws.getCell(r, 5).fill = stateFill(st); ws.getCell(r, 5).alignment = { horizontal: "center" };
    for (const [col, ci] of [["J", 10], ["K", 11], ["L", 12]] as const) {
      const c = ws.getCell(r, ci);
      c.value = (sums ? (last >= first ? { formula: `SUM(${col}${first}:${col}${last})` } : 0) : { formula: `${col}${inc1}` }) as ExcelJS.CellValue; c.numFmt = "#,##0";
    }
    ws.getCell(r, 13).value = { formula: `K${r}*${N.mult}` } as ExcelJS.CellValue; ws.getCell(r, 13).numFmt = "#,##0";
    ws.getCell(r, 14).value = rating.ratingA ? breakerLabel(rating) : "";
    ws.getCell(r, 15).value = note;
    ws.getCell(r, 17).value = rating.ratingA ?? null; ws.getCell(r, 17).numFmt = "#,##0";
    const c = ws.getCell(r, 18); c.value = { formula: `IF(N(Q${r})>0,M${r}/Q${r},"")` } as ExcelJS.CellValue; c.numFmt = "0.0%"; c.fill = fill(WITHIN);
  };
  totalsRow(inc1, `Incomer 1 – supplied from ${feeder(1)}`, "INC 1", "ON", imcc.incomer1, INCOMER, "Beginning of the panel. Totals of the loads below; rating from the working load of the whole IMCC.", true);
  bus.loads.forEach((x, n) => writeLoad(ws, first + n, x));
  totalsRow(inc2, `Incomer 2 – supplied from ${feeder(2)}`, "INC 2", "OFF", imcc.incomer2, INCOMER, "Other end of the panel, next to the bus coupler. Same totals and rating as Incomer 1.", false);
  totalsRow(coup, "Bus coupler (side by side with Incomer 2)", "BC", "OFF", imcc.coupler, COUPLER, "Switched ON with Incomer 2 if Incomer 1 is unhealthy; interlock: Incomer 1 OFF first. Same rating as Incomer 1.", false);
  cf(`R${inc1}:R${coup}`, 2);

  let row = coup + 2;
  for (const i of imcc.issues) { ws.getCell(row, 2).value = `Check: ${i.message}`; ws.getCell(row, 2).font = F({ color: { argb: "FF7F6000" } }); row++; }
  ws.views = [{ showGridLines: false }];
}

type SwitchLike = typeof NORMAL_STATE.incomer1;
void FAILOVER_STATE;

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
    ["Power factor, general / power feeder", a.pfGeneral, "Current of Feeder rows given in kW."],
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
    ["Design script", `Working loads in list order fill an IMCC until its design current would exceed ${set.maxDesignA.toFixed(0)} A (${set.maxIncomerA} A ACB at ${(set.acbMaxLoading * 100).toFixed(0)} % maximum loading, adjustable in the application); the next load starts a new IMCC, so the number of IMCCs follows from the loads.`],
    ["Status", "W and w = working (counted, used for the design). S = standby and Spare (SP) are kept with their group's section but not counted."],
    ["Operating philosophy", "Panel: Incomer 1 at the beginning; Incomer 2 and the bus coupler side by side at the other end. Normal: Incomer 1 ON, Incomer 2 and bus coupler OFF. Incomer 1 unhealthy: Incomer 2 and bus coupler both ON, with an interlock that switches Incomer 1 OFF first. ON is shown light green, OFF light red."],
    ["Incomer ratings", "Incomer 1, Incomer 2 and the bus coupler each carry the whole IMCC, so they have the same calculated design rating (calculated, not nameplate or installed-device ratings). An ACB is loaded at most to the maximum loading above."],
    ["Automatic PMCC feeders", "When the switch is ON, each incomer is an outgoing feeder of the PMCC load list (Incomer 1 = W, Incomer 2 = S, same group, each with the IMCC working current). When OFF the IMCC design is independent."],
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
