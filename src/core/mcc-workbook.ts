/**
 * Builds the "MCC Load List / Breaker Selection" workbook (real .xlsx with working formulas, filters,
 * a frozen header, drop-downs and cell notes). Rules: docs/mcc-load-list.md. Preliminary aid only.
 */
import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  ACB_THRESHOLD_A, CT_PRIMARIES_A, MCCB_RATINGS_A, DEFAULT_MCC_ASSUMPTIONS, breakerLabel, ctPrimaryFor, currentFor, selectBreaker,
  type MccAssumptions, type MccLoadType,
} from "./mcc-breaker";
import { parseLoadSheet } from "./transformer-sizing";

type Cell = unknown;

export interface MccWorkbookResult {
  workbook: ExcelJS.Workbook;
  /** Load rows read from the source sheet. */
  rows: number;
  /** Rows with no current available (nothing could be calculated). */
  missing: number;
}

const SOURCE = "FFDDEBF7"; // input data
const CALC = "FFE2EFDA"; // calculated
const PRELIM = "FFFFF2CC"; // preliminary selections
const ORANGE = "FFFFE0B2"; // title block / notes headers (same light orange as the other report)
const THIN = { style: "thin" as const, color: { argb: "FFBFBFBF" } };
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };

/** Light tint of an ARGB colour for alternating rows. */
const lighter = (argb: string) => {
  const c = [2, 4, 6].map((i) => parseInt(argb.slice(i, i + 2), 16));
  return "FF" + c.map((v) => Math.round(v + (255 - v) * 0.55).toString(16).padStart(2, "0")).join("").toUpperCase();
};

const letter = (n: number) => {
  let s = "";
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  return s;
};

const norm = (h: Cell) => String(h ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const blank = (c: Cell) => c === null || c === undefined || (typeof c === "string" && c.trim() === "");

function loadType(raw: string): MccLoadType {
  const t = raw.toUpperCase();
  if (t.includes("VFD")) return "VFD";
  if (t.includes("DOL")) return "DOL";
  return "FEEDER";
}

type Kind = "input" | "calc" | "select";
interface Col {
  key: string;
  header: string;
  kind: Kind;
  width: number;
  format?: string;
  comment?: string;
}

// Notes!B4:B10 hold the assumptions the formulas use.
const A = {
  pfGeneral: "Notes!$B$4", pfDol: "Notes!$B$5", pfVfd: "Notes!$B$6", eff: "Notes!$B$7",
  mult: "Notes!$B$8", calcV: "Notes!$B$9", threshold: "Notes!$B$10", mccb630: "Notes!$B$11",
} as const;

const COLS: Col[] = [
  { key: "type", header: "Load type read (DOL / VFD / Feeder)", kind: "input", width: 12, comment: "From the Type column: text containing VFD = VFD, containing DOL = DOL, anything else = general feeder." },
  { key: "volt", header: "Voltage read (V)", kind: "input", width: 10, format: "#,##0" },
  { key: "kw", header: "kW read", kind: "input", width: 9, format: "#,##0" },
  { key: "rated", header: "Estimated / entered load current (A)", kind: "input", width: 14, format: "#,##0", comment: "Amps rating from the load list. If present and valid it is used in preference to a current calculated from kW." },
  { key: "calcV", header: "Calculation voltage (V)", kind: "calc", width: 12, format: "#,##0", comment: "Line voltage used in the formula. The 433 V level is calculated at the value on the Notes sheet (433 V by default)." },
  { key: "pf", header: "Power factor used", kind: "calc", width: 9, format: "0.00", comment: "General feeder 0.7, DOL 0.7, VFD 0.9 (Notes sheet)." },
  { key: "eff", header: "Motor efficiency used (provisional)", kind: "calc", width: 12, format: "0%", comment: "Applied to DOL motors given in kW only (kW taken as motor output). Provisional default 93% (Notes sheet). Overwrite with the real value. Blank for other loads." },
  { key: "calcI", header: "Calculated Current (A) – specified PF", kind: "calc", width: 14, format: "#,##0", comment: "I = kW x 1000 / (sqrt(3) x V x PF), divided also by efficiency for DOL motors. Blank when kW, voltage or (DOL) efficiency is missing." },
  { key: "usedI", header: "Current used (A)", kind: "calc", width: 11, format: "#,##0", comment: "Entered current when valid, otherwise the calculated current." },
  { key: "src", header: "Current source", kind: "calc", width: 15 },
  { key: "design", header: "Design Current (A)", kind: "calc", width: 11, format: "#,##0", comment: "Design Current = Current used x multiplier (1.25 on the Notes sheet). A screening factor only, not a final protective-device rating." },
  { key: "protType", header: "Feeder Protection Type", kind: "select", width: 17, comment: "Preliminary. Design current up to 400 A: MCCB. 400 to 500 A: engineering review (or a 630 A MCCB if the Notes sheet says so). 500 A and above: EDO ACB." },
  { key: "protRating", header: "Indicative Feeder Breaker Rating (A)", kind: "select", width: 13, format: "#,##0", comment: "MCCB: 6, 10, 16, 25, 32, 50, 63, 100, 125, 160, 200, 250, 400 A (630 A if chosen). EDO ACB: 630, 1000, 1250, 1600, 2000, 2500, 3200, 4000 A (630 to 800 A is built as 1000 A). Blank when engineering review is required." },
  { key: "selected", header: "Selected Breaker (MCCB / EDO ACB)", kind: "select", width: 22, comment: "Preliminary selection. Engineering approval required before procurement." },
  { key: "relay", header: "Overload Relay Setting (A)", kind: "select", width: 12, format: "#,##0.0", comment: "Indicative: DOL motors only, at the motor current used. VFD overload protection is set in the drive. Confirm against the nameplate." },
  { key: "incType", header: "Panel Incomer Breaker Type (MCCB / ACB)", kind: "select", width: 18, comment: "Only for rows whose description contains 'incomer'. An existing 800 A ACB incomer is standardised to 1000 A EDO ACB. Not sized from the sum of feeder ratings: needs maximum demand, diversity, fault level, etc." },
  { key: "incRating", header: "Panel Incomer Breaker Rating (A)", kind: "select", width: 13, format: "#,##0" },
  { key: "ct", header: "CT Ratio (Primary/1 A)", kind: "select", width: 12, comment: "Indicative, incomer rows only: smallest standard primary not below the incomer rating, with a 1 A secondary. No CT is assigned to motor feeders. Confirm against metering/protection scheme, burden, accuracy class and saturation." },
  { key: "ctUse", header: "CT Use / Metering", kind: "select", width: 20 },
  { key: "cableI", header: "Cable Design Current (A)", kind: "calc", width: 12, format: "#,##0", comment: "Taken equal to the Design Current. Cable capacity, derating and voltage drop still to be checked." },
  { key: "cableSize", header: "Cable size / runs", kind: "input", width: 14, comment: "Taken from the source column if the load list has one; otherwise an input field left blank for the designer." },
  { key: "cableLen", header: "Cable length (m)", kind: "input", width: 11, format: "#,##0", comment: "Taken from the source column if the load list has one; otherwise an input field left blank." },
  { key: "notes", header: "Engineering Notes / Checks", kind: "calc", width: 60 },
];

export function buildMccWorkbook(
  rawSheet: readonly (readonly Cell[])[],
  sourceName: string,
  assumptions: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS,
): MccWorkbookResult {
  const raw = rawSheet.map((r) => [...r]);
  const parsed = parseLoadSheet(raw);
  if (parsed.errors.length) throw new Error(parsed.errors.join(" "));
  const headerIdx = raw.findIndex((r) => r.some((c) => !blank(c)));
  const header = raw[headerIdx];
  const width = Math.max(...raw.slice(headerIdx).map((r) => r.length));
  const descCol = header.findIndex((h) => norm(h).startsWith("desc"));
  const cableSizeCol = header.findIndex((h) => norm(h).includes("cable") && (norm(h).includes("size") || norm(h).includes("run")));
  const cableLenCol = header.findIndex((h) => norm(h).includes("length") || (norm(h).includes("cable") && norm(h).includes("len")));

  const wb = new ExcelJS.Workbook();
  // No author metadata.
  wb.creator = ""; wb.lastModifiedBy = "";
  wb.calcProperties.fullCalcOnLoad = true;
  const ws = wb.addWorksheet("Load list", { views: [{ state: "frozen", ySplit: 4, xSplit: 0 }] });
  const notes = wb.addWorksheet("Notes");

  // Column map: source columns first (kept as they are), then the added ones.
  const L: Record<string, string> = {};
  COLS.forEach((c, i) => { L[c.key] = letter(width + 1 + i); });
  const last = width + COLS.length;

  ws.getCell("A1").value = "MCC Load List / Breaker Selection (preliminary)";
  ws.getCell("A1").font = { bold: true, size: 14 };
  ws.getCell("A2").value = `Source: ${sourceName || "load list"}. Blue = input data, green = calculated, yellow = preliminary selection. Not a substitute for approved electrical design: see the Notes sheet.`;
  ws.getCell("A2").font = { italic: true };

  const HEAD = 4;
  const headRow = ws.getRow(HEAD);
  headRow.height = 62;
  const style = (cell: ExcelJS.Cell, fill: string, bold = false) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    cell.border = BORDER;
    cell.font = { bold };
    cell.alignment = { vertical: "middle", wrapText: true };
  };
  for (let c = 0; c < width; c++) {
    const cell = headRow.getCell(c + 1);
    cell.value = (header[c] as ExcelJS.CellValue) ?? "";
    style(cell, SOURCE, true);
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    ws.getColumn(c + 1).width = Math.min(45, Math.max(12, String(header[c] ?? "").length + 2));
  }
  COLS.forEach((c, i) => {
    const cell = headRow.getCell(width + 1 + i);
    cell.value = c.header;
    style(cell, c.kind === "input" ? SOURCE : c.kind === "calc" ? CALC : PRELIM, true);
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    if (c.comment) cell.note = c.comment;
    ws.getColumn(width + 1 + i).width = c.width;
  });

  let missing = 0;
  parsed.rows.forEach((row, n) => {
    const r = HEAD + 1 + n;
    const src = raw[row.rowNumber - 1];
    const fillSrc = n % 2 === 0 ? "FFFFFFFF" : "FFF7F7F7";
    for (let c = 0; c < width; c++) {
      const cell = ws.getCell(r, c + 1);
      const v = src[c];
      cell.value = v === undefined ? null : v instanceof Date ? v : (v as ExcelJS.CellValue);
      style(cell, fillSrc);
      if (typeof v === "number") cell.alignment = { horizontal: "center", vertical: "middle" };
    }

    const type = loadType(row.type);
    const res = currentFor({ type, kw: row.kw, voltageV: row.voltageV, ratedA: row.amps }, assumptions);
    if (res.usedA === undefined) missing++;
    const sel = selectBreaker(res.designA, { acbThresholdA: ACB_THRESHOLD_A, mccb630: assumptions.mccb630 });
    const isIncomer = descCol >= 0 && /incomer/i.test(String(src[descCol] ?? ""));
    const is800 = isIncomer && row.amps === 800;
    const incSel = isIncomer ? (is800 ? selectBreaker(res.designA, { existing800AcbIncomer: true, mccb630: assumptions.mccb630 }) : sel) : undefined;
    const ct = incSel?.ratingA !== undefined ? ctPrimaryFor(incSel.ratingA) : undefined;

    const f = (key: string) => `${L[key]}${r}`;
    const desc = descCol >= 0 ? `${letter(descCol + 1)}${r}` : '""';
    const calcVResult = res.usedA === undefined && !(row.voltageV && row.voltageV > 0) ? "" : row.voltageV !== undefined ? (Math.round(row.voltageV) === 433 ? assumptions.calcVoltage433 : row.voltageV) : "";
    const ctLadder = CT_PRIMARIES_A.reduceRight<string>((acc, p) => `IF(${f("incRating")}<=${p},"${p}/1 A",${acc})`, '""');
    const ratingLadder = [...MCCB_RATINGS_A, 630, 1000, 1250, 1600, 2000, 2500, 3200].reduceRight<string>((acc, p) => `IF(${f("design")}<=${p},${p},${acc})`, "4000");

    const noteParts =
      `IF(NOT(ISNUMBER(${f("usedI")})),"Missing input: give a rated current, or kW and voltage. ","")` +
      `&IF(AND(${f("type")}="DOL",ISNUMBER(${f("kw")}),NOT(ISNUMBER(${f("eff")}))),"Efficiency required for the DOL calculation. ","")` +
      `&IF(${f("src")}="Entered current","Rated current from the load list used. ","")` +
      `&IF(AND(${f("src")}="Calculated",${f("type")}="DOL"),"Motor efficiency is a provisional assumption: confirm. ","")` +
      `&IF(AND(${f("src")}="Calculated",${f("type")}="FEEDER"),"Feeder current calculated from kW at the specified PF: confirm. ","")` +
      `&IF(${f("type")}="VFD","Verify VFD input current and protection with the manufacturer's data. ","")` +
      `&IF(${f("protType")}="Engineering review required",IF(${f("design")}>4000,"Design current above the largest standard ACB (4000 A): engineering review required. ","Design current above 400 A but below the ${ACB_THRESHOLD_A} A ACB threshold: engineering review required (no undersized 400 A MCCB selected; or take 630 A as an MCCB on the Notes sheet). "),"")` +
      `&IF(AND(ISNUMBER(${f("protRating")}),${f("protRating")}>630),"Above 630 A: ACB rating needs engineering verification. ","")` +
      `&IF(AND(ISNUMBER(${f("rated")}),${f("rated")}=800,ISNUMBER(SEARCH("incomer",${desc}))),"Existing 800 A ACB incomer standardised to 1000 A EDO ACB; re-evaluate CT ratio, busbar and fault level. ","")` +
      `&IF(${f("incType")}<>"","Incomer breaker and CT are provisional: size from maximum demand, diversity, fault level and discrimination, not the sum of feeders. ","")` +
      `&"Preliminary: engineering approval required."`;

    const fm = (key: string, formula: string, result?: string | number | boolean) => {
      const col = COLS.findIndex((c) => c.key === key);
      const cell = ws.getCell(r, width + 1 + col);
      cell.value = { formula, result } as ExcelJS.CellValue;
      return cell;
    };
    const val = (key: string, v: ExcelJS.CellValue) => {
      const col = COLS.findIndex((c) => c.key === key);
      const cell = ws.getCell(r, width + 1 + col);
      cell.value = v;
      return cell;
    };
    // Inputs read (values), then formulas.
    val("type", type === "FEEDER" ? "Feeder" : type);
    val("volt", row.voltageV ?? null);
    val("kw", row.kw ?? null);
    val("rated", row.amps ?? null);
    fm("calcV", `IF(ISNUMBER(${f("volt")}),IF(ROUND(${f("volt")},0)=433,${A.calcV},${f("volt")}),"")`, calcVResult);
    fm("pf", `IF(${f("type")}="VFD",${A.pfVfd},IF(${f("type")}="DOL",${A.pfDol},${A.pfGeneral}))`, res.pf);
    fm("eff", `IF(AND(${f("type")}="DOL",ISNUMBER(${f("kw")})),${A.eff},"")`, type === "DOL" && row.kw !== undefined ? assumptions.efficiency : "");
    fm("calcI",
      `IF(AND(ISNUMBER(${f("kw")}),ISNUMBER(${f("calcV")}),ISNUMBER(${f("pf")})),IF(AND(${f("kw")}>0,${f("calcV")}>0),` +
      `IF(${f("type")}="DOL",IF(ISNUMBER(${f("eff")}),${f("kw")}*1000/(SQRT(3)*${f("calcV")}*${f("pf")}*${f("eff")}),""),` +
      `${f("kw")}*1000/(SQRT(3)*${f("calcV")}*${f("pf")})),""),"")`, res.calculatedA ?? "");
    fm("usedI", `IF(AND(ISNUMBER(${f("rated")}),${f("rated")}>0),${f("rated")},${f("calcI")})`, res.usedA ?? "");
    fm("src", `IF(AND(ISNUMBER(${f("rated")}),${f("rated")}>0),"Entered current",IF(ISNUMBER(${f("calcI")}),"Calculated","Missing input"))`, res.source);
    fm("design", `IF(ISNUMBER(${f("usedI")}),${f("usedI")}*${A.mult},"")`, res.designA ?? "");
    fm("protType",
      `IF(NOT(ISNUMBER(${f("design")})),"",IF(${f("design")}<=400,"MCCB",IF(AND(${A.mccb630}=1,${f("design")}<=630),"MCCB",IF(${f("design")}<${A.threshold},"Engineering review required",IF(${f("design")}<=4000,"EDO ACB","Engineering review required")))))`,
      res.designA === undefined ? "" : sel.kind === "REVIEW" ? "Engineering review required" : sel.kind);
    fm("protRating", `IF(OR(${f("protType")}="",${f("protType")}="Engineering review required"),"",${ratingLadder})`, sel.ratingA ?? "");
    fm("selected", `IF(${f("protType")}="","",IF(${f("protRating")}="","Engineering review required",${f("protRating")}&" A "&${f("protType")}))`, res.designA === undefined ? "" : breakerLabel(sel));
    fm("relay", `IF(AND(${f("type")}="DOL",ISNUMBER(${f("usedI")})),ROUND(${f("usedI")},1),"")`, type === "DOL" && res.usedA !== undefined ? Math.round(res.usedA * 10) / 10 : "");
    fm("incType",
      `IF(ISNUMBER(SEARCH("incomer",${desc})),IF(AND(ISNUMBER(${f("rated")}),${f("rated")}=800),"EDO ACB",${f("protType")}),"")`,
      incSel ? (incSel.kind === "REVIEW" ? "Engineering review required" : incSel.kind) : "");
    fm("incRating",
      `IF(${f("incType")}="","",IF(${f("incType")}="Engineering review required","",IF(AND(ISNUMBER(${f("rated")}),${f("rated")}=800),1000,${f("protRating")})))`,
      incSel?.ratingA ?? "");
    fm("ct", `IF(NOT(ISNUMBER(${f("incRating")})),"",${ctLadder})`, ct !== undefined ? `${ct}/1 A` : "");
    fm("ctUse", `IF(${f("ct")}="","","Metering / protection: confirm scheme")`, ct !== undefined ? "Metering / protection: confirm scheme" : "");
    fm("cableI", `IF(ISNUMBER(${f("design")}),${f("design")},"")`, res.designA ?? "");
    if (cableSizeCol >= 0) fm("cableSize", `IF(${letter(cableSizeCol + 1)}${r}="","",${letter(cableSizeCol + 1)}${r})`);
    else val("cableSize", null);
    if (cableLenCol >= 0) fm("cableLen", `IF(${letter(cableLenCol + 1)}${r}="","",${letter(cableLenCol + 1)}${r})`);
    else val("cableLen", null);
    fm("notes", noteParts);

    COLS.forEach((c, i) => {
      const cell = ws.getCell(r, width + 1 + i);
      const base = c.kind === "input" ? SOURCE : c.kind === "calc" ? CALC : PRELIM;
      style(cell, n % 2 === 0 ? lighter(base) : base);
      if (c.format) cell.numFmt = c.format;
      cell.alignment = { vertical: "middle", horizontal: c.key === "notes" || c.key === "selected" || c.key === "src" || c.key === "ctUse" ? "left" : "center", wrapText: c.key === "notes" };
    });

    for (const key of ["protType", "incType"]) {
      ws.getCell(r, width + 1 + COLS.findIndex((c) => c.key === key)).dataValidation = {
        type: "list", allowBlank: true, showErrorMessage: true, errorStyle: "warning",
        errorTitle: "Breaker type", error: "Use MCCB, EDO ACB or Engineering review required.",
        formulae: ['"MCCB,EDO ACB,Engineering review required"'],
      };
    }
  });

  ws.autoFilter = { from: { row: HEAD, column: 1 }, to: { row: Math.max(HEAD, HEAD + parsed.rows.length), column: last } };

  buildNotes(notes, assumptions);
  return { workbook: wb, rows: parsed.rows.length, missing };
}

function buildNotes(ws: ExcelJS.Worksheet, a: MccAssumptions) {
  ws.getColumn(1).width = 52;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 90;
  ws.getCell("A1").value = "Notes: MCC Load List / Breaker Selection";
  ws.getCell("A1").font = { bold: true, size: 14 };
  ws.getCell("A2").value = "Preliminary engineering aid. Edit the yellow assumption cells; the Load list formulas follow them.";
  ws.getCell("A2").font = { italic: true };
  const h = ws.getRow(3);
  ["Assumption (editable)", "Value", "Used for"].forEach((t, i) => {
    const c = h.getCell(i + 1);
    c.value = t; c.font = { bold: true }; c.border = BORDER;
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ORANGE } };
  });
  const rows: [string, number, string, string][] = [
    ["Power factor, general feeder", a.pfGeneral, "Feeders and any load that is not DOL or VFD.", "0.00"],
    ["Power factor, DOL feeder / motor", a.pfDol, "DOL motors.", "0.00"],
    ["Power factor, VFD feeder", a.pfVfd, "VFD feeders. Actual VFD input current must be verified with the manufacturer's data.", "0.00"],
    ["Motor efficiency (provisional)", a.efficiency, "DOL motors given in kW (kW taken as motor output). Provisional 93%: not a confirmed value. Overwrite per row in the Load list if known.", "0%"],
    ["Design current multiplier", a.designMultiplier, "Design Current = Current used x multiplier. Screening factor only.", "0.00"],
    ["Calculation voltage for the 433 V level (V)", a.calcVoltage433, "Line voltage used in the current formula for 433 V loads (other voltages are used as given).", "#,##0"],
    ["ACB threshold (A)", ACB_THRESHOLD_A, "Design current at or above this: 630 A EDO ACB (or larger ACB). Between 400 A and this value: engineering review.", "#,##0"],
    ["630 A as MCCB instead of ACB (1 = yes, 0 = no)", a.mccb630 ? 1 : 0, "1: a requirement between 400 A and 630 A is a 630 A MCCB (customer preference). 0: it is an EDO ACB from 500 A, with engineering review between 400 and 500 A.", "0"],
  ];
  rows.forEach(([name, value, use, fmt], i) => {
    const r = 4 + i;
    ws.getCell(r, 1).value = name;
    const v = ws.getCell(r, 2);
    v.value = value; v.numFmt = fmt; v.alignment = { horizontal: "center" };
    v.fill = { type: "pattern", pattern: "solid", fgColor: { argb: PRELIM } };
    ws.getCell(r, 3).value = use;
    for (let c = 1; c <= 3; c++) { ws.getCell(r, c).border = BORDER; ws.getCell(r, c).alignment = { ...ws.getCell(r, c).alignment, wrapText: true, vertical: "middle" }; }
  });

  const text: [string, string[]][] = [
    ["Colour key", [
      "Blue: input data (source columns and values read from them). Green: calculated fields. Yellow: preliminary selections.",
    ]],
    ["Formulas", [
      "Current from kW: I (A) = kW x 1000 / (sqrt(3) x line voltage (V) x PF). For DOL motors also divided by the efficiency (kW taken as motor output). The efficiency is not applied to VFD or general feeders.",
      "If a valid rated current is in the load list it is used in preference to a calculation. Missing or invalid inputs leave the calculated cells blank and the Engineering Notes say what is missing: no zero is shown.",
      "Design Current = Current used x multiplier (1.25 above). Cable Design Current is taken equal to the Design Current.",
      "Overload relay setting (indicative): DOL motors only, at the current used. VFD overload protection is set in the drive.",
    ]],
    ["Power factors", ["General feeder 0.7, DOL 0.7, VFD 0.9. For VFDs the specified PF of 0.9 is used as requested, but actual VFD input current should be checked against the manufacturer's data."]],
    ["Breaker ratings", [
      "MCCB / MPCB / MCB ratings (the BOM list plus 32 and 63 A): 6, 10, 16, 25, 32, 50, 63, 100, 125, 160, 200, 250 and 400 A. A 630 A feeder may be an MCCB or an EDO ACB as the customer prefers (setting above). The device class per rating is to be defined later.",
      "Standard ACB ratings: 630, 800, 1000, 1250, 1600, 2000, 2500, 3200 and 4000 A. A 630 to 800 A requirement is built as a 1000 A ACB; the 800 A sheets are kept for later use.",
      "Design current up to 400 A: next MCCB rating in the list. 500 A and above: 630 A EDO ACB up to 630 A, then the next standard ACB (1000 / 1250 / 1600 / 2000 / 2500 / 3200 / 4000 A) flagged for engineering verification. Above 4000 A: engineering review.",
      "Between 400 A and 500 A: no breaker is selected. The row says 'Engineering review required' because MCCBs stop at 400 A and the ACB starts at 500 A; an undersized 400 A MCCB is never chosen silently.",
      "A load that exceeds the selected breaker is never capped silently; it is flagged.",
    ]],
    ["Incomer and CT", [
      "Incomer columns are filled only for rows whose description contains 'incomer'. An existing 800 A ACB incomer in the source is standardised to a 1000 A EDO ACB (this applies only where the source identifies it; other 800 A values are left alone).",
      "The incomer is never sized from the sum of the feeder ratings. It needs maximum demand, diversity, future expansion, busbar rating, fault level, breaking capacity and discrimination/coordination. Without that data the incomer is provisional.",
      "CT ratios are indicative: the smallest standard primary (100, 150, 200, 250, 300, 400, 500, 600, 750, 800, 1000, 1250, 1600, 2000, 2500, 3200, 4000 A) not below the incomer rating, with a 1 A secondary. No CT is assigned to motor feeders. An 800/1 A ratio is not kept when the incomer is standardised to 1000 A. Confirm against maximum operating current, meter/relay requirements, accuracy class, burden, saturation and the protection scheme.",
    ]],
    ["Cable columns", ["Cable size / runs and cable length are taken from the source columns when the load list has them; otherwise they are blank input fields for the designer."]],
    ["Final engineering checks (not done here)", [
      "Maximum demand and diversity; cable current capacity and derating; voltage drop; short-circuit withstand; prospective fault level and breaking capacity; motor starting; earth-fault protection; discrimination/coordination; busbar rating; applicable standards.",
    ]],
    ["Disclaimer", ["This workbook is a preliminary engineering aid and not a substitute for approved electrical design. All selections require engineering approval before procurement."]],
  ];
  let r = 4 + rows.length + 1;
  for (const [title, lines] of text) {
    const t = ws.getCell(r, 1);
    t.value = title; t.font = { bold: true }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ORANGE } };
    ws.getCell(r, 2).fill = t.fill; ws.getCell(r, 3).fill = t.fill;
    r++;
    for (const line of lines) {
      ws.mergeCells(r, 1, r, 3);
      const c = ws.getCell(r, 1);
      c.value = line; c.alignment = { wrapText: true, vertical: "top" };
      ws.getRow(r).height = Math.max(18, Math.ceil(line.length / 140) * 16);
      r++;
    }
    r++;
  }
}

/** ExcelJS writes "Unknown" as author when none is set; remove the author fields entirely. */
function stripAuthor(bytes: Uint8Array): Uint8Array {
  const parts = unzipSync(bytes);
  const core = parts["docProps/core.xml"];
  if (core) {
    parts["docProps/core.xml"] = strToU8(
      strFromU8(core).replace(/<dc:creator>.*?<\/dc:creator>/, "").replace(/<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/, ""),
    );
  }
  return zipSync(parts);
}

export async function mccWorkbookBuffer(
  rawSheet: readonly (readonly Cell[])[],
  sourceName: string,
  assumptions?: MccAssumptions,
): Promise<{ buffer: ArrayBuffer; rows: number; missing: number }> {
  const { workbook, rows, missing } = buildMccWorkbook(rawSheet, sourceName, assumptions);
  const out = new Uint8Array(await workbook.xlsx.writeBuffer());
  const clean = stripAuthor(out);
  const buffer = clean.buffer.slice(clean.byteOffset, clean.byteOffset + clean.byteLength) as ArrayBuffer;
  return { buffer, rows, missing };
}
