/**
 * Builds the BOM workbook: an Index sheet plus one sheet per unique item, each cloned from the unit-BOM template
 * with its layout, fonts, widths and colours unchanged and only the rating text set (docs/bom.md).
 */
import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { ratingReplacements, type BomItem } from "./bom";

const ORANGE = "FFFFE0B2";
const RED = "FFFF0000";
const THIN = { style: "thin" as const, color: { argb: "FFBFBFBF" } };
const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const MAX_COL = 20;

export interface BomBuildResult {
  /** Items that got a sheet of their own. */
  sheets: number;
  /** Items with an action (red). */
  actions: number;
  /** Name of the sheet built for each item key (items without a sheet are absent). */
  sheetNames: Record<string, string>;
}

/** Excel sheet names: at most 31 characters, none of \ / ? * [ ] : */
const safe = (n: string) => n.replace(/[\\/?*[\]:]/g, "-").slice(0, 31).trim();

function copySheet(src: ExcelJS.Worksheet, dst: ExcelJS.Worksheet, repl: [RegExp, string][]) {
  for (let c = 1; c <= MAX_COL; c++) {
    const w = src.getColumn(c).width;
    if (w) dst.getColumn(c).width = w;
  }
  src.eachRow({ includeEmpty: true }, (row, r) => {
    const out = dst.getRow(r);
    if (row.height) out.height = row.height;
    for (let c = 1; c <= MAX_COL; c++) {
      const s = row.getCell(c);
      const d = out.getCell(c);
      let v = s.value;
      if (typeof v === "string") for (const [re, to] of repl) v = v.replace(re, to);
      else if (v && typeof v === "object" && "hyperlink" in v) v = { text: "BACK TO INDEX", hyperlink: "#'Index'!A1" } as ExcelJS.CellValue;
      d.value = v;
      d.style = JSON.parse(JSON.stringify(s.style));
    }
  });
  for (const m of src.model.merges ?? []) dst.mergeCells(m);
  dst.properties.tabColor = src.properties.tabColor;
  dst.views = src.views.map((v) => ({ ...v }));
}

/** ExcelJS writes "Unknown" as author when none is set: remove the author fields. */
function stripAuthor(bytes: Uint8Array): Uint8Array {
  const parts = unzipSync(bytes);
  const core = parts["docProps/core.xml"];
  if (core) {
    parts["docProps/core.xml"] = strToU8(strFromU8(core).replace(/<dc:creator>.*?<\/dc:creator>/, "").replace(/<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/, ""));
  }
  return zipSync(parts);
}

export async function buildBomWorkbook(
  template: ArrayBuffer | Uint8Array,
  items: BomItem[],
  sourceName: string,
): Promise<{ buffer: ArrayBuffer } & BomBuildResult> {
  const tpl = new ExcelJS.Workbook();
  await tpl.xlsx.load(template as ArrayBuffer);
  const out = new ExcelJS.Workbook();
  out.creator = ""; out.lastModifiedBy = "";
  const index = out.addWorksheet("Index", { views: [{ state: "frozen", ySplit: 5 }] });

  const sheetNames: Record<string, string> = {};
  const used = new Set<string>(["Index"]);
  let actions = 0;
  let n = 0;

  const itemRows: { item: BomItem; names: string[] }[] = [];
  for (const item of items) {
    const names: string[] = [];
    item.sheets.forEach((tplName) => {
      const src = tpl.getWorksheet(tplName);
      if (!src) return;
      n++;
      const base = safe(`${String(n).padStart(2, "0")} ${tplName.trim()}`);
      let name = base;
      for (let k = 2; used.has(name); k++) name = safe(`${base.slice(0, 28)} ${k}`);
      used.add(name);
      const dst = out.addWorksheet(name);
      copySheet(src, dst, ratingReplacements(item, tplName));
      names.push(name);
    });
    if (names.length) sheetNames[item.key] = names.join(" / ");
    if (item.action) actions++;
    itemRows.push({ item, names });
  }

  // Index
  index.getCell("A1").value = "BOM index: one unit BOM per unique item (type + rating)";
  index.getCell("A1").font = { bold: true, size: 14 };
  index.getCell("A2").value = `Source load list: ${sourceName || "load list"}. Quantities come from the load list. Sheet layouts are those of the unit-BOM template. Preliminary: engineering approval required.`;
  index.getCell("A3").value = "Class column: logic to be given later (reminder). Red = action item, to be reviewed.";
  index.getCell("A3").font = { italic: true };
  const heads = ["SL no", "Type", "Breaker / rating", "Class", "Motor kW", "Quantity", "Loads (sheet rows)", "Sheet(s) in this workbook", "Action"];
  const widths = [7, 20, 20, 12, 10, 10, 22, 50, 70];
  heads.forEach((h, i) => {
    const c = index.getCell(5, i + 1);
    c.value = h; c.font = { bold: true }; c.border = BORDER;
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ORANGE } };
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    index.getColumn(i + 1).width = widths[i];
  });
  index.getCell("D5").note = "Class: the logic will be given later.";
  itemRows.forEach(({ item, names }, i) => {
    const r = 6 + i;
    const vals: (string | number | null)[] = [
      i + 1, item.type, item.breaker || "(review)", null, item.kw ?? null, item.count, item.rows.join(", "),
      item.noSheet ? "No BOM sheet required" : names.join("; ") || "(none)", item.action ?? "",
    ];
    vals.forEach((v, j) => {
      const c = index.getCell(r, j + 1);
      c.value = v; c.border = BORDER;
      c.alignment = { vertical: "top", wrapText: true, horizontal: j === 0 || j === 4 || j === 5 ? "center" : "left" };
    });
    // each item's sheet name links to its sheet
    if (names.length === 1) index.getCell(r, 8).value = { text: names[0], hyperlink: `#'${names[0]}'!A1` } as ExcelJS.CellValue;
    if (item.action) {
      const a = index.getCell(r, 9);
      a.fill = { type: "pattern", pattern: "solid", fgColor: { argb: RED } };
      a.font = { bold: true, color: { argb: "FFFFFFFF" } };
    }
  });
  index.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5 + Math.max(1, itemRows.length), column: heads.length } };

  const bytes = new Uint8Array(await out.xlsx.writeBuffer());
  const clean = stripAuthor(bytes);
  return {
    buffer: clean.buffer.slice(clean.byteOffset, clean.byteOffset + clean.byteLength) as ArrayBuffer,
    sheets: Object.keys(sheetNames).length, actions, sheetNames,
  };
}
