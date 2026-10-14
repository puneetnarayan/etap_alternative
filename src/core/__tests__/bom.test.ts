import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import ExcelJS from "exceljs";
import { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import { DEFAULT_MCC_ASSUMPTIONS, bomItems, mccRows } from "..";
import { buildBomWorkbook } from "../bom-workbook";

const buf = (name: string) => {
  const b = readFileSync(new URL(name, import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};
const sample = () => readSheet(buf("../../../public/sample_load_list.xlsx"));
const template = () => buf("../../../public/bom/unit_bom_template.xlsx");

const sheetOf = (rows: unknown[][]) => [["Description", "Type", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)"], ...rows];

describe("unique BOM items (type + rating)", () => {
  it("loads with the same type, breaker rating and kW are one item, counted", () => {
    const items = bomItems(mccRows(sheetOf([
      ["Fan A", "VFD", null, 75, 433], ["Fan B", "VFD", null, 75, 433], ["Fan C", "VFD", null, 110, 433],
    ])).rows);
    expect(items.map((i) => [i.type, i.breaker, i.kw, i.count])).toEqual([["VFD", "250 A MCCB", 110, 1], ["VFD", "160 A MCCB", 75, 2]]);
    expect(items[1].rows).toEqual([2, 3]);
  });

  it("sheet matching: incomer, feeders, VFD, DOL; plain 6 / 10 A feeders need no sheet", () => {
    const items = bomItems(mccRows(sheetOf([
      ["MCC-2 incomer", "Feeder", 800, null, 433], // design 1000 A -> 1000 A ACB
      ["Panel", "Feeder", 200, null, 433], // 250 A -> 250 A MCCB
      ["Small", "Feeder", 25, null, 433], // 31 A -> 32 A MCCB
      ["Tiny", "Feeder", 4, null, 433], // 5 A -> 6 A: no sheet
      ["Pump", "DOL", null, 30, 433], // 30 kW DOL
      ["Big", "Feeder", 2000, null, 433], // 2500 A -> above 1250 A: no outgoing sheet
    ])).rows);
    const by = (b: string) => items.find((i) => i.breaker === b)!;
    expect(by("1000 A EDO ACB")).toMatchObject({ type: "Incomer", sheets: ["Incomer 1000A IMCC "] });
    expect(by("250 A MCCB")).toMatchObject({ type: "Feeder", sheets: ["OG 250 to 630A drawout Fdr MCCB"] });
    expect(by("32 A MCCB")).toMatchObject({ type: "Feeder", sheets: ["OG <250A  power FDR - MCCB"] });
    expect(by("6 A MCCB")).toMatchObject({ noSheet: true, sheets: [] });
    expect(by("6 A MCCB").action).toBeUndefined(); // not required: no action
    expect(by("2500 A EDO ACB").action).toMatch(/No outgoing ACB sheet above 1250 A/);
    expect(items.find((i) => i.type === "DOL")).toMatchObject({ sheets: ["DOL < 75KW"], kw: 30 });
  });

  it("soft-starter ACB feeders keep both sheets (225 kW and 315 kW) with an action to review", () => {
    const items = bomItems(mccRows(sheetOf([["Mill", "DOL", null, 315, 433]])).rows);
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ type: "DOL + soft starter", breaker: "1000 A EDO ACB" });
    expect(items[0].sheets).toEqual(["OG 315KW Soft starter ", "OG 225KW Soft starter"]);
    expect(items[0].action).toMatch(/225 kW or 315 kW/);
  });

  it("a load without a breaker (engineering review) is an item with an action, not a guess", () => {
    const items = bomItems(mccRows(sheetOf([["Pump", "DOL", null, 185, 433]])).rows); // 474 A design: 400 to 500 A band
    expect(items[0]).toMatchObject({ breaker: "", sheets: [] });
    expect(items[0].action).toMatch(/engineering review/);
    const asMccb = bomItems(mccRows(sheetOf([["Pump", "DOL", null, 185, 433]]), { ...DEFAULT_MCC_ASSUMPTIONS, mccb630: true }).rows);
    expect(asMccb[0]).toMatchObject({ breaker: "630 A MCCB", sheets: ["DOL < 75KW"] });
  });

  it("a row with no current data gives no item", () => {
    expect(bomItems(mccRows(sheetOf([["Mystery", "DOL", null, null, 433]])).rows)).toEqual([]);
  });
});

describe("the BOM workbook", () => {
  it("has an Index plus one sheet per matched item, built from the template layout", async () => {
    const items = bomItems(mccRows(await sample()).rows);
    const r = await buildBomWorkbook(template(), items, "sample_load_list.xlsx");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(r.buffer);
    const matched = items.flatMap((i) => i.sheets).length;
    expect(wb.worksheets).toHaveLength(1 + matched);
    expect(wb.worksheets[0].name).toBe("Index");
    expect(r.sheets).toBe(items.filter((i) => i.sheets.length).length);
    // index: one row per item; the Class column is there and blank
    const idx = wb.getWorksheet("Index")!;
    const head = (idx.getRow(5).values as unknown[]).map(String);
    expect(head).toContain("Class");
    expect(idx.rowCount).toBe(5 + items.length);
    for (let i = 0; i < items.length; i++) expect(idx.getRow(6 + i).getCell(head.indexOf("Class")).value ?? null).toBeNull();
    // action items are red
    const actionCol = head.indexOf("Action");
    const redRows = items.map((it, i) => ({ it, cell: idx.getRow(6 + i).getCell(actionCol) })).filter((x) => x.it.action);
    expect(redRows.length).toBeGreaterThan(0);
    for (const x of redRows) expect(x.cell.fill).toMatchObject({ fgColor: { argb: "FFFF0000" } });
  });

  it("a cloned sheet keeps the template's layout; only the rating text changes (VFD 160 A / 75 kW)", async () => {
    const items = bomItems(mccRows(sheetOf([["Fan", "VFD", null, 75, 433]])).rows);
    const r = await buildBomWorkbook(template(), items, "t");
    const out = new ExcelJS.Workbook(); await out.xlsx.load(r.buffer);
    const tpl = new ExcelJS.Workbook(); await tpl.xlsx.load(template());
    const a = tpl.getWorksheet("55kW VFD")!;
    const b = out.worksheets[1];
    for (let c = 1; c <= 8; c++) expect(b.getColumn(c).width).toBe(a.getColumn(c).width);
    let changed = 0;
    a.eachRow({ includeEmpty: true }, (row, rn) => {
      for (let c = 1; c <= 14; c++) {
        const x = row.getCell(c); const y = b.getRow(rn).getCell(c);
        expect(y.font?.bold, `${rn},${c}`).toBe(x.font?.bold);
        expect(JSON.stringify(y.fill ?? null), `${rn},${c}`).toBe(JSON.stringify(x.fill ?? null));
        if (String(x.value) !== String(y.value)) { changed++; }
      }
    });
    expect(changed).toBe(2); // the "< 200A 4P" line and the sheet title
    const text = b.getCell("A4").value;
    expect(String(text)).toBe("MCCB - Drawout type,Double break mechanism");
    expect(String(b.getCell("B4").value)).toMatch(/^160A 4P/);
    expect(String(b.getCell("B3").value)).toBe("OG 160A VFD feeder (75 kW)");
  });

  it("the template has the CT secondary at 1 A everywhere and the new sheets", async () => {
    const tpl = new ExcelJS.Workbook(); await tpl.xlsx.load(template());
    const names = tpl.worksheets.map((w) => w.name);
    for (const n of ["OG 630A ACB Feeder", "Incomer 1600A", "Incomer 2000A", "Incomer 2500A", "Buscoupler 2500A", "OG 630A ACB", "OG 800A ACB", "Incomer 3200A"]) expect(names).toContain(n);
    let ct5 = 0; let ct1 = 0;
    for (const ws of tpl.worksheets) ws.eachRow((row) => row.eachCell((c) => {
      const v = String(c.value ?? "");
      if (/CT CTR:\s*\d+\s*\/\s*5\s*A/i.test(v)) ct5++;
      if (/CT CTR:\s*\d+\s*\/\s*1\s*A/i.test(v)) ct1++;
    }));
    expect(ct5).toBe(0);
    expect(ct1).toBeGreaterThan(20);
    const inc = tpl.getWorksheet("Incomer 4000A")!;
    expect(inc.getCell("C24").value).toMatch(/4000\/1A/);
  });

  it("has no author metadata", async () => {
    const r = await buildBomWorkbook(template(), bomItems(mccRows(await sample()).rows), "s");
    const core = strFromU8(unzipSync(new Uint8Array(r.buffer))["docProps/core.xml"]);
    expect(core).not.toMatch(/creator|lastModifiedBy|Unknown/i);
  });
});
