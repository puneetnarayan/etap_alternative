import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import ExcelJS from "exceljs";
import { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import {
  breakerLabel, ctPrimaryFor, currentFor, mccRows, selectBreaker,
} from "..";
import { mccWorkbookBuffer } from "../mcc-workbook";

const label = (d: number | undefined, o: Parameters<typeof selectBreaker>[1] = {}) => breakerLabel(selectBreaker(d, o));

describe("breaker selection", () => {
  it("MCCB: next rating in the BOM list 6, 10, 16, 25, 32, 50, 63, 100, 125, 160, 200, 250, 400 A", () => {
    expect(label(20)).toBe("25 A MCCB");
    expect(label(50)).toBe("50 A MCCB");
    expect(label(120)).toBe("125 A MCCB");
    expect(label(300)).toBe("400 A MCCB");
    expect(label(5)).toBe("6 A MCCB");
    expect(label(32)).toBe("32 A MCCB"); // 32 and 63 A are kept
    expect(label(60)).toBe("63 A MCCB");
    expect(label(400)).toBe("400 A MCCB");
  });

  it("450 A -> engineering review, not an undersized 400 A MCCB (unless 630 A is taken as an MCCB)", () => {
    const s = selectBreaker(450);
    expect(s).toMatchObject({ kind: "REVIEW", status: "review" });
    expect(s.ratingA).toBeUndefined();
    expect(s.message).toMatch(/engineering review/i);
    expect(selectBreaker(400.5).kind).toBe("REVIEW");
    expect(selectBreaker(499.9).kind).toBe("REVIEW");
    expect(label(450, { mccb630: true })).toBe("630 A MCCB");
    expect(label(630, { mccb630: true })).toBe("630 A MCCB");
    expect(label(631, { mccb630: true })).toBe("1000 A EDO ACB");
  });

  it("500 A and 620 A -> 630 A EDO ACB (the default for a 630 A feeder)", () => {
    expect(label(500)).toBe("630 A EDO ACB");
    expect(label(620)).toBe("630 A EDO ACB");
    expect(label(630)).toBe("630 A EDO ACB");
    for (let d = 1; d <= 1300; d += 7) {
      const s = selectBreaker(d);
      if (s.kind === "MCCB") expect([6, 10, 16, 25, 32, 50, 63, 100, 125, 160, 200, 250, 400]).toContain(s.ratingA);
    }
  });

  it("an existing 800 A ACB incomer -> 1000 A EDO ACB; 630 to 800 A is built as 1000 A", () => {
    expect(label(1000, { existing800AcbIncomer: true })).toBe("1000 A EDO ACB");
    expect(selectBreaker(undefined, { existing800AcbIncomer: true }).ratingA).toBe(1000);
    expect(selectBreaker(700)).toMatchObject({ kind: "EDO ACB", ratingA: 1000, status: "verify" });
    expect(selectBreaker(800).ratingA).toBe(1000);
  });

  it("above 630 A -> next standard ACB (1000, 1250, 1600, 2000, 2500, 3200, 4000), flagged for verification; above 4000 A -> review", () => {
    for (const [d, r] of [[631, 1000], [1000, 1000], [1001, 1250], [1250, 1250], [1251, 1600], [1700, 2000], [2001, 2500], [2600, 3200], [3300, 4000], [4000, 4000]] as const) {
      expect(selectBreaker(d), String(d)).toMatchObject({ kind: "EDO ACB", ratingA: r, status: "verify" });
    }
    expect(selectBreaker(4001).kind).toBe("REVIEW");
  });

  it("no design current -> no breaker, never a made-up value", () => {
    for (const d of [undefined, null, NaN, 0, -5]) expect(selectBreaker(d as number | undefined).kind).toBe("REVIEW");
  });

  it("indicative CT primary: smallest standard primary not below the rating (1000 A -> 1000, 4000 A -> 4000)", () => {
    expect(ctPrimaryFor(1000)).toBe(1000);
    expect(ctPrimaryFor(630)).toBe(750);
    expect(ctPrimaryFor(3200)).toBe(3200);
    expect(ctPrimaryFor(4000)).toBe(4000);
    expect(ctPrimaryFor(undefined)).toBeUndefined();
  });
});

describe("current and power factor", () => {
  it("PF: general feeder 0.7, DOL 0.7, VFD 0.9", () => {
    expect(currentFor({ type: "FEEDER", kw: 100, voltageV: 415 }).pf).toBe(0.7);
    expect(currentFor({ type: "DOL", kw: 100, voltageV: 415 }).pf).toBe(0.7);
    expect(currentFor({ type: "VFD", kw: 100, voltageV: 415 }).pf).toBe(0.9);
  });

  it("hand-calculated: 185 kW DOL at 433 V, 0.7, 93 % = 378.9 A; VFD 250 kW = 370.4 A; feeder 100 kW at 690 V = 119.5 A", () => {
    expect(currentFor({ type: "DOL", kw: 185, voltageV: 433 }).calculatedA).toBeCloseTo(378.92, 1); // 185000 / (1.7321 x 433 x 0.7 x 0.93)
    expect(currentFor({ type: "VFD", kw: 250, voltageV: 433 }).calculatedA).toBeCloseTo(370.38, 1); // no efficiency for VFD
    expect(currentFor({ type: "FEEDER", kw: 100, voltageV: 690 }).calculatedA).toBeCloseTo(119.5, 1);
    expect(currentFor({ type: "DOL", kw: 185, voltageV: 433 }).designA).toBeCloseTo(378.92 * 1.25, 1);
  });

  it("a valid rated current is preferred to a calculation", () => {
    const r = currentFor({ type: "DOL", kw: 185, voltageV: 433, ratedA: 300 });
    expect(r).toMatchObject({ usedA: 300, source: "Entered current" });
    expect(r.designA).toBeCloseTo(375, 9);
  });

  it("missing voltage or kW leaves everything blank (no zero)", () => {
    for (const i of [{ type: "DOL" as const, kw: 100 }, { type: "VFD" as const, voltageV: 433 }, { type: "FEEDER" as const }]) {
      const r = currentFor(i);
      expect(r.calculatedA).toBeUndefined();
      expect(r.usedA).toBeUndefined();
      expect(r.designA).toBeUndefined();
      expect(r.source).toBe("Missing input");
    }
  });
});

describe("the workbook", () => {
  const sample = async () => {
    const b = readFileSync(new URL("../../../public/sample_load_list.xlsx", import.meta.url));
    return readSheet(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer);
  };

  it("is a real xlsx with a Load list and a Notes sheet, formulas, a frozen header, filters, drop-downs and notes", async () => {
    const raw = await sample();
    const { buffer, rows } = await mccWorkbookBuffer(raw, "sample_load_list.xlsx");
    expect(rows).toBe(29);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    expect(wb.worksheets.map((w) => w.name)).toEqual(["Load list", "Notes"]);
    const ws = wb.getWorksheet("Load list")!;
    expect(ws.views[0]).toMatchObject({ state: "frozen", ySplit: 4 });
    expect(ws.autoFilter).toBeTruthy();
    const head = ws.getRow(4).values as unknown[];
    const names = head.map((h) => String(h ?? ""));
    for (const h of [
      "Estimated / entered load current (A)", "Calculated Current (A) – specified PF", "Design Current (A)", "Feeder Protection Type",
      "Indicative Feeder Breaker Rating (A)", "Selected Breaker (MCCB / EDO ACB)", "Overload Relay Setting (A)",
      "Panel Incomer Breaker Type (MCCB / ACB)", "Panel Incomer Breaker Rating (A)", "CT Ratio (Primary/1 A)", "CT Use / Metering",
      "Cable Design Current (A)", "Cable size / runs", "Cable length (m)", "Engineering Notes / Checks",
    ]) expect(names, h).toContain(h);
    // source columns kept intact, in order
    expect(names.slice(1, 8)).toEqual(raw[0].map(String));
    // calculated cells are formulas, not pasted numbers
    const col = (h: string) => names.indexOf(h);
    const calc = ws.getRow(5).getCell(col("Calculated Current (A) – specified PF")).value as { formula?: string };
    expect(calc.formula).toMatch(/SQRT\(3\)/);
    const design = ws.getRow(5).getCell(col("Design Current (A)")).value as { formula?: string };
    expect(design.formula).toContain("Notes!$B$8");
    expect(ws.getRow(5).getCell(col("Feeder Protection Type")).dataValidation?.type).toBe("list");
    expect(ws.getRow(4).getCell(col("Design Current (A)")).note).toBeTruthy();
    // original source values untouched
    expect(ws.getRow(5).getCell(1).value).toBe(raw[1][0]);
    const notes = wb.getWorksheet("Notes")!;
    expect(notes.getCell("B4").value).toBe(0.7);
    expect(notes.getCell("B6").value).toBe(0.9);
    expect(notes.getCell("B7").value).toBe(0.93);
    expect(notes.getCell("B8").value).toBe(1.25);
  });

  it("has no author metadata", async () => {
    const { buffer } = await mccWorkbookBuffer(await sample(), "s.xlsx");
    const core = strFromU8(unzipSync(new Uint8Array(buffer))["docProps/core.xml"]);
    expect(core).not.toMatch(/creator|lastModifiedBy|Unknown/i);
  });

  it("an incomer row at 800 A gets the 1000 A EDO ACB and a 1000/1 A CT; motor feeders get no CT or incomer", async () => {
    const raw = await sample();
    const { buffer } = await mccWorkbookBuffer(raw, "s.xlsx");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet("Load list")!;
    const names = (ws.getRow(4).values as unknown[]).map((h) => String(h ?? ""));
    const rowOf = (d: string) => {
      for (let r = 5; r <= ws.rowCount; r++) if (String(ws.getRow(r).getCell(1).value) === d) return ws.getRow(r);
      throw new Error(d);
    };
    const result = (row: ExcelJS.Row, h: string) => {
      const v = row.getCell(names.indexOf(h)).value as { formula?: string; result?: unknown } | null;
      return v && typeof v === "object" && "formula" in v ? (v.result ?? "") : v; // an empty cached result is not stored
    };
    const inc = rowOf("MCC-2 incomer");
    expect(result(inc, "Panel Incomer Breaker Rating (A)")).toBe(1000);
    expect(result(inc, "CT Ratio (Primary/1 A)")).toBe("1000/1 A");
    const pump = rowOf("Boiler feed water pump A (duty)");
    expect(result(pump, "Panel Incomer Breaker Rating (A)")).toBe("");
    expect(result(pump, "CT Ratio (Primary/1 A)")).toBe("");
  });

  it("a row with no current data shows blanks and a missing-input note, not zeros", async () => {
    const raw = [["Description", "Type", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)"], ["Mystery", "DOL", null, null, 433]];
    const { buffer, missing } = await mccWorkbookBuffer(raw, "s.xlsx");
    expect(missing).toBe(1);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet("Load list")!;
    const names = (ws.getRow(4).values as unknown[]).map((h) => String(h ?? ""));
    const cell = (h: string) => ws.getRow(5).getCell(names.indexOf(h)).value as { formula?: string; result?: unknown };
    for (const h of ["Calculated Current (A) – specified PF", "Design Current (A)", "Indicative Feeder Breaker Rating (A)"]) {
      expect(cell(h).formula, h).toBeTruthy();
      expect(cell(h).result ?? "", h).toBe(""); // blank, not 0
    }
  });
});

describe("rows for the page", () => {
  const sheet = [
    ["Description", "Type", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)"],
    ["Pump", "DOL", null, 185, 433], // 378.9 A calculated, design 473.6 A: review
    ["Fan", "VFD", null, 200, 433], // 305.5 A, design 381.8 A: 400 A MCCB
    ["Crusher", "DOL", null, 315, 433], // 665.2 A, design 831.4 A: 1000 A EDO ACB
    ["Dosing", "DOL", 180, null, 433], // entered 180 A, design 225 A: 250 A MCCB
    ["MCC-2 incomer", "Feeder", 800, null, 433], // design 1000 A: 1000 A EDO ACB, CT 1000/5
    ["Unknown", "DOL", null, null, 433],
  ];
  const r = mccRows(sheet);

  it("matches the workbook rules, row by row", () => {
    expect(r.errors).toEqual([]);
    expect(r.rows.map((x) => breakerLabel(x.breaker))).toEqual([
      "Engineering review required", "400 A MCCB", "1000 A EDO ACB", "250 A MCCB", "1000 A EDO ACB", "Engineering review required",
    ]);
    expect(r.rows[0].current.designA).toBeCloseTo(473.6, 1);
    expect(r.rows[3].current.source).toBe("Entered current");
  });

  it("incomer and CT only on the incomer row; relay only on DOL motors with a current", () => {
    expect(r.rows.map((x) => x.isIncomer)).toEqual([false, false, false, false, true, false]);
    expect(breakerLabel(r.rows[4].incomer!)).toBe("1000 A EDO ACB");
    expect(r.rows[4].ctPrimaryA).toBe(1000);
    expect(r.rows.filter((x) => x.ctPrimaryA !== undefined)).toHaveLength(1);
    expect(r.rows.map((x) => x.relayA)).toEqual([378.9, undefined, 645.2, 180, undefined, undefined]);
  });

  it("summary counts, with the row that has no data counted as missing, not as review", () => {
    expect(r.summary).toEqual({ loads: 6, mccb: 2, acb: 2, review: 1, missing: 1, verify: 2 });
  });

  it("a sheet without the required columns gives an error and no rows", () => {
    const bad = mccRows([["Name"], ["x"]]);
    expect(bad.rows).toEqual([]);
    expect(bad.errors[0]).toMatch(/Missing column/);
  });
});
