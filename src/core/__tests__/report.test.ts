import { readFileSync } from "node:fs";
import readExcelFile, { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import { DEFAULT_OPTIONS, buildReportBlob, parseLoadSheet, sizeTransformers, type RawLoadRow } from "..";

const file = readFileSync(new URL("../../../public/sample_load_list.xlsx", import.meta.url));
const buf = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;

async function report(includePower: boolean, volts = 433) {
  const parsed = parseLoadSheet(await readSheet(buf));
  const result = sizeTransformers(parsed.rows, volts);
  const blob = await buildReportBlob(result, {
    transformerLabel: `6600 / ${volts} V`, fileName: "sample.xlsx", options: DEFAULT_OPTIONS, includePower,
  });
  const sheets = await readExcelFile(await blob.arrayBuffer());
  return { result, sheets: Object.fromEntries(sheets.map((s) => [s.sheet, s.data])) as Record<string, unknown[][]> };
}
const find = (rows: unknown[][], label: string) => rows.find((r) => r[0] === label);

describe("Excel report", () => {
  it("has summary, loads, one sheet per transformer combination, and issues", async () => {
    const { sheets } = await report(true);
    expect(Object.keys(sheets)).toEqual(["Summary", "Loads", "2 Transformers", "3 Transformers", "Issues"]);
  });

  it("summary numbers match the calculation", async () => {
    const { result, sheets } = await report(true);
    const s = sheets["Summary"];
    expect(find(s, "Transformer")![1]).toBe("6600 / 433 V");
    expect(find(s, "Recommended")![1]).toBe("3 × 2000 kVA");
    expect(find(s, "Total kVA used for sizing")![1] as number).toBeCloseTo(result.totalKva, 6);
    expect(find(s, "Total kW (input)")![1] as number).toBeCloseTo(result.totals.kw, 6);
    expect(find(s, "Total kvar")![1] as number).toBeCloseTo(result.totals.kvar, 6);
    expect(find(s, "Working (W) / standby (S)")![1]).toBe("12 W / 2 S");
    expect(find(s, "Total kVA, vector sum")).toBeUndefined();
    expect(find(s, "Incomers")![1]).toBe(3);
    expect(find(s, "Bus couplers")![1]).toBe(2);
  });

  it("the kW / kvar / power-factor switch removes those columns and totals", async () => {
    const on = await report(true);
    const off = await report(false);
    const hdr = (r: Awaited<ReturnType<typeof report>>) => r.sheets["Loads"][0] as string[];
    expect(hdr(on)).toEqual(expect.arrayContaining(["W/S", "kW (input)", "kvar", "Power factor"]));
    expect(hdr(off)).toContain("W/S");
    for (const h of ["kW (input)", "kvar", "Power factor"]) expect(hdr(off)).not.toContain(h);
    expect(find(off.sheets["Summary"], "Total kW (input)")).toBeUndefined();
    expect(find(off.sheets["Summary"], "Total kvar")).toBeUndefined();
    expect(find(off.sheets["Summary"], "Total kVA used for sizing")).toBeDefined();
    // kVA is always present
    expect(hdr(off)).toContain("kVA");
    // the combination sheets follow the switch too
    const flat = (rows: unknown[][]) => rows.flat();
    expect(flat(off.sheets["3 Transformers"])).not.toContain("kvar");
    expect(flat(on.sheets["3 Transformers"])).toContain("kvar");
  });

  it("each combination sheet lists every included load exactly once, under its transformer", async () => {
    const { result, sheets } = await report(true);
    for (const n of [2, 3]) {
      const rows = sheets[`${n} Transformers`];
      const start = rows.findIndex((r) => r[0] === "Loads by transformer");
      const body = rows.slice(start + 2);
      expect(body).toHaveLength(result.loads.length);
      const names = new Set(body.map((r) => r[0]));
      expect([...names].sort()).toEqual(Array.from({ length: n }, (_, i) => `T${i + 1}`));
      expect(body.map((r) => r[1]).sort()).toEqual(result.loads.map((l) => l.rowNumber).sort());
    }
  });

  it("2-transformer sheet shows 2 x 4000 with one coupler; one-out table is present", async () => {
    const { sheets } = await report(true);
    const rows = sheets["2 Transformers"];
    expect(rows[0][0]).toBe("2 × 4000 kVA transformers, 6600 / 433 V");
    expect(String(rows[2][0])).toContain("2 incomers, 1 bus coupler");
    expect(rows.some((r) => r[0] === "T1" && r[4] === "OK")).toBe(true);
  });

  it("loads sheet carries W and S for each load", async () => {
    const { sheets } = await report(true);
    const rows = sheets["Loads"];
    const col = (rows[0] as string[]).indexOf("W/S");
    const values = rows.slice(1, -1).map((r) => r[col]);
    expect(values.filter((v) => v === "S")).toHaveLength(2);
    expect(values.filter((v) => v === "W")).toHaveLength(12);
  });

  it("issues sheet lists the excluded 690 V rows", async () => {
    const { sheets } = await report(true);
    const issues = sheets["Issues"];
    expect(issues.filter((r) => r[2] === "Excluded")).toHaveLength(2);
  });
});

describe("Excel report with added transformers", () => {
  it("has a sheet for each combination including added ones, and says so", async () => {
    const parsed = parseLoadSheet(await readSheet(buf));
    // Only the 2000 kVA rating and 2 transformers forced: 3471.6 kVA does not fit 2, so 3 is added.
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [2] };
    const result = sizeTransformers(parsed.rows, 433, opts);
    const blob = await buildReportBlob(result, { transformerLabel: "6600 / 433 V", options: opts, includePower: true });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((s) => [s.sheet, s.data]));
    expect(Object.keys(sheets)).toEqual(["Summary", "Loads", "2 Transformers", "3 Transformers", "Issues"]);
    expect(String(sheets["3 Transformers"][2][0])).toContain("Added because the selected number");
    expect(sheets["2 Transformers"][1][0]).toContain("DOES NOT FIT");
    expect(find(sheets["Summary"], "Recommended")![1]).toBe("3 × 2000 kVA");
  });
});

describe("Excel report with sets of transformers", () => {
  it("5 transformers = sets of 3 + 2: sets, couplers, names, and every load once", async () => {
    // 60 feeders of 83.33 kVA = 5,000 kVA; 3 x 2000 kVA (3,600) and 4 (2+2: 3,600) do not fit, 5 (3+2: 5,400) does
    const rows: RawLoadRow[] = Array.from({ length: 60 }, (_, i) => ({
      rowNumber: i + 2, description: `F${i}`, type: "Feeder", voltageV: 433, status: "W" as const,
      amps: (5000 / 60 * 1000) / (Math.sqrt(3) * 433),
    }));
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [3] };
    const result = sizeTransformers(rows, 433, opts);
    const blob = await buildReportBlob(result, { transformerLabel: "6600 / 433 V", options: opts, includePower: true });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((s) => [s.sheet, s.data]));
    // selected (3) and the final one (5) get a sheet; 4 is only in the summary table
    expect(Object.keys(sheets)).toEqual(["Summary", "Loads", "3 Transformers", "5 Transformers"]);
    const sum = sheets["Summary"];
    expect(find(sum, "Recommended")![1]).toBe("5 × 2000 kVA (sets of 3 + 2)");
    expect(find(sum, "Incomers")![1]).toBe(5);
    expect(find(sum, "Bus couplers")![1]).toBe(3);
    const tried = sum.filter((r) => typeof r[0] === "number" && typeof r[1] === "string" && /^[23]( \+ [23])*$/.test(r[1] as string));
    expect(tried.map((r) => [r[0], r[1], r[3]])).toEqual([[3, "3", "No"], [4, "2 + 2", "No"], [5, "3 + 2", "Yes"]]);
    const five = sheets["5 Transformers"];
    expect(String(five[3][0])).toContain("5 incomers, 3 bus couplers, in 2 independent sets (3 + 2)");
    expect(five.some((r) => String(r[0]).startsWith("Set 1: 3 transformers"))).toBe(true);
    expect(five.some((r) => String(r[0]).startsWith("Set 2: 2 transformers"))).toBe(true);
    const start = five.findIndex((r) => r[0] === "Loads by transformer");
    const body = five.slice(start + 2);
    expect(body).toHaveLength(60);
    expect([...new Set(body.map((r) => r[0]))].sort()).toEqual(["S1-T1", "S1-T2", "S1-T3", "S2-T1", "S2-T2"]);
  });
});
