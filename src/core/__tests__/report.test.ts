import { readFileSync } from "node:fs";
import readExcelFile, { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import { DEFAULT_OPTIONS, buildReportBlob, parseLoadSheet, sizeTransformers } from "..";

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
