import { readFileSync } from "node:fs";
import { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_OPTIONS, candidatesInOrder, distributeLoads, evaluateCandidate, parseLoadSheet,
  parseVoltage, sizeLoad, sizeTransformers, type RawLoadRow, type SizedLoad,
} from "..";

const row = (o: Partial<RawLoadRow>): RawLoadRow => ({ rowNumber: 2, description: "x", type: "", voltageV: 433, ...o });
const loadOf = (o: Partial<RawLoadRow>, v = 433): SizedLoad => {
  const r = sizeLoad(row(o), v);
  if (!("load" in r)) throw new Error(r.issue.message);
  return r.load;
};
const SQ3 = Math.sqrt(3);

describe("sizeLoad (hand-calculated)", () => {
  it("DOL motor from kW: kW / (0.7 * 0.95)", () => {
    expect(loadOf({ type: "DOL", kw: 100 }).kva).toBeCloseTo(100 / 0.665, 6); // 150.3759
  });
  it("VFD motor from kW: kW / (0.9 * 0.95)", () => {
    expect(loadOf({ type: "VFD", kw: 100 }).kva).toBeCloseTo(100 / 0.855, 6); // 116.9591
  });
  it("amps-only feeder: sqrt(3) * V * I", () => {
    const l = loadOf({ type: "Feeder", amps: 400 });
    expect(l.kind).toBe("feeder");
    expect(l.kva).toBeCloseTo((SQ3 * 433 * 400) / 1000, 6); // 299.996
    expect(l.kw).toBeCloseTo(l.kva * 0.7, 6);
  });
  it("amps-only with no type is a feeder; with DOL/VFD is a motor", () => {
    expect(loadOf({ type: "", amps: 100 }).kind).toBe("feeder");
    expect(loadOf({ type: "DOL", amps: 100 }).kind).toBe("motor");
    expect(loadOf({ type: "VFD", amps: 100 }).powerFactor).toBe(0.9);
  });
  it("kW wins over amps, and a feeder type with kW is treated as a DOL motor with a warning", () => {
    const l = loadOf({ type: "Feeder", kw: 50, amps: 999 });
    expect(l.kind).toBe("motor");
    expect(l.kva).toBeCloseTo(50 / 0.665, 6);
    expect(l.warnings.length).toBe(1);
  });
  it("uses the 690 V basis when selected", () => {
    expect(loadOf({ type: "Feeder", amps: 100, voltageV: 690 }, 690).kva).toBeCloseTo((SQ3 * 690 * 100) / 1000, 6);
  });
  it("excludes other voltages, errors on missing or bad data, warns on missing voltage", () => {
    const ex = sizeLoad(row({ type: "DOL", kw: 10, voltageV: 690 }), 433);
    expect("issue" in ex && ex.issue.severity).toBe("excluded");
    expect("issue" in sizeLoad(row({ type: "DOL" }), 433)).toBe(true);
    expect("issue" in sizeLoad(row({ type: "DOL", kw: -5 }), 433)).toBe(true);
    expect("issue" in sizeLoad(row({ kw: 5, badFields: ["Amps"] }), 433)).toBe(true);
    const nv = sizeLoad(row({ type: "DOL", kw: 10, voltageV: undefined }), 433);
    expect("load" in nv && nv.load.warnings[0]).toMatch(/assumed 433/);
  });
});

describe("candidates", () => {
  it("max loads: 80 % normal and 90 % with one out", () => {
    const m = (c: number, r: number) => evaluateCandidate(0, c, r).maxLoadKva;
    expect(m(2, 2000)).toBeCloseTo(1800); // one unit at 90 %
    expect(m(3, 2000)).toBeCloseTo(3600); // two units at 90 %
    expect(m(2, 2500)).toBeCloseTo(2250);
    expect(m(3, 2500)).toBeCloseTo(4500);
    expect(m(2, 4000)).toBeCloseTo(3600);
    expect(m(3, 4000)).toBeCloseTo(7200);
  });
  it("prefers 2 x 2000 and steps up only when it does not fit", () => {
    const first = (kva: number, stepUp: "min-installed" | "count-first" = "min-installed") =>
      candidatesInOrder(kva, { ...DEFAULT_OPTIONS, stepUp }).find((c) => c.fits);
    expect(first(1799)).toMatchObject({ count: 2, ratingKva: 2000 });
    expect(first(1801)).toMatchObject({ count: 2, ratingKva: 2500 });
    expect(first(2300)).toMatchObject({ count: 3, ratingKva: 2000 });
    expect(first(3601)).toMatchObject({ count: 3, ratingKva: 2500 });
    expect(first(4501)).toMatchObject({ count: 3, ratingKva: 4000 });
    expect(first(7201)).toBeUndefined();
    expect(first(1801, "count-first")).toMatchObject({ count: 3, ratingKva: 2000 });
  });
  it("normal-running limit can also be the binding one", () => {
    const o = { ...DEFAULT_OPTIONS, normalLimit: 0.5, contingencyLimit: 0.9 };
    expect(evaluateCandidate(1600, 3, 2000, o).normalOk).toBe(true); // 26.7 %
    expect(evaluateCandidate(2500, 2, 2000, o).normalOk).toBe(false); // 62.5 %
  });
});

describe("distributeLoads", () => {
  const mk = (kvas: number[]): SizedLoad[] => kvas.map((kva, i) => ({
    rowNumber: i + 2, description: `L${i}`, kind: "motor", type: "DOL", voltageV: 433, kw: kva * 0.7, kva, powerFactor: 0.7, basis: "", warnings: [],
  }));
  it("keeps every load exactly once and balances within the largest load", () => {
    const loads = mk([500, 400, 300, 250, 200, 150, 100, 90, 60]);
    const s = distributeLoads(loads, 3, 2000);
    expect(s.flatMap((x) => x.loads).length).toBe(loads.length);
    const total = loads.reduce((a, l) => a + l.kva, 0);
    expect(s.reduce((a, x) => a + x.kva, 0)).toBeCloseTo(total);
    const kvas = s.map((x) => x.kva);
    expect(Math.max(...kvas) - Math.min(...kvas)).toBeLessThanOrEqual(500);
  });
});

describe("spreadsheet parsing", () => {
  it("parses voltages", () => {
    expect(parseVoltage("433").value).toBe(433);
    expect(parseVoltage(433).value).toBe(433);
    expect(parseVoltage("690 V").value).toBe(690);
    expect(parseVoltage(0.433).value).toBe(433);
    expect(parseVoltage("0.69 kV").value).toBe(690);
    expect(parseVoltage("abc").bad).toBe(true);
    expect(parseVoltage("").bad).toBe(false);
  });
  it("finds columns in any order and skips blank rows", () => {
    const p = parseLoadSheet([
      ["Voltage", "kW", "Description", "Amps", "Type"],
      [433, 100, "Pump", null, "DOL"],
      [null, null, null, null, null],
      [433, null, "DB", 200, "Feeder"],
    ]);
    expect(p.errors).toEqual([]);
    expect(p.rows).toHaveLength(2);
    expect(p.rows[1]).toMatchObject({ rowNumber: 4, description: "DB", amps: 200, type: "Feeder" });
  });
  it("reports missing columns and empty sheets", () => {
    expect(parseLoadSheet([["Description", "kW"]]).errors[0]).toMatch(/Missing column/);
    expect(parseLoadSheet([]).errors[0]).toMatch(/empty/);
  });
  it("flags unreadable numbers", () => {
    const p = parseLoadSheet([["Description", "Type", "Amps", "kW", "Voltage"], ["x", "DOL", "abc", 10, 433]]);
    expect(p.rows[0].badFields).toEqual(["Amps"]);
  });
});

describe("sample workbook end to end", () => {
  const file = readFileSync(new URL("../../../public/sample_load_list.xlsx", import.meta.url));
  const buf = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;

  it("reads the sample and gives the hand-calculated 433 V result", async () => {
    const parsed = parseLoadSheet(await readSheet(buf));
    expect(parsed.errors).toEqual([]);
    expect(parsed.rows).toHaveLength(15);

    // Independent arithmetic from the sample's numbers:
    const dolKw = 185 + 160 + 160 + 132 + 90 + 315;
    const vfdKw = 250 + 200 + 75;
    const expected =
      dolKw / (0.7 * 0.95) + vfdKw / (0.9 * 0.95) + (SQ3 * 433 * (180 + 120 + 250 + 800)) / 1000;

    const r = sizeTransformers(parsed.rows, 433);
    expect(r.loads).toHaveLength(13);
    expect(r.issues.filter((i) => i.severity === "excluded")).toHaveLength(2);
    expect(r.totalKva).toBeCloseTo(expected, 6); // ~3193.4 kVA
    expect(r.recommended).toMatchObject({ count: 3, ratingKva: 2000, fits: true });
    expect(r.incomers).toBe(3);
    expect(r.couplers).toBe(2);
    expect(r.recommended!.contingencyPct).toBeCloseTo((expected / 4000) * 100, 6);
    expect(r.sections!.reduce((a, s) => a + s.kva, 0)).toBeCloseTo(expected, 6);
  });

  it("690 V group fits 2 x 2000", async () => {
    const parsed = parseLoadSheet(await readSheet(buf));
    const r = sizeTransformers(parsed.rows, 690);
    expect(r.loads).toHaveLength(2);
    expect(r.totalKva).toBeCloseTo(450 / 0.665 + 355 / 0.855, 6); // ~1091.9
    expect(r.recommended).toMatchObject({ count: 2, ratingKva: 2000 });
    expect(r.couplers).toBe(1);
  });
});
