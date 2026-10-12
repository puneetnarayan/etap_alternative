import { readFileSync } from "node:fs";
import { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_OPTIONS, candidatesInOrder, distributeLoads, evaluateCandidate, parseLoadSheet,
  parseStatus, parseVoltage, sizeLoad, sizeTransformers, type RawLoadRow, type SizedLoad,
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
  const mk = (specs: [number, string?][]): SizedLoad[] => specs.map(([kva, group], i) => ({
    rowNumber: i + 2, description: `L${i}`, kind: "motor", type: "DOL", voltageV: 433, kw: kva * 0.7,
    kvar: Math.sqrt(kva * kva - (kva * 0.7) ** 2), kva, powerFactor: 0.7, basis: "", group, warnings: [],
  }));
  it("keeps every load exactly once and balances within the largest load", () => {
    const loads = mk([[500], [400], [300], [250], [200], [150], [100], [90], [60]]);
    const { sections: s } = distributeLoads(loads, 3, 2000);
    expect(s.flatMap((x) => x.loads).length).toBe(loads.length);
    const total = loads.reduce((a, l) => a + l.kva, 0);
    expect(s.reduce((a, x) => a + x.kva, 0)).toBeCloseTo(total);
    const kvas = s.map((x) => x.kva);
    expect(Math.max(...kvas) - Math.min(...kvas)).toBeLessThanOrEqual(500);
  });
  it("puts members of a duty/standby group on different sections", () => {
    // The two 500 kVA pumps would otherwise both be placed first on the two emptiest sections anyway;
    // add many other loads so the standby would prefer to share a section if it could.
    const loads = mk([[500, "P"], [500, "P"], [450, "Q"], [450, "Q"], [100], [100], [100]]);
    for (const n of [2, 3]) {
      const { sections, notes } = distributeLoads(loads, n, 4000);
      expect(notes).toEqual([]);
      for (const g of ["P", "Q"]) {
        const homes = sections.map((s) => s.loads.filter((l) => l.group === g).length);
        expect(Math.max(...homes)).toBe(1);
      }
    }
  });
  it("is case/space-insensitive on group names and reports groups larger than the section count", () => {
    const loads = mk([[300, "cwp"], [300, " CWP "], [300, "Cwp"]]);
    const { sections, notes } = distributeLoads(loads, 2, 2000);
    expect(notes).toHaveLength(1);
    expect(sections.flatMap((s) => s.loads)).toHaveLength(3);
  });
  it("section figures follow the total basis", () => {
    const loads = mk([[100], [100]]);
    const arith = distributeLoads(loads, 1, 2000, { ...DEFAULT_OPTIONS, totalBasis: "arithmetic" }).sections[0];
    expect(arith.kva).toBeCloseTo(200);
    const sameKw = mk([[100], [100]]).map((l, i) => ({ ...l, kw: 70, kvar: i === 0 ? 71.414 : -71.414 }));
    const vec = distributeLoads(sameKw, 1, 2000, { ...DEFAULT_OPTIONS, totalBasis: "vector" }).sections[0];
    expect(vec.kva).toBeCloseTo(140, 1); // kvar cancels, kW adds
  });
});

describe("kW / kvar / totals", () => {
  it("motor: input kW = rated kW / efficiency, kvar = input kW * tan(acos pf), kVA^2 = kW^2 + kvar^2", () => {
    const l = loadOf({ type: "DOL", kw: 100 });
    expect(l.ratedKw).toBe(100);
    expect(l.kw).toBeCloseTo(100 / 0.95, 6); // 105.263
    expect(l.kvar).toBeCloseTo((100 / 0.95) * Math.tan(Math.acos(0.7)), 6); // 107.39
    expect(l.kva ** 2).toBeCloseTo(l.kw ** 2 + l.kvar ** 2, 6);
    expect(l.kw / l.kva).toBeCloseTo(0.7, 9);
  });
  it("amps-only feeder at pf 0.7 reports kW = 0.7 kVA", () => {
    const l = loadOf({ type: "Feeder", amps: 400 });
    expect(l.kw / l.kva).toBeCloseTo(0.7, 9);
  });
  it("arithmetic total is never below the vector total; basis switches the sizing figure", () => {
    const rows = [row({ rowNumber: 2, type: "DOL", kw: 100 }), row({ rowNumber: 3, type: "VFD", kw: 100 })];
    const a = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, totalBasis: "arithmetic" });
    const v = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, totalBasis: "vector" });
    expect(a.totals.kvaArithmetic).toBeCloseTo(100 / 0.665 + 100 / 0.855, 6);
    expect(a.totals.kvaVector).toBeLessThan(a.totals.kvaArithmetic);
    expect(a.totalKva).toBeCloseTo(a.totals.kvaArithmetic, 9);
    expect(v.totalKva).toBeCloseTo(v.totals.kvaVector, 9);
    expect(a.totals.powerFactor).toBeCloseTo(a.totals.kw / a.totals.kvaVector, 9);
  });
});

describe("combos and one-out check", () => {
  const rows = [row({ rowNumber: 2, type: "DOL", kw: 1000 }), row({ rowNumber: 3, type: "Feeder", amps: 1000 })];
  // kVA = 1000/0.665 + sqrt(3)*433*1000/1000 = 1503.76 + 749.99 = 2253.75
  const exp = 1000 / 0.665 + (SQ3 * 433 * 1000) / 1000;
  it("gives a result per transformer count, honouring the counts option", () => {
    const both = sizeTransformers(rows, 433);
    expect(both.combos.map((c) => c.count)).toEqual([2, 3]);
    const only3 = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, counts: [3] });
    expect(only3.combos.map((c) => c.count)).toEqual([3]);
    expect(only3.recommended).toMatchObject({ count: 3, ratingKva: 2000 });
    expect(only3.incomers).toBe(3);
    expect(only3.couplers).toBe(2);
  });
  it("each combination picks the smallest fitting rating; 2 x 2000 and 2 x 2500 fail, 2 x 4000 fits", () => {
    const r = sizeTransformers(rows, 433);
    expect(r.totalKva).toBeCloseTo(exp, 6);
    const two = r.combos.find((c) => c.count === 2)!;
    expect(two.fits).toBe(true);
    expect(two.chosen.ratingKva).toBe(4000); // 2253.75 > 2250 (2 x 2500 at 90 %)
    expect(two.couplers).toBe(1);
  });
  it("one-out scenarios: remaining units share the total and must stay <= 90 %", () => {
    const r = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [3] });
    const c = r.combos[0];
    expect(c.oneOut).toHaveLength(3);
    for (const s of c.oneOut) {
      expect(s.eachKva).toBeCloseTo(exp / 2, 6);
      expect(s.loadingPct).toBeCloseTo((exp / 2 / 2000) * 100, 6); // 56.3 %
      expect(s.ok).toBe(true);
    }
    expect(c.oneOut.map((s) => s.movedKva).reduce((a, b) => a + b, 0)).toBeCloseTo(exp, 6);
  });
  it("a single configured rating that is too small reports no fit but still shows the distribution", () => {
    const big = [row({ rowNumber: 2, type: "Feeder", amps: 4000 })]; // ~3000 kVA
    const r = sizeTransformers(big, 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [2] });
    expect(r.recommended).toBeUndefined();
    expect(r.combos[0].fits).toBe(false);
    expect(r.combos[0].chosen.ratingKva).toBe(2000);
    expect(r.combos[0].oneOut[0].ok).toBe(false);
    expect(r.combos[0].sections).toHaveLength(2);
  });
  it("custom rating list is used", () => {
    const r = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: [1600], counts: [2] });
    expect(r.combos[0].chosen.ratingKva).toBe(1600);
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
  it("reads an optional duty/standby group column (needs 'group' or 'pair' in the header)", () => {
    const p = parseLoadSheet([
      ["Description", "Type", "Amps", "kW", "Voltage", "Duty/standby group"],
      ["Pump A", "DOL", null, 100, 433, "CWP"],
      ["Pump B", "DOL", null, 100, 433, "CWP"],
      ["Fan", "VFD", null, 50, 433, null],
    ]);
    expect(p.rows.map((r) => r.group)).toEqual(["CWP", "CWP", undefined]);
    const q = parseLoadSheet([["Description", "Type", "Amps", "kW", "Voltage", "Duty/Standby"], ["a", "DOL", null, 1, 433, "Standby"]]);
    expect(q.rows[0].group).toBeUndefined();
  });
  it("reads a Working/Standby column as W / S", () => {
    expect(parseStatus("W").status).toBe("W");
    expect(parseStatus(" working ").status).toBe("W");
    expect(parseStatus("Duty").status).toBe("W");
    expect(parseStatus("S").status).toBe("S");
    expect(parseStatus("Stand-by").status).toBe("S");
    expect(parseStatus("").status).toBeUndefined();
    expect(parseStatus("maybe").unrecognised).toBe("maybe");
    const p = parseLoadSheet([
      ["Description", "Type", "Amps", "kW", "Voltage", "Working/Standby (W/S)", "Duty/standby group"],
      ["Pump A", "DOL", null, 100, 433, "W", "CWP"],
      ["Pump B", "DOL", null, 100, 433, "S", "CWP"],
      ["Fan", "VFD", null, 50, 433, "x", null],
    ]);
    // the group column is not mistaken for the W/S column and vice versa
    expect(p.rows.map((r) => r.status)).toEqual(["W", "S", undefined]);
    expect(p.rows.map((r) => r.group)).toEqual(["CWP", "CWP", undefined]);
    const r = sizeTransformers(p.rows, 433);
    expect(r.loads.map((l) => l.status)).toEqual(["W", "S", undefined]);
    expect(r.issues.some((i) => /not recognised \(use W or S\)/.test(i.message))).toBe(true);
  });
  it("accepts a W/S header on its own", () => {
    const p = parseLoadSheet([["Description", "Type", "Amps", "kW", "Voltage", "W/S"], ["Pump", "DOL", null, 10, 433, "S"]]);
    expect(p.rows[0].status).toBe("S");
    expect(p.rows[0].group).toBeUndefined();
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
    expect(parsed.rows).toHaveLength(16);
    expect(parsed.rows.filter((r) => r.group).map((r) => r.group)).toEqual(["BFW", "BFW", "CWP", "CWP"]);
    expect(parsed.rows.filter((r) => r.status === "S").map((r) => r.description)).toEqual(["Boiler feed water pump B", "Cooling water pump B"]);
    expect(parsed.rows.every((r) => r.status === "W" || r.status === "S")).toBe(true);

    // Independent arithmetic from the sample's numbers:
    const dolKw = 185 + 185 + 160 + 160 + 132 + 90 + 315;
    const vfdKw = 250 + 200 + 75;
    const expected =
      dolKw / (0.7 * 0.95) + vfdKw / (0.9 * 0.95) + (SQ3 * 433 * (180 + 120 + 250 + 800)) / 1000;

    const r = sizeTransformers(parsed.rows, 433);
    expect(r.loads).toHaveLength(14);
    expect(r.issues.filter((i) => i.severity === "excluded")).toHaveLength(2);
    expect(r.totalKva).toBeCloseTo(expected, 6); // ~3471.6 kVA
    expect(r.recommended).toMatchObject({ count: 3, ratingKva: 2000, fits: true });
    expect(r.incomers).toBe(3);
    expect(r.couplers).toBe(2);
    expect(r.recommended!.contingencyPct).toBeCloseTo((expected / 4000) * 100, 6); // ~86.8 %
    const three = r.combos.find((c) => c.count === 3)!;
    expect(three.sections.reduce((a, s) => a + s.kva, 0)).toBeCloseTo(expected, 6);
    // duty/standby pairs end up on different sections in both combinations
    for (const combo of r.combos) {
      for (const g of ["BFW", "CWP"]) {
        const per = combo.sections.map((s) => s.loads.filter((l) => l.group === g).length);
        expect(Math.max(...per)).toBe(1);
      }
    }
    // 2 transformers: 3471.6 kVA needs 2 x 4000 (2 x 2500 carries only 2250 with one out)
    expect(r.combos.find((c) => c.count === 2)!.chosen.ratingKva).toBe(4000);
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
