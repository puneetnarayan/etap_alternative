import { readFileSync } from "node:fs";
import { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_OPTIONS, STANDARD_MCC_FEEDERS, candidatesInOrder, displayCombos, distributeLoads, evaluateCandidate, failureReason, layoutFor, parseLoadSheet, spareFeeders, spareKey,
  parseStatus, parseVoltage, setCapacityKva, sizeLoad, sizeTransformers, type RawLoadRow, type SizedLoad,
} from "..";

const row = (o: Partial<RawLoadRow>): RawLoadRow => ({ rowNumber: 2, description: "x", type: "", voltageV: 433, ...o });
const loadOf = (o: Partial<RawLoadRow>, v = 433): SizedLoad => {
  const r = sizeLoad(row(o), v);
  if (!("load" in r)) throw new Error(r.issue.message);
  return r.load;
};
const SQ3 = Math.sqrt(3);
/** Start with both a set of 2 and a set of 3 (the page's default starts with 3 only). */
const BOTH = { ...DEFAULT_OPTIONS, counts: [2, 3] };

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
  it("amps-only loads at the 433 V level are converted at 433 V (sqrt(3) x 433 x A)", () => {
    const l = loadOf({ type: "Feeder", amps: 400 });
    expect(l.kva).toBeCloseTo((SQ3 * 433 * 400) / 1000, 6); // 300.0 kVA
    expect(l.basis).toContain("√3 × 433 V × 400 A");
    expect(l.voltageV).toBe(433);
    expect(l.kw / l.kva).toBeCloseTo(0.7, 9);
    // motors given in kW do not depend on the voltage
    expect(loadOf({ type: "DOL", kw: 100 }).kva).toBeCloseTo(100 / 0.665, 6);
  });
  it("the calculation voltage can be changed or removed; unlisted levels use their own voltage", () => {
    const at = (calc: Record<number, number>, amps: number, v = 433) => {
      const r = sizeLoad(row({ type: "Feeder", amps, voltageV: v }), v, { ...DEFAULT_OPTIONS, calculationVoltageV: calc });
      if (!("load" in r)) throw new Error(r.issue.message);
      return r.load.kva;
    };
    expect(DEFAULT_OPTIONS.calculationVoltageV).toEqual({}); // 433 V loads are calculated at 433 V
    expect(at({}, 1000)).toBeCloseTo((SQ3 * 433 * 1000) / 1000, 6); // no mapping: 433 V as labelled
    expect(at({ 433: 415 }, 1000)).toBeCloseTo((SQ3 * 415 * 1000) / 1000, 6);
    expect(at({ 433: 420 }, 1000, 690)).toBeCloseTo((SQ3 * 690 * 1000) / 1000, 6); // 690 V is not remapped
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
      candidatesInOrder(kva, { ...BOTH, stepUp }).find((c) => c.fits);
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
    kvar: Math.sqrt(kva * kva - (kva * 0.7) ** 2), kva, current: 0, powerFactor: 0.7, basis: "", group, warnings: [],
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
  const rows = [row({ rowNumber: 2, type: "DOL", kw: 1000 }), row({ rowNumber: 3, type: "Feeder", amps: 950 })];
  // kVA = 1000/0.665 + sqrt(3)*433*950/1000 = 1503.76 + 712.46 = 2216.2 (amps are converted at 433 V)
  const exp = 1000 / 0.665 + (SQ3 * 433 * 950) / 1000;
  it("gives a result per transformer count, honouring the counts option", () => {
    expect(sizeTransformers(rows, 433).combos.map((c) => c.count)).toEqual([3]); // default: start with 3
    const both = sizeTransformers(rows, 433, BOTH);
    expect(both.combos.map((c) => c.count)).toEqual([2, 3]);
    const only3 = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, counts: [3] });
    expect(only3.combos.map((c) => c.count)).toEqual([3]);
    expect(only3.recommended).toMatchObject({ count: 3, ratingKva: 2000 });
    expect(only3.incomers).toBe(3);
    expect(only3.couplers).toBe(2);
  });
  it("each combination picks the smallest fitting rating (2 x 2000 carries 1800 kVA, 2 x 2500 carries 2250 kVA)", () => {
    const r = sizeTransformers(rows, 433, BOTH);
    expect(r.totalKva).toBeCloseTo(exp, 6); // 2,216.2 kVA
    const two = r.combos.find((c) => c.count === 2)!;
    expect(two.fits).toBe(true);
    expect(two.chosen.ratingKva).toBe(2500); // 2,216.2 <= 2,250 (2 x 2000 would carry only 1,800)
    expect(two.candidates.map((c) => [c.ratingKva, c.fits])).toEqual([[2000, false], [2500, true], [4000, true]]);
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
    const r = sizeTransformers(big, 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [2], maxTransformers: 2 });
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
  const file = readFileSync(new URL("./fixtures/small_plant_load_list.xlsx", import.meta.url));
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

    const r = sizeTransformers(parsed.rows, 433, BOTH);
    expect(r.loads).toHaveLength(14);
    expect(sizeTransformers(parsed.rows, 433).recommended).toMatchObject({ count: 3, ratingKva: 2000 }); // starts with 3
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
    const r = sizeTransformers(parsed.rows, 690, BOTH);
    expect(r.loads).toHaveLength(2);
    expect(r.totalKva).toBeCloseTo(450 / 0.665 + 355 / 0.855, 6); // ~1091.9
    expect(r.recommended).toMatchObject({ count: 2, ratingKva: 2000 }); // 2 x 2000 carries 1800 kVA
    expect(r.couplers).toBe(1);
    // the page starts with 3 transformers; a set of 2 is used first only when asked for
    expect(sizeTransformers(parsed.rows, 690).recommended).toMatchObject({ count: 3, ratingKva: 2000 });
    expect(sizeTransformers(parsed.rows, 690, { ...DEFAULT_OPTIONS, counts: [2] }).recommended).toMatchObject({ count: 2 });
  });
});

describe("more transformers = independent sets of 2 or 3", () => {
  // For 2000 kVA units: a set of 2 carries 0.9 * 2000 = 1800 kVA, a set of 3 carries 0.9 * 2 * 2000 = 3600 kVA
  // (the one-transformer-out limit binds). Capacity by number of transformers:
  //   2: 1800 | 3: 3600 | 4 (2+2): 3600 | 5 (3+2): 5400 | 6 (3+3): 7200 | 7 (3+2+2): 7200
  //   8 (3+3+2): 9000 | 9 (3+3+3): 10800 | 10 (3+3+2+2): 10800 | 11: 12600 | 12: 14400 | 20 (6x3 + 2): 23400
  const loadsKva = (kva: number, parts = 60): RawLoadRow[] =>
    Array.from({ length: parts }, (_, i) =>
      row({ rowNumber: i + 2, description: `F${i}`, type: "Feeder", amps: (kva / parts * 1000) / (SQ3 * 433) }));
  const only2000 = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
  const cap = (n: number) => evaluateCandidate(0, n, 2000).maxLoadKva;

  it("splits into sets of 2 or 3 with as few sets as possible", () => {
    expect([2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map(layoutFor)).toEqual([
      [2], [3], [2, 2], [3, 2], [3, 3], [3, 2, 2], [3, 3, 2], [3, 3, 3], [3, 3, 2, 2], [3, 3, 3, 2], [3, 3, 3, 3],
    ]);
    for (let n = 2; n <= 20; n++) {
      const l = layoutFor(n);
      expect(l.reduce((a, k) => a + k, 0)).toBe(n);
      expect(l.every((k) => k === 2 || k === 3)).toBe(true);
    }
    expect(layoutFor(1)).toEqual([]);
  });
  it("capacity by number of transformers (hand-calculated)", () => {
    const expected: Record<number, number> = { 2: 1800, 3: 3600, 4: 3600, 5: 5400, 6: 7200, 7: 7200, 8: 9000, 9: 10800, 10: 10800, 11: 12600, 12: 14400, 20: 23400 };
    for (const [n, c] of Object.entries(expected)) expect(cap(Number(n))).toBeCloseTo(c, 6);
  });
  it("incomers and bus couplers: one coupler between neighbouring incomers within a set, none between sets", () => {
    const rows = loadsKva(7000);
    const r = sizeTransformers(rows, 433, { ...only2000, counts: [7] });
    const c = r.combos[0];
    expect(c.layout).toEqual([3, 2, 2]);
    expect(c.incomers).toBe(7);
    expect(c.couplers).toBe(4); // 2 + 1 + 1
    expect(c.sets.map((s) => [s.size, s.incomers, s.couplers])).toEqual([[3, 3, 2], [2, 2, 1], [2, 2, 1]]);
    expect(c.sets.flatMap((s) => s.sections).map((s) => s.name)).toEqual([
      "S1-T1", "S1-T2", "S1-T3", "S2-T1", "S2-T2", "S3-T1", "S3-T2",
    ]);
  });
  it("a selected number that fits is left alone (nothing added); the default starts with 3", () => {
    const r = sizeTransformers(loadsKva(1500), 433, only2000);
    expect(r.combos.map((c) => [c.count, c.added])).toEqual([[3, false]]);
    expect(r.recommended).toMatchObject({ count: 3, ratingKva: 2000 });
    const r2 = sizeTransformers(loadsKva(1500), 433, { ...only2000, counts: [2, 3] });
    expect(r2.combos.map((c) => [c.count, c.added])).toEqual([[2, false], [3, false]]);
    expect(r2.recommended).toMatchObject({ count: 2, ratingKva: 2000 });
  });
  it("forced 2 that does not fit adds 3", () => {
    const r = sizeTransformers(loadsKva(3000), 433, { ...only2000, counts: [2] });
    expect(r.combos.map((c) => [c.count, c.fits, c.added])).toEqual([[2, false, false], [3, true, true]]);
    expect(r.recommended).toMatchObject({ count: 3 });
    expect(r.couplers).toBe(2);
  });
  it("3 with 5000 kVA does not fit; 4 (2+2) adds nothing so it is skipped; 5 (3+2, 5400) does", () => {
    const r = sizeTransformers(loadsKva(5000), 433, { ...only2000, counts: [3] });
    expect(r.combos.map((c) => [c.count, c.fits])).toEqual([[3, false], [5, true]]);
    expect(r.recommended).toMatchObject({ count: 5 });
    const five = r.combos[1];
    expect(five.layout).toEqual([3, 2]);
    expect(five.incomers).toBe(5);
    expect(five.couplers).toBe(3);
    // 5000 / ((5 - 2 sets) * 2000) = 83.3 % with one transformer out in each set
    expect(five.chosen.contingencyPct).toBeCloseTo(83.333, 2);
    expect(five.chosen.normalPct).toBeCloseTo(50, 6);
  });
  it("starts with 3, adds sets of 3 and a set of 2 at the end, skipping numbers that add nothing (10,000 kVA -> 9)", () => {
    const r = sizeTransformers(loadsKva(10000), 433, only2000);
    // 4 = 2+2 and 7 = 3+2+2 carry no more than 3 and 6, so they are skipped
    expect(r.combos.map((c) => c.count)).toEqual([3, 5, 6, 8, 9]);
    expect(r.combos.map((c) => c.layout)).toEqual([[3], [3, 2], [3, 3], [3, 3, 2], [3, 3, 3]]);
    expect(r.combos.filter((c) => c.added).map((c) => c.count)).toEqual([5, 6, 8, 9]);
    expect(r.combos.filter((c) => c.fits).map((c) => c.count)).toEqual([9]);
    expect(r.recommended).toMatchObject({ count: 9 });
    // only the one started with and the final one are shown in full
    expect(displayCombos(r).map((c) => c.count)).toEqual([3, 9]);
  });
  it("exact boundary: 3600 kVA fits 3, 3601 does not", () => {
    expect(sizeTransformers(loadsKva(3599.9), 433, only2000).combos).toHaveLength(1);
    const over = sizeTransformers(loadsKva(3601), 433, only2000);
    expect(over.combos.map((c) => [c.count, c.fits])).toEqual([[3, false], [5, true]]);
  });
  it("each set carries its share: loads are filled in proportion to what each set can carry", () => {
    const r = sizeTransformers(loadsKva(5000), 433, { ...only2000, counts: [5] });
    const [s3, s2] = r.combos[0].sets;
    expect(s3.kva + s2.kva).toBeCloseTo(5000, 3);
    // proportional to 3600 : 1800, within one 83 kVA load
    expect(s3.kva).toBeGreaterThan(3333.3 - 90);
    expect(s3.kva).toBeLessThan(3333.3 + 90);
    for (const s of [s3, s2]) {
      expect(s.kva).toBeLessThanOrEqual(s.capacityKva + 90);
      for (const o of s.oneOut) {
        expect(o.eachKva).toBeCloseTo(s.kva / (s.size - 1), 6);
        expect(o.loadingPct).toBeCloseTo((s.kva / (s.size - 1) / 2000) * 100, 6);
        expect(o.loadingPct).toBeLessThan(90 + 5);
      }
    }
  });
  it("8 indivisible feeders of 750 kVA on 6 transformers (3 + 3): 4 per set, no overload, no notes", () => {
    const rows: RawLoadRow[] = Array.from({ length: 8 }, (_, i) =>
      row({ rowNumber: i + 2, description: `F${i}`, type: "Feeder", amps: 750 * 1000 / (SQ3 * 433) }));
    const r = sizeTransformers(rows, 433, { ...only2000, counts: [6] });
    const c = r.combos[0];
    expect(c.layout).toEqual([3, 3]);
    expect(c.sets.map((x) => x.sections.reduce((a, sec) => a + sec.loads.length, 0))).toEqual([4, 4]);
    expect(c.sets.map((x) => x.kva)).toEqual([expect.closeTo(3000, 1), expect.closeTo(3000, 1)]);
    expect(c.notes).toEqual([]);
    expect(c.oneOut.every((o) => o.loadingPct < 90 && o.ok)).toBe(true);
    expect(c.oneOut[0].loadingPct).toBeCloseTo((3000 / 2 / 2000) * 100, 1); // 75 %
  });
  it("sets are balanced before sections: sections of one set stay within one load of each other", () => {
    const c = sizeTransformers(loadsKva(5000, 7), 433, { ...only2000, counts: [5] }).combos[0];
    for (const set of c.sets) {
      const per = set.sections.map((x) => x.loads.length);
      expect(Math.max(...per) - Math.min(...per)).toBeLessThanOrEqual(1);
    }
  });
  it("up to 20 transformers by default; gives up there with no recommendation", () => {
    expect(DEFAULT_OPTIONS.maxTransformers).toBe(20);
    const r = sizeTransformers(loadsKva(30000, 120), 433, only2000); // 23,400 kVA is the most 20 can carry
    expect(r.recommended).toBeUndefined();
    expect(Math.max(...r.combos.map((c) => c.count))).toBe(20);
    expect(r.combos.every((c) => !c.fits)).toBe(true);
    expect(r.combos.map((c) => c.count)).toEqual([3, 5, 6, 8, 9, 11, 12, 14, 15, 17, 18, 20]);
    expect(displayCombos(r).map((c) => c.count)).toEqual([3, 20]);
    // 22,000 kVA in 120 loads of 183.3 kVA can be placed on 20 (a set of 3 holds 19, the set of 2 holds 9)
    const fits20 = sizeTransformers(loadsKva(22000, 120), 433, only2000);
    expect(fits20.recommended).toMatchObject({ count: 20 });
    // 23,000 kVA is within the ideal 23,400 kVA but whole loads cannot be packed that tightly (117 of 120 fit)
    const tight = sizeTransformers(loadsKva(23000, 120), 433, only2000);
    const last = tight.combos[tight.combos.length - 1];
    expect(last.count).toBe(20);
    expect(last.chosen.normalOk && last.chosen.contingencyOk).toBe(true);
    expect(last.chosen.balanced).toBe(false);
    expect(tight.recommended).toBeUndefined();
    expect(failureReason(last.chosen)).toBe("loads cannot be balanced within the limits");
    expect(sizeTransformers(loadsKva(30000, 120), 433, { ...only2000, maxTransformers: 3 }).combos.map((c) => c.count)).toEqual([3]);
  });
  it("standby pairs stay on different sections with several sets", () => {
    const rows = [
      ...loadsKva(5000),
      row({ rowNumber: 100, description: "A", type: "DOL", kw: 100, status: "W", group: "P" }),
      row({ rowNumber: 101, description: "B", type: "DOL", kw: 100, status: "S", group: "P" }),
    ];
    const c = sizeTransformers(rows, 433, { ...only2000, counts: [5] }).combos[0];
    expect(c.sections.map((s) => s.loads.filter((l) => l.group === "P").length).every((n) => n <= 1)).toBe(true);
  });
  it("every load is placed exactly once, with sets", () => {
    const r = sizeTransformers(loadsKva(9000), 433, { ...only2000, counts: [8] });
    expect(r.combos[0].sections.flatMap((s) => s.loads)).toHaveLength(60);
    expect(r.combos[0].sets.reduce((a, s) => a + s.kva, 0)).toBeCloseTo(9000, 3);
  });
});

describe("real load balancing decides whether a number of transformers fits", () => {
  // For 2000 kVA units: a section carries at most 0.8 * 2000 = 1600 kVA, a set of 3 carries 3600 kVA, a set of 2 carries 1800 kVA.
  const kvaRows = (kvas: number[], extra: Partial<RawLoadRow>[] = []): RawLoadRow[] =>
    kvas.map((kva, i) => row({ rowNumber: i + 2, description: `L${i}`, type: "Feeder", amps: (kva * 1000) / (SQ3 * 433), ...(extra[i] ?? {}) }));
  const only2000 = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };

  it("5 loads of 1000 kVA: 5 transformers (3+2) would carry 5400 kVA ideally but the loads cannot be placed, so 6 is used", () => {
    // set of 3 holds 3 loads (3000 <= 3600; a 4th makes 4000), set of 2 holds 1 load (2000 > 1800): only 4 of 5 fit on 5 transformers
    const r = sizeTransformers(kvaRows([1000, 1000, 1000, 1000, 1000]), 433, only2000);
    expect(r.combos.map((c) => c.count)).toEqual([3, 5, 6]);
    const five = r.combos[1];
    expect(five.chosen.normalOk && five.chosen.contingencyOk).toBe(true); // ideal sharing is within both limits
    expect(five.chosen.balanced).toBe(false); // ... but whole loads cannot be balanced
    expect(five.fits).toBe(false);
    expect(failureReason(five.chosen)).toBe("loads cannot be balanced within the limits");
    expect(five.notes.some((n) => /above/.test(n))).toBe(true);
    expect(r.combos[2]).toMatchObject({ count: 6, fits: true, added: true });
    expect(r.recommended).toMatchObject({ count: 6 });
    expect(r.combos[2].layout).toEqual([3, 3]);
  });

  it("the placement is repaired when the first attempt leaves a set over its limit (but a valid one exists)", () => {
    // Largest-first placement gives the set of 2 1900 kVA (> 1800); moving loads fixes it: e.g. 1500 alone in the set of 2.
    for (const kvas of [[1500, 1100, 900, 800, 250, 200], [1500, 1050, 800, 800, 450, 400, 350], [1550, 1200, 1100, 700, 450, 200], [1350, 1100, 1100, 800, 350, 250]]) {
      const r = sizeTransformers(kvaRows(kvas), 433, { ...only2000, counts: [5] });
      const c = r.combos[0];
      expect(c.added).toBe(false);
      expect(c.fits, kvas.join(",")).toBe(true);
      expect(c.chosen.balanced).toBe(true);
      expect(r.combos).toHaveLength(1); // nothing had to be added
      const [s3, s2] = c.sets;
      expect(s3.kva).toBeLessThanOrEqual(3600 + 1e-6);
      expect(s2.kva).toBeLessThanOrEqual(1800 + 1e-6);
      for (const sec of c.sections) expect(sec.kva).toBeLessThanOrEqual(1600 + 1e-6);
      expect(c.sections.flatMap((s) => s.loads)).toHaveLength(kvas.length);
      expect(s3.kva + s2.kva).toBeCloseTo(kvas.reduce((a, b) => a + b, 0), 6);
      expect(c.oneOut.every((o) => o.ok)).toBe(true);
      expect(c.notes).toEqual([]);
    }
  });

  it("repairing keeps duty/standby pairs on different sections", () => {
    const kvas = [1500, 1100, 900, 800, 250, 200];
    const groups = kvas.map((_, i) => ({ group: i === 4 || i === 5 ? "G" : i === 0 || i === 2 ? "H" : undefined }));
    const c = sizeTransformers(kvaRows(kvas, groups), 433, { ...only2000, counts: [5] }).combos[0];
    expect(c.fits).toBe(true);
    for (const g of ["G", "H"]) {
      expect(Math.max(...c.sections.map((s) => s.loads.filter((l) => l.group === g).length))).toBe(1);
    }
  });

  it("a load too big for any section (> 80 % of the rating) cannot be helped by more transformers", () => {
    // 1700 kVA is 85 % of 2000 kVA: not allowed in normal running on any single transformer
    const r = sizeTransformers(kvaRows([1700, 300]), 433, only2000);
    expect(r.blockers).toHaveLength(1);
    expect(r.blockers[0]).toContain("L0");
    expect(r.blockers[0]).toContain("1,700 kVA");
    expect(r.blockers[0]).toContain("1,600 kVA"); // whole numbers; only percentages have a decimal
    expect(r.recommended).toBeUndefined();
    expect(r.combos.map((c) => c.count)).toEqual([3]); // nothing added: it would never help
    // a larger rating makes it placeable: 80 % of 2500 = 2000 kVA
    const big = sizeTransformers(kvaRows([1700, 300]), 433, { ...DEFAULT_OPTIONS, ratingsKva: [2500] });
    expect(big.blockers).toEqual([]);
    expect(big.recommended).toMatchObject({ count: 3, ratingKva: 2500 });
    expect(sizeTransformers(kvaRows([100, 200]), 433, only2000).blockers).toEqual([]);
  });

  it("the page's default starts with 3 and uses 2 only at the end, as the remainder", () => {
    expect(DEFAULT_OPTIONS.counts).toEqual([3]);
    // 4750 kVA (several loads) needs 3 + 2 = 5: sets of 3 first, a set of 2 for the remainder
    const r = sizeTransformers(kvaRows([1500, 1100, 900, 800, 250, 200]), 433, only2000);
    expect(r.combos.map((c) => c.layout)).toEqual([[3], [3, 2]]);
    expect(r.recommended).toMatchObject({ count: 5 });
    // every order of attempts ends sets of 2 last
    for (const c of r.combos) expect(c.layout).toEqual([...c.layout].sort((a, b) => b - a));
  });

  it("a total that fits ideally is not enough on its own: 'fits' needs the real placement", () => {
    const rows = kvaRows([1000, 1000, 1000, 1000, 1000]);
    const ideal = evaluateCandidate(5000, 5, 2000);
    expect(ideal.fits).toBe(true); // 5000 <= 5400
    const real = sizeTransformers(rows, 433, { ...only2000, counts: [5] }).combos[0];
    expect(real.fits).toBe(false);
  });

  it("fuzz: whatever the loads, each load is placed exactly once and 'fits' really means within every limit", () => {
    let seed = 12345;
    const rnd = () => { seed = (seed * 1664525 + 1013904223) % 4294967296; return seed / 4294967296; };
    for (let trial = 0; trial < 120; trial++) {
      const n = 3 + Math.floor(rnd() * 40);
      const kvas = Array.from({ length: n }, () => 50 + Math.floor(rnd() * 70) * 20); // 50 .. 1430 kVA
      const extra = kvas.map(() => (rnd() < 0.15 ? { group: `G${Math.floor(rnd() * 3)}` } : {}));
      const opts = { ...only2000, counts: [[2], [3], [3]][Math.floor(rnd() * 3)] };
      const r = sizeTransformers(kvaRows(kvas, extra), 433, opts);
      for (const c of r.combos) {
        const placed = c.sections.flatMap((x) => x.loads);
        expect(placed).toHaveLength(kvas.length);
        expect(new Set(placed.map((l) => l.rowNumber)).size).toBe(kvas.length);
        if (c.fits) {
          for (const sec of c.sections) expect(sec.kva).toBeLessThanOrEqual(1600 + 1e-6);
          for (const set of c.sets) expect(set.kva).toBeLessThanOrEqual(set.capacityKva + 1e-6);
          for (const g of ["G0", "G1", "G2"]) {
            const size = placed.filter((l) => l.group === g).length;
            if (size <= c.count) expect(Math.max(0, ...c.sections.map((x) => x.loads.filter((l) => l.group === g).length))).toBeLessThanOrEqual(1);
          }
        }
      }
      if (r.recommended) expect(r.combos.some((c) => c.fits && c.count === r.recommended!.count)).toBe(true);
    }
  });

  it("stays fast when no placement exists (tight 20-transformer case)", () => {
    const t0 = Date.now();
    const rows = kvaRows(Array.from({ length: 120 }, () => 23000 / 120));
    const r = sizeTransformers(rows, 433, only2000);
    expect(r.recommended).toBeUndefined();
    expect(Date.now() - t0).toBeLessThan(5000);
  });
});

describe("spare feeders: one of each type and rating, output only", () => {
  const mk = (rows: Partial<RawLoadRow>[]) => rows.map((r, i) => row({ rowNumber: i + 2, description: `L${i}`, ...r }));
  const loadsOf = (rows: Partial<RawLoadRow>[]) => sizeTransformers(mk(rows), 433).loads;

  it("one spare per type and rating (kW and amps are different ratings), in the order DOL, VFD, Feeder", () => {
    const loads = loadsOf([
      { type: "VFD", kw: 75 }, { type: "DOL", kw: 185 }, { type: "DOL", kw: 185 }, { type: "DOL", kw: 90 },
      { type: "DOL", amps: 180 }, { type: "VFD", amps: 120 }, { type: "Feeder", amps: 800 }, { type: "Feeder", amps: 250 }, { type: "Feeder", amps: 250 },
      { type: "VFD", kw: 75 },
    ]);
    const spares = spareFeeders(loads);
    expect(spares.map((x) => `${x.type} ${x.rating} ${x.unit} x${x.count}`)).toEqual([
      "DOL 90 kW x1", "DOL 185 kW x2", "DOL 180 A x1", "VFD 75 kW x2", "VFD 120 A x1", "FEEDER 250 A x2", "FEEDER 800 A x1",
    ]);
    expect(spares.map((x) => x.description)).toEqual([
      "Spare DOL motor feeder, 90 kW (Spare)", "Spare DOL motor feeder, 185 kW (Spare)", "Spare DOL motor feeder, 180 A (Spare)",
      "Spare VFD motor feeder, 75 kW (Spare)", "Spare VFD motor feeder, 120 A (Spare)", "Spare feeder, 250 A (Spare)", "Spare feeder, 800 A (Spare)",
    ]);
    expect(new Set(spares.map((x) => x.key)).size).toBe(spares.length);
    // the key is the same one the loads carry
    for (const l of loads) expect(spares.some((x) => x.key === spareKey(l))).toBe(true);
  });

  it("a feeder with a kW rating is a DOL motor, so it gets a DOL spare, not a feeder spare", () => {
    expect(spareFeeders(loadsOf([{ type: "Feeder", kw: 50 }])).map((x) => x.description)).toEqual(["Spare DOL motor feeder, 50 kW (Spare)"]);
  });

  it("the sample file gives 12 spares (5 DOL kW, 1 DOL amps, 3 VFD kW, 1 VFD amps, 2 feeders)", () => {
    const small = sizeTransformers(
      mk([
        { type: "DOL", kw: 185 }, { type: "DOL", kw: 185 }, { type: "DOL", kw: 160 }, { type: "DOL", kw: 160 }, { type: "VFD", kw: 250 }, { type: "VFD", kw: 250 },
        { type: "VFD", kw: 200 }, { type: "DOL", kw: 132 }, { type: "DOL", kw: 90 }, { type: "VFD", kw: 75 }, { type: "DOL", kw: 315 },
        { type: "DOL", amps: 180 }, { type: "VFD", amps: 120 }, { type: "Feeder", amps: 250 }, { type: "Feeder", amps: 800 },
      ]),
      433,
    );
    expect(small.spares).toHaveLength(12);
  });

  it("spares never enter the calculation: no total, section, set or recommendation includes them", () => {
    const rows = mk([{ type: "DOL", kw: 185 }, { type: "DOL", kw: 185 }, { type: "VFD", kw: 250 }, { type: "Feeder", amps: 800 }, { type: "Feeder", amps: 250 }]);
    const r = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000] });
    expect(r.spares.length).toBeGreaterThan(0);
    const expected = 185 / 0.665 * 2 + 250 / 0.855 + (SQ3 * 433 * (800 + 250)) / 1000; // loads only
    expect(r.totalKva).toBeCloseTo(expected, 6);
    expect(r.loads).toHaveLength(5); // the spares are not loads
    for (const c of r.combos) {
      expect(c.sections.flatMap((x) => x.loads)).toHaveLength(5);
      expect(c.sections.reduce((a, x) => a + x.kva, 0)).toBeCloseTo(expected, 6);
      expect(c.sets.reduce((a, x) => a + x.kva, 0)).toBeCloseTo(expected, 6);
    }
    // adding more loads of an existing type/rating changes the totals but not the number of spares
    const more = sizeTransformers([...rows, ...mk([{ type: "DOL", kw: 185 }])], 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000] });
    expect(more.spares).toHaveLength(r.spares.length);
    expect(more.totalKva).toBeGreaterThan(r.totalKva);
    // a spare carries no electrical figures
    for (const sp of r.spares) expect(Object.keys(sp).sort()).toEqual(["count", "description", "key", "kind", "rating", "type", "unit"]);
  });

  it("no loads, no spares", () => {
    expect(sizeTransformers([], 433).spares).toEqual([]);
  });
});

describe("PMCC / IMCC feeders in every set", () => {
  const mk = (rows: Partial<RawLoadRow>[]) => rows.map((r, i) => row({ rowNumber: i + 2, description: `L${i}`, ...r }));
  const MCC = { ...DEFAULT_OPTIONS, ratingsKva: [2000], mccFeeders: STANDARD_MCC_FEEDERS };
  // 28 installed per board, of which 4 x 63 + 4 x 32 + 4 x 16 + 6 x 6 = 480 A are counted = 360.0 kVA
  const perBoardKva = (SQ3 * 433 * (4 * 63 + 4 * 32 + 4 * 16 + 6 * 6)) / 1000;
  const bigRows = () => mk([{ type: "DOL", kw: 185 }, { type: "DOL", kw: 185 }, { type: "VFD", kw: 250 }]);

  it("is off unless asked for", () => {
    expect(DEFAULT_OPTIONS.mccFeeders).toEqual([]);
    const r = sizeTransformers(bigRows(), 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000] });
    expect(r.combos[0].mccLoads).toEqual([]);
  });

  it("adds 28 feeders per set, counted in the combination's total (480 A = 360.0 kVA counted per board)", () => {
    const r = sizeTransformers(bigRows(), 433, MCC);
    const c = r.combos[0];
    expect(c.layout).toEqual([3]);
    expect(c.mccLoads).toHaveLength(28);
    expect(c.mccLoads.every((l) => l.source === "mcc" && l.rowNumber === 0 && l.kind === "feeder")).toBe(true);
    expect(c.mccLoads.reduce((a, l) => a + l.kva, 0)).toBeCloseTo(perBoardKva, 6);
    expect(perBoardKva).toBeCloseTo(360.0, 1);
    expect(c.mccLoads.filter((l) => l.kva === 0)).toHaveLength(28 - 18); // 10 installed but not counted
    expect(c.totalKva).toBeCloseTo(r.totalKva + perBoardKva, 6);
    expect(c.sets[0].kva).toBeCloseTo(c.totalKva, 6);
    expect(r.totalKva).toBeCloseTo(r.loads.reduce((a, l) => a + l.kva, 0), 6); // the list total excludes them
  });

  it("set of 3: half of each kind on each end section, none on the middle one", () => {
    const c = sizeTransformers(bigRows(), 433, MCC).combos[0];
    const count = (si: number, amps: number) => c.sections[si].loads.filter((l) => l.source === "mcc" && l.amps === amps).length;
    for (const [amps, half] of [[63, 3], [32, 3], [16, 3], [6, 5]] as const) {
      expect([count(0, amps), count(1, amps), count(2, amps)]).toEqual([half, 0, half]);
    }
  });

  it("set of 2: half of each kind on each side of the bus coupler", () => {
    const c = sizeTransformers(bigRows(), 433, { ...MCC, counts: [2] }).combos[0];
    expect(c.layout).toEqual([2]);
    const count = (si: number, amps: number) => c.sections[si].loads.filter((l) => l.source === "mcc" && l.amps === amps).length;
    for (const [amps, half] of [[63, 3], [32, 3], [16, 3], [6, 5]] as const) expect([count(0, amps), count(1, amps)]).toEqual([half, half]);
  });

  it("every set of a combination gets its own board of feeders (3 + 2 = 56)", () => {
    const rows = mk(Array.from({ length: 12 }, () => ({ type: "Feeder", amps: 1500 })));
    const c = sizeTransformers(rows, 433, { ...MCC, maxTransformers: 5, counts: [5] }).combos.find((x) => x.count === 5)!;
    expect(c.layout).toEqual([3, 2]);
    expect(c.mccLoads).toHaveLength(56);
    expect(c.totalKva).toBeCloseTo(12 * ((SQ3 * 433 * 1500) / 1000) + 2 * perBoardKva, 6);
    for (const set of c.sets) {
      expect(set.sections.flatMap((s) => s.loads).filter((l) => l.source === "mcc")).toHaveLength(28);
    }
  });

  it("can tip a list that just fits into needing more transformers", () => {
    // 3 x 2000 kVA carry at most 4,800 kVA (3 x 1600 normal, 2 x 1800 one out = 3,600 -> limit 3,600)
    const cap = setCapacityKva(3, 2000);
    const rows = mk([{ type: "Feeder", amps: Math.floor(((cap - 100) * 1000) / (SQ3 * 433) / 3) }, { type: "Feeder", amps: Math.floor(((cap - 100) * 1000) / (SQ3 * 433) / 3) }, { type: "Feeder", amps: Math.floor(((cap - 100) * 1000) / (SQ3 * 433) / 3) }]);
    const without = sizeTransformers(rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000] });
    const withMcc = sizeTransformers(rows, 433, MCC);
    expect(without.combos[0].fits).toBe(true);
    expect(withMcc.combos[0].fits).toBe(false);
  });
});

describe("current of each load (kVA ÷ (√3 × V))", () => {
  it("hand-calculated: 185 kW DOL = 370.9 A, 250 kW VFD = 389.9 A at 433 V; 450 kW DOL = 566.2 A at 690 V", () => {
    expect(loadOf({ type: "DOL", kw: 185 }).current).toBeCloseTo(370.9, 1); // 185 / (0.7 x 0.95) = 278.2 kVA
    expect(loadOf({ type: "VFD", kw: 250 }).current).toBeCloseTo(389.9, 1); // 250 / (0.9 x 0.95) = 292.4 kVA
    expect(loadOf({ type: "DOL", kw: 450, voltageV: 690 }, 690).current).toBeCloseTo(566.2, 1);
  });

  it("an amps-only load gets its own amps back, whatever the type", () => {
    for (const type of ["DOL", "VFD", "Feeder"]) expect(loadOf({ type, amps: 180 }).current).toBeCloseTo(180, 9);
    expect(loadOf({ type: "Feeder", amps: 800, voltageV: 690 }, 690).current).toBeCloseTo(800, 9);
  });

  it("the PMCC/IMCC feeders show their rating when counted and 0 when not", () => {
    const c = sizeTransformers(mkRows(), 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000], mccFeeders: STANDARD_MCC_FEEDERS }).combos[0];
    const counted = c.mccLoads.filter((l) => l.kva > 0);
    expect(counted).toHaveLength(18);
    for (const l of counted) expect(l.current).toBeCloseTo(l.amps!, 9);
    expect(c.mccLoads.filter((l) => l.kva === 0).every((l) => l.current === 0)).toBe(true);
  });
});
function mkRows(): RawLoadRow[] {
  return [{ rowNumber: 2, description: "a", type: "DOL", kw: 185, voltageV: 433 }];
}
