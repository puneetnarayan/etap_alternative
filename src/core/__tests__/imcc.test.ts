import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_MCC_ASSUMPTIONS, DEFAULT_RULE, DEFAULT_OPTIONS, MAX_INCOMER_A, NORMAL_STATE, appendFeedersToSheet, designImccs, imccFeederRows, loadResult, newId,
  normalizeMethod, normalizeStatus, parseImccSheet, parseLoadSheet, sizeTransformers, validateOperatingState, type ImccLoad,
} from "..";
import { imccWorkbookBuffer } from "../imcc-workbook";

const SQ3 = Math.sqrt(3);
// Hand formulas (docs/imcc.md): I = kW x 1000 / (sqrt3 x 433 x pf x eff), pf VFD 0.9, DOL 0.7 (eff 0.93), feeder 0.7.
const iVfd = (kw: number) => (kw * 1000) / (SQ3 * 433 * 0.9);
const iDol = (kw: number) => (kw * 1000) / (SQ3 * 433 * 0.7 * 0.93);
const kvaOf = (a: number) => (SQ3 * 433 * a) / 1000;

const L = (description: string, status: string, method: ImccLoad["method"], v: { kw?: number; a?: number; group?: string } = {}): ImccLoad => ({
  id: newId("load"), description, kw: v.kw, ratedA: v.a, status: normalizeStatus(status), statusText: status, method, group: v.group,
});
/** 300 A working feeders: design 375 A each, so two per section (750 A); a third would make 1125 A > 1000 A. */
const fdr = (n: number, status = "W") => L(`Feeder ${n}`, status, "PWR FDR", { a: 300 });

describe("the design script: working loads fill incomers up to 1000 A of design current", () => {
  it("7 working loads of 300 A (design 375 A) -> sections of 2, 2, 2, 1 -> 2 IMCCs; no incomer above 1000 A", () => {
    const set = designImccs(Array.from({ length: 7 }, (_, n) => fdr(n + 1)));
    expect(set.imccs.map((i) => i.tag)).toEqual(["IMCC-01", "IMCC-02"]);
    const secs = set.imccs.flatMap((i) => i.sections);
    expect(secs.map((s) => s.working)).toEqual([2, 2, 2, 1]);
    expect(secs.map((s) => s.workingA)).toEqual([600, 600, 600, 300]);
    expect(secs.map((s) => s.designA)).toEqual([750, 750, 750, 375]);
    for (const s of secs) expect(s.designA).toBeLessThanOrEqual(MAX_INCOMER_A);
    expect(set.imccs[0].incomer1).toMatchObject({ kind: "EDO ACB", ratingA: 1000 }); // 750 A -> 1000 A (existing ladder)
  });
  it("the number of IMCCs changes with the load: 4 loads -> 1 IMCC, 9 loads -> 3 IMCCs, 5 loads of 300 A in order", () => {
    expect(designImccs(Array.from({ length: 4 }, (_, n) => fdr(n))).imccs).toHaveLength(1);
    expect(designImccs(Array.from({ length: 5 }, (_, n) => fdr(n))).imccs).toHaveLength(2); // 2+2 | 1
    expect(designImccs(Array.from({ length: 9 }, (_, n) => fdr(n))).imccs).toHaveLength(3); // 2+2 | 2+2 | 1
    expect(designImccs([]).imccs).toHaveLength(0);
  });
  it("uses exactly 1000 A as a boundary: 800 A of working current (design 1000 A) still fits one incomer", () => {
    const set = designImccs([L("A", "W", "PWR FDR", { a: 400 }), L("B", "W", "PWR FDR", { a: 400 }), L("C", "W", "PWR FDR", { a: 1 })]);
    expect(set.imccs[0].sections[0].working).toBe(2);
    expect(set.imccs[0].sections[1].working).toBe(1);
    expect(set.imccs[0].sections[0].designA).toBeCloseTo(1000, 9);
  });
  it("a single load above the limit is flagged and placed alone", () => {
    const set = designImccs([fdr(1), L("Huge", "W", "PWR FDR", { a: 900 })]);
    expect(set.issues.some((x) => x.severity === "error" && /exceeds the 1000 A/.test(x.message))).toBe(true);
    expect(set.imccs[0].sections[0].loads.map((x) => x.load.description)).toEqual(["Feeder 1"]);
    expect(set.imccs[0].sections[1].loads.map((x) => x.load.description)).toEqual(["Huge"]);
  });
  it("only W / w loads drive the design; S and Spare do not change any rating", () => {
    const base = [fdr(1), fdr(2), fdr(3)];
    const withExtra = [...base, L("Stby", "S", "PWR FDR", { a: 300 }), L("Spare", "Spare", "PWR FDR", { a: 300 }), L("Lower w", "w", "PWR FDR", { a: 1 })];
    const a = designImccs(base), b = designImccs(withExtra);
    expect(b.imccs[0].sections.map((s) => s.working)).toEqual([2, 2]); // 'w' counted as working, in order
    expect(b.imccs[0].sections[0].workingA).toBe(a.imccs[0].sections[0].workingA);
  });
});

describe("status classification (W, w, S, Spare)", () => {
  it("W and w equal; S standby; Spare -> SP; others invalid", () => {
    expect(["W", "w", "S", "s", "Spare", "spare", "SP", "x", ""].map(normalizeStatus)).toEqual(["W", "W", "S", "S", "SP", "SP", "SP", undefined, undefined]);
  });
  it("standby and spare rows stay with their group's section, are counted as such, and carry no working load", () => {
    const loads = [
      L("P1 duty", "W", "PWR FDR", { a: 300, group: "G" }), L("P2 duty", "W", "PWR FDR", { a: 300, group: "H" }),
      L("P3 duty", "W", "PWR FDR", { a: 300, group: "I" }), // new section (3rd 300 A would exceed 1000 A... design 1125)
      L("P1 standby", "S", "PWR FDR", { a: 300, group: "G" }), L("Spare feeder", "Spare", "PWR FDR", { a: 300 }),
    ];
    const set = designImccs(loads);
    const [s1, s2] = set.imccs[0].sections;
    expect(s1.loads.map((x) => x.load.description)).toEqual(["P1 duty", "P2 duty", "P1 standby"]);
    expect(s2.loads.map((x) => x.load.description)).toEqual(["P3 duty", "Spare feeder"]); // no group: with the preceding working load
    expect([s1.working, s1.standby, s1.spare, s2.spare]).toEqual([2, 1, 0, 1]);
    expect(s1.workingA).toBe(600);
    expect(s1.standbyKva).toBeCloseTo(kvaOf(300), 9);
  });
  it("a missing or invalid status is reported and the row is not counted or placed", () => {
    const set = designImccs([fdr(1), L("Mystery", "maybe", "PWR FDR", { a: 300 }), L("Blank", "", "PWR FDR", { a: 300 })]);
    expect(set.issues.filter((x) => x.severity === "error" && /status/.test(x.message))).toHaveLength(2);
    expect(set.imccs[0].sections[0].workingA).toBe(300);
    expect(set.imccs[0].sections[0].loads).toHaveLength(1);
  });
  it("no working load at all: nothing to design", () => {
    const set = designImccs([L("S only", "S", "DOL", { kw: 30 })]);
    expect(set.imccs).toEqual([]);
    expect(set.issues[0].message).toMatch(/No working/);
  });
});

describe("currents follow the application's conventions; SS = soft starter", () => {
  it("VFD 0.9, DOL 0.7 x 0.93, soft starter calculated like DOL, hand-checked", () => {
    expect(normalizeMethod("SS")).toBe("SS");
    expect(normalizeMethod("Soft starter")).toBe("SS");
    expect(loadResult(L("v", "W", "VFD", { kw: 55 }), 433).usedA).toBeCloseTo(iVfd(55), 9);
    expect(loadResult(L("d", "W", "DOL", { kw: 30 }), 433).usedA).toBeCloseTo(iDol(30), 9);
    expect(loadResult(L("s", "W", "SS", { kw: 30 }), 433).usedA).toBeCloseTo(iDol(30), 9);
    expect(loadResult(L("s", "W", "SS", { kw: 30 }), 433).kva).toBeCloseTo(kvaOf(iDol(30)), 9);
  });
  it("section sums and design current, hand-calculated", () => {
    const set = designImccs([L("v", "W", "VFD", { kw: 55 }), L("d", "W", "DOL", { kw: 30 }), L("s", "W", "SS", { kw: 75 })]);
    const s = set.imccs[0].sections[0];
    expect(s.workingA).toBeCloseTo(iVfd(55) + iDol(30) + iDol(75), 9);
    expect(s.designA).toBeCloseTo((iVfd(55) + iDol(30) + iDol(75)) * 1.25, 9);
    expect(s.workingKva).toBeCloseTo(kvaOf(s.workingA), 9);
    expect(s.lceKva).toBeCloseTo(s.workingKva * 0.8, 9);
  });
});

describe("Incomer 2 and the bus coupler follow Incomer 1", () => {
  it("always equal to Incomer 1's calculated rating, also after the loads change", () => {
    for (const n of [3, 4, 6]) {
      const set = designImccs(Array.from({ length: n }, (_, k) => fdr(k)));
      for (const i of set.imccs) { expect(i.incomer2.ratingA).toBe(i.incomer1.ratingA); expect(i.coupler.ratingA).toBe(i.incomer1.ratingA); }
    }
  });
  it("warns when section 2 would need more than Incomer 1", () => {
    // 100 A (design 125 A) + 790 A (design 987.5 A) do not fit one incomer (1112.5 A): section 2 starts with the big load
    const set = designImccs([L("small", "W", "PWR FDR", { a: 100 }), L("big", "W", "PWR FDR", { a: 790 })]);
    const i = set.imccs[0];
    expect(i.sections[0].working).toBe(1);
    expect(i.incomer1).toMatchObject({ kind: "MCCB", ratingA: 125 }); // 125 A design
    expect(i.sections[1].breaker).toMatchObject({ kind: "EDO ACB", ratingA: 1000 });
    expect(i.incomer2.ratingA).toBe(125); // still follows Incomer 1
    expect(i.issues.some((x) => /Incomer 2 section needs/.test(x.message))).toBe(true);
  });
});

describe("operating-state validation", () => {
  it("permitted: one incomer ON, the other OFF, coupler OFF", () => {
    expect(validateOperatingState(NORMAL_STATE).ok).toBe(true);
    expect(validateOperatingState({ incomer1: "OFF", incomer2: "ON", coupler: "OFF" }).ok).toBe(true);
  });
  it("prohibited: both ON, one ON with coupler ON, none ON, coupler ON", () => {
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "OFF" }).errors.join(" ")).toMatch(/Both incomers are ON/);
    expect(validateOperatingState({ incomer1: "ON", incomer2: "OFF", coupler: "ON" }).errors.join(" ")).toMatch(/not assumed to supply both/);
    expect(validateOperatingState({ incomer1: "OFF", incomer2: "OFF", coupler: "OFF" }).ok).toBe(false);
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "ON" }).errors.length).toBeGreaterThanOrEqual(2);
  });
  it("changing the philosophy needs a named authoriser", () => {
    const loose = { exactlyOneIncomerOn: false, couplerMustBeOff: false };
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "OFF" }, loose).errors.join(" ")).toMatch(/engineering change/);
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "OFF" }, { ...loose, authorisedBy: "J. Doe" }).ok).toBe(true);
    expect(DEFAULT_RULE).toEqual({ exactlyOneIncomerOn: true, couplerMustBeOff: true });
  });
});

describe("the Automatic PMCC feeders switch: generated incomer feeders", () => {
  const set = () => designImccs(Array.from({ length: 5 }, (_, n) => fdr(n + 1))); // 2 IMCCs: sections 2,2 | 1 (IMCC-02 has no section 2)
  it("one feeder per incomer with a working load: Incomer 1 = W, Incomer 2 = S, group = IMCC tag, current = section working current", () => {
    const rows = imccFeederRows(set());
    expect(rows.map((r) => [r.description, r.status, r.group, r.amps])).toEqual([
      ["IMCC-01 Incomer 1", "W", "IMCC-01", 600], ["IMCC-01 Incomer 2", "S", "IMCC-01", 600], ["IMCC-02 Incomer 1", "W", "IMCC-02", 300],
    ]);
    expect(rows[0].kva).toBeCloseTo(kvaOf(600), 9);
    expect(rows[0].voltageV).toBe(433);
  });
  it("appends the rows to a PMCC list, originals unchanged; no rows -> unchanged; running twice from the raw list gives no duplicates", () => {
    const sheet = [["Description", "W/S", "Type (DOL/VFD/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)", "Duty/standby group"], ["Pump", "W", "DOL", null, 90, 433, null]];
    const rows = imccFeederRows(set());
    const out = appendFeedersToSheet(sheet, rows);
    expect(out.slice(0, 2)).toEqual(sheet);
    expect(out).toHaveLength(2 + 3);
    expect(out[2]).toEqual(["IMCC-01 Incomer 1", "W", "Feeder", 600, null, 433, "IMCC-01"]);
    expect(out[3][1]).toBe("S");
    expect(appendFeedersToSheet(sheet, [])).toEqual(sheet);
    expect(appendFeedersToSheet(sheet, rows)).toEqual(out); // pure: same input, same output
    const parsed = parseLoadSheet(out as never);
    expect(parsed.errors).toEqual([]);
    expect(parsed.rows.filter((r) => r.group === "IMCC-01")).toHaveLength(2);
  });
  it("adds W/S and group columns when the PMCC list has none", () => {
    const sheet = [["Description", "Type", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)"], ["Pump", "DOL", null, 90, 433]];
    const out = appendFeedersToSheet(sheet, imccFeederRows(set()));
    expect(out[0]).toEqual(["Description", "Type", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)", "W/S", "Duty/standby group"]);
    expect(out[2]).toEqual(["IMCC-01 Incomer 1", "Feeder", 600, null, 433, "W", "IMCC-01"]);
  });
  it("switch ON: the transformer sizing is recalculated with the incomers (more kVA, number of transformers may change); the incomers of one IMCC go on different sections", () => {
    // 10 feeders of 467 A at 433 V = 350.2 kVA each -> 3502 kVA: fits 3 x 2000 kVA (3600 kVA)
    const base = [["Description", "W/S", "Type (DOL/VFD/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)", "Duty/standby group"],
      ...Array.from({ length: 10 }, (_, n) => [`Fdr ${n + 1}`, "W", "Feeder", 467, null, 433, null])];
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const off = sizeTransformers(parseLoadSheet(base as never).rows, 433, opts);
    expect(off.recommended?.count).toBe(3);
    // an IMCC set whose two incomers carry 800 A (600 kVA) each
    const big = designImccs([L("a", "W", "PWR FDR", { a: 400 }), L("b", "W", "PWR FDR", { a: 400 }), L("c", "W", "PWR FDR", { a: 400 }), L("d", "W", "PWR FDR", { a: 400 })]);
    const rows = imccFeederRows(big);
    expect(rows.map((r) => r.amps)).toEqual([800, 800]);
    const onSheet = appendFeedersToSheet(base, rows);
    const on = sizeTransformers(parseLoadSheet(onSheet as never).rows, 433, opts);
    expect(on.totalKva).toBeCloseTo(off.totalKva + 2 * kvaOf(800), 6);
    expect(on.recommended!.count).toBeGreaterThan(3); // 3502 + 1200 = 4702 kVA > 3600 kVA
    const placed = on.combos.find((c) => c.count === on.recommended!.count)!.sections;
    const sectionOfInc = placed.map((s, n) => ({ n, names: s.loads.map((l) => l.description).filter((d) => d.startsWith("IMCC-01")) }));
    expect(sectionOfInc.filter((x) => x.names.length > 0)).toHaveLength(2); // different sections
    // switch OFF: the list is exactly the uploaded one, so the sizing is the same as before
    expect(sizeTransformers(parseLoadSheet(base as never).rows, 433, opts).totalKva).toBe(off.totalKva);
  });
});

describe("importing the reference layout", () => {
  const sheet = [
    [null, "ECR-X MCC FEEDERS"],
    [null, "S. NO.", null, null, "KW", "Current in Amps", "W/S/Spare", "Method of starting ", "KVA", "Group", "Working KVA"],
    [null, "INCOMER- 1"],
    [null, 1, "Linear screen - 1", null, 5.5, 8.5, "W", "VFD", 6.1, "Linear screen operation"],
    [null, 2, "Linear screen - Spare", null, 5.5, 8.5, "Spare", "VFD"],
    [null, 3, "Spillage pump", "Rating changed from 5.5kW to 30kW", 30, 60, "w", "SS"],
    [null, null, null, null, null, null, null, null, 6.1],
    [null, "INCOMER- 2"],
    [null, 4, "Pump 2", null, 15, null, "S", "DOL"],
    [null, 5, "1-ph 6A Feeder-1"],
  ];
  it("reads statuses (W, w, Spare, S, missing), methods incl. SS, groups, notes; skips totals and incomer rows; ignores the reference's current", () => {
    const r = parseImccSheet(sheet);
    expect(r.errors).toEqual([]);
    expect(r.incomerRows).toBe(2);
    expect(r.loads.map((l) => [l.description, l.kw, l.status, l.method])).toEqual([
      ["Linear screen - 1", 5.5, "W", "VFD"], ["Linear screen - Spare", 5.5, "SP", "VFD"], ["Spillage pump", 30, "W", "SS"],
      ["Pump 2", 15, "S", "DOL"], ["1-ph 6A Feeder-1", undefined, undefined, undefined],
    ]);
    expect(r.loads[0].group).toBe("Linear screen operation");
    expect(r.loads[2].note).toBe("Rating changed from 5.5kW to 30kW");
    expect(r.loads.every((l) => l.ratedA === undefined)).toBe(true);
  });
  it("reports a sheet without the expected header", () => {
    expect(parseImccSheet([["a", "b"], [1, 2]]).errors[0]).toMatch(/S\. NO\./);
  });
});

describe("Excel output", () => {
  const loads = () => [
    L("Screen", "W", "VFD", { kw: 55 }), L("Pump", "w", "DOL", { kw: 30 }), L("Spare pump", "Spare", "DOL", { kw: 30 }), L("Stby pump", "S", "SS", { kw: 30 }),
    ...Array.from({ length: 4 }, (_, n) => fdr(n + 1)),
  ];
  const argbOf = (c: ExcelJS.Cell) => (c.fill as ExcelJS.FillPattern | undefined)?.fgColor?.argb;
  it("Summary, one sheet per IMCC and Notes; bus layout in order with the app's colours; spare as SP; totals as formulas; Incomer 2 and coupler equal Incomer 1", async () => {
    const set = designImccs(loads());
    for (const linked of [true, false]) {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await imccWorkbookBuffer(set, DEFAULT_MCC_ASSUMPTIONS, { sourceName: "x.xlsx", linked }));
      expect(wb.worksheets.map((w) => w.name)).toEqual(["Summary", ...set.imccs.map((i) => i.tag), "Notes"]);
      const ws = wb.getWorksheet("IMCC-01")!;
      const texts: string[] = [];
      ws.eachRow((r) => r.eachCell((c) => { if (typeof c.value === "string") texts.push(c.value); }));
      expect(texts.some((t) => t.includes(linked ? 'PMCC outgoing feeder "IMCC-01 Incomer 1"' : "switch OFF"))).toBe(true);
      // order along the bus: INC 1, loads, BC, loads, INC 2
      const marks: string[] = [];
      ws.eachRow((r) => { const v = r.getCell(1).value; if (v === "INC 1" || v === "BC" || v === "INC 2") marks.push(String(v)); });
      expect(marks).toEqual(["INC 1", "BC", "INC 2"]);
      const rowOf = (m: string) => { let n = 0; ws.eachRow((r) => { if (r.getCell(1).value === m) n = r.number; }); return n; };
      expect(rowOf("INC 1")).toBeLessThan(rowOf("BC"));
      expect(rowOf("BC")).toBeLessThan(rowOf("INC 2"));
      // colours as in the other reports
      expect(argbOf(ws.getRow(rowOf("INC 1")).getCell(3))).toBe("FF80DEEA"); // incomer cyan
      expect(argbOf(ws.getRow(rowOf("BC")).getCell(3))).toBe("FFFFF9C4"); // coupler light yellow
      const fills: Record<string, string | undefined> = {};
      ws.eachRow((r) => { const v = r.getCell(5).value; if (typeof v === "string" && ["W", "S", "SP"].includes(v)) fills[v] = argbOf(r.getCell(5)); });
      expect(fills).toEqual({ W: "FFDEB1FB", S: "FFFBB1DE", SP: "FFFFD966" });
      const sums: string[] = [];
      ws.eachRow((r) => { const v = r.getCell(11).value as { formula?: string } | null; if (r.number > 17 && v && typeof v === "object" && v.formula?.startsWith("SUM(")) sums.push(String(r.getCell(1).value)); });
      expect(sums).toEqual(["INC 1", "INC 2"]);
      const ratingA = [rowOf("INC 1"), rowOf("BC"), rowOf("INC 2")].map((r) => ws.getRow(r).getCell(17).value);
      expect(new Set(ratingA).size).toBe(1);
      expect(ratingA[0]).toBe(set.imccs[0].incomer1.ratingA);
      const status: string[] = [];
      ws.eachRow((r) => { const v = r.getCell(5).value; if (typeof v === "string" && v.length <= 2 && r.number > 19) status.push(v); });
      expect(status.filter((x) => x === "SP")).toHaveLength(1);
    }
  });
});
