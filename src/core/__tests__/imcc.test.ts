import { readFileSync } from "node:fs";
import ExcelJS from "exceljs";
import readExcelFile from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_MCC_ASSUMPTIONS, DEFAULT_OPTIONS, DEFAULT_RULE, FAILOVER_STATE, INCOMER_ACB_OPTIONS, NORMAL_STATE, appendFeedersToSheet, classifyFile, designImccs, imccFeederRows,
  loadResult, newId, normalizeMethod, normalizeStatus, parseImccLoadList, parseLoadSheet, planIntake, sizeTransformers, switchToFailover, switchToNormal,
  validateOperatingState, type ImccLoad,
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
/** n working feeders of `a` A (design a x 1.25). */
const many = (n: number, a = 300) => Array.from({ length: n }, (_, k) => L(`Feeder ${k + 1}`, "W", "Feeder", { a }));

describe("the design script: working loads fill an IMCC up to ACB size x maximum loading of design current", () => {
  it("1000 A ACB at 90 % = 900 A: 300 A feeders (design 375 A) go two per IMCC (750 A); 7 loads -> 4 IMCCs", () => {
    const set = designImccs(many(7));
    expect(set.maxDesignA).toBeCloseTo(900, 9);
    expect(set.imccs.map((i) => i.tag)).toEqual(["IMCC-01", "IMCC-02", "IMCC-03", "IMCC-04"]);
    expect(set.imccs.map((i) => i.bus.working)).toEqual([2, 2, 2, 1]);
    expect(set.imccs.map((i) => i.bus.workingA)).toEqual([600, 600, 600, 300]);
    expect(set.imccs.map((i) => i.bus.designA)).toEqual([750, 750, 750, 375]);
    for (const i of set.imccs) expect(i.bus.designA).toBeLessThanOrEqual(900);
  });
  it("the number of IMCCs changes with the load", () => {
    expect([2, 3, 4, 5, 9].map((n) => designImccs(many(n)).imccs.length)).toEqual([1, 2, 2, 3, 5]);
    expect(designImccs([]).imccs).toHaveLength(0);
  });
  it("the boundary is inclusive: 720 A working = 900 A design fits; one more ampere starts a new IMCC", () => {
    const fits = designImccs([L("A", "W", "Feeder", { a: 400 }), L("B", "W", "Feeder", { a: 320 })]);
    expect(fits.imccs).toHaveLength(1);
    expect(fits.imccs[0].bus.designA).toBeCloseTo(900, 9);
    expect(designImccs([L("A", "W", "Feeder", { a: 400 }), L("B", "W", "Feeder", { a: 320 }), L("C", "W", "Feeder", { a: 1 })]).imccs).toHaveLength(2);
  });
  it("the incomer ACB is selectable (630, 800, 1000, 1250 A; default 1000) and sets the limit and the rating of all three devices", () => {
    expect([...INCOMER_ACB_OPTIONS]).toEqual([630, 800, 1000, 1250]);
    const byAcb = (a: number) => designImccs(many(5), DEFAULT_MCC_ASSUMPTIONS, { maxIncomerA: a });
    // limit = ACB x 0.9 / 1.25 of working current: 630 -> 453.6 A (1 feeder), 800 -> 576 A (1), 1000 -> 720 A (2), 1250 -> 900 A (3)
    expect([630, 800, 1000, 1250].map((a) => byAcb(a).imccs.length)).toEqual([5, 5, 3, 2]);
    for (const a of [630, 800, 1000, 1250]) {
      const set = byAcb(a);
      expect(set.maxIncomerA).toBe(a);
      for (const i of set.imccs) {
        expect([i.incomer1.ratingA, i.incomer2.ratingA, i.coupler.ratingA]).toEqual([a, a, a]);
        expect(i.incomer1.kind).toBe("EDO ACB");
        expect(i.bus.designA).toBeLessThanOrEqual(a * 0.9 + 1e-9);
        expect(i.loading).toBeCloseTo(i.bus.designA / a, 9);
      }
    }
    expect(designImccs(many(5)).maxIncomerA).toBe(1000); // default
  });
  it("the ACB maximum loading is adjustable: 100 % = 1000 A, 70 % = 700 A", () => {
    expect(designImccs(many(4), DEFAULT_MCC_ASSUMPTIONS, { acbMaxLoading: 1 }).imccs.map((i) => i.bus.working)).toEqual([2, 2]); // 750 A, a third would be 1125 A
    const tight = designImccs(many(3), DEFAULT_MCC_ASSUMPTIONS, { acbMaxLoading: 0.7 });
    expect(tight.maxDesignA).toBeCloseTo(700, 9);
    expect(tight.imccs.map((i) => i.bus.working)).toEqual([1, 1, 1]); // 750 A already exceeds 700 A
    expect(tight.imccs[0].loadingLimit).toBe(0.7);
    expect(designImccs(many(3), DEFAULT_MCC_ASSUMPTIONS, { acbMaxLoading: 1 }).imccs[0].loadingLimit).toBe(1);
  });
  it("a single load above the limit is flagged and placed alone", () => {
    const set = designImccs([...many(1), L("Huge", "W", "Feeder", { a: 800 })]);
    expect(set.issues.some((x) => x.severity === "error" && /Huge.*exceeds 900 A/.test(x.message))).toBe(true);
    expect(set.imccs.map((i) => i.bus.loads.map((x) => x.load.description))).toEqual([["Feeder 1"], ["Huge"]]);
    expect(set.imccs[1].issues.some((x) => /engineering review/.test(x.message))).toBe(true);
  });
  it("only W / w loads drive the design; S and Spare change no current", () => {
    const base = many(3);
    const extra = [...base, L("Stby", "S", "Feeder", { a: 300 }), L("Spare", "Spare", "Feeder", { a: 300 })];
    const a = designImccs(base), b = designImccs(extra);
    expect(b.imccs[0].bus.workingA).toBe(a.imccs[0].bus.workingA);
    expect(b.imccs.map((i) => i.bus.working)).toEqual(a.imccs.map((i) => i.bus.working));
  });
});

describe("status classification (W, w, S, Spare)", () => {
  it("W and w equal; S standby; Spare -> SP; others invalid", () => {
    expect(["W", "w", "S", "s", "Spare", "spare", "SP", "x", ""].map(normalizeStatus)).toEqual(["W", "W", "S", "S", "SP", "SP", "SP", undefined, undefined]);
  });
  it("standby and spare rows stay with their group's IMCC, else the preceding working load; they carry no working load", () => {
    const loads = [
      L("P1 duty", "W", "Feeder", { a: 300, group: "G" }), L("P2 duty", "W", "Feeder", { a: 300, group: "H" }),
      L("P3 duty", "W", "Feeder", { a: 300, group: "I" }), // 3rd 300 A would make 1125 A: second IMCC
      L("P1 standby", "S", "Feeder", { a: 300, group: "G" }), L("Spare feeder", "Spare", "Feeder", { a: 300 }),
    ];
    const [i1, i2] = designImccs(loads).imccs;
    expect(i1.bus.loads.map((x) => x.load.description)).toEqual(["P1 duty", "P2 duty", "P1 standby"]);
    expect(i2.bus.loads.map((x) => x.load.description)).toEqual(["P3 duty", "Spare feeder"]);
    expect([i1.bus.working, i1.bus.standby, i1.bus.spare, i2.bus.spare]).toEqual([2, 1, 0, 1]);
    expect(i1.bus.workingA).toBe(600);
    expect(i1.bus.standbyKva).toBeCloseTo(kvaOf(300), 9);
  });
  it("a missing or invalid status is reported and the row is not counted or placed", () => {
    const set = designImccs([...many(1), L("Mystery", "maybe", "Feeder", { a: 300 }), L("Blank", "", "Feeder", { a: 300 })]);
    expect(set.issues.filter((x) => x.severity === "error" && /status/.test(x.message))).toHaveLength(2);
    expect(set.imccs[0].bus.workingA).toBe(300);
    expect(set.imccs[0].bus.loads).toHaveLength(1);
  });
  it("no working load at all: nothing to design", () => {
    const set = designImccs([L("S only", "S", "DOL", { kw: 30 })]);
    expect(set.imccs).toEqual([]);
    expect(set.issues[0].message).toMatch(/No working/);
  });
});

describe("currents follow the application's conventions; SS = soft starter; Feeder replaces PWR FDR", () => {
  it("VFD 0.9, DOL 0.7 x 0.93, soft starter like DOL, hand-checked", () => {
    expect(normalizeMethod("SS")).toBe("SS");
    expect(normalizeMethod("Soft starter")).toBe("SS");
    expect(loadResult(L("v", "W", "VFD", { kw: 55 }), 433).usedA).toBeCloseTo(iVfd(55), 9);
    expect(loadResult(L("d", "W", "DOL", { kw: 30 }), 433).usedA).toBeCloseTo(iDol(30), 9);
    expect(loadResult(L("s", "W", "SS", { kw: 30 }), 433).usedA).toBeCloseTo(iDol(30), 9);
    expect(loadResult(L("s", "W", "SS", { kw: 30 }), 433).kva).toBeCloseTo(kvaOf(iDol(30)), 9);
  });
  it("PWR FDR (and power feeder) is read as Feeder", () => {
    expect(["Feeder", "feeder", "PWR FDR", "pwr fdr", "Power feeder"].map(normalizeMethod)).toEqual(["Feeder", "Feeder", "Feeder", "Feeder", "Feeder"]);
  });
  it("IMCC sums and design current, hand-calculated", () => {
    const set = designImccs([L("v", "W", "VFD", { kw: 55 }), L("d", "W", "DOL", { kw: 30 }), L("s", "W", "SS", { kw: 75 })]);
    const b = set.imccs[0].bus;
    expect(b.workingA).toBeCloseTo(iVfd(55) + iDol(30) + iDol(75), 9);
    expect(b.designA).toBeCloseTo((iVfd(55) + iDol(30) + iDol(75)) * 1.25, 9);
    expect(b.workingKva).toBeCloseTo(kvaOf(b.workingA), 9);
    expect(b.lceKva).toBeCloseTo(b.workingKva * 0.8, 9);
  });
});

describe("operating states: normal, Incomer 1 unhealthy, interlock", () => {
  it("permitted: normal (Incomer 1 ON, Incomer 2 and bus coupler OFF) and unhealthy (Incomer 1 OFF, Incomer 2 and bus coupler ON)", () => {
    expect(validateOperatingState(NORMAL_STATE).ok).toBe(true);
    expect(validateOperatingState(FAILOVER_STATE).ok).toBe(true);
    expect(FAILOVER_STATE).toEqual({ incomer1: "OFF", incomer2: "ON", coupler: "ON" });
  });
  it("prohibited: Incomer 1 ON with Incomer 2 and / or the coupler ON (interlock), Incomer 2 or the coupler ON alone, nothing ON", () => {
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "OFF" }).errors.join(" ")).toMatch(/Interlock.*Incomer 2 must stay OFF/);
    expect(validateOperatingState({ incomer1: "ON", incomer2: "OFF", coupler: "ON" }).errors.join(" ")).toMatch(/Interlock.*bus coupler must stay OFF/);
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "ON" }).errors.join(" ")).toMatch(/Incomer 2 and the bus coupler must stay OFF/);
    expect(validateOperatingState({ incomer1: "OFF", incomer2: "ON", coupler: "OFF" }).errors.join(" ")).toMatch(/switched ON together/);
    expect(validateOperatingState({ incomer1: "OFF", incomer2: "OFF", coupler: "ON" }).errors.join(" ")).toMatch(/switched ON together/);
    expect(validateOperatingState({ incomer1: "OFF", incomer2: "OFF", coupler: "OFF" }).errors.join(" ")).toMatch(/No source is ON/);
  });
  it("interlock sequence: from normal, Incomer 1 is switched OFF first, then Incomer 2 and the coupler ON; and back in reverse", () => {
    const f = switchToFailover(NORMAL_STATE);
    expect(f.state).toEqual(FAILOVER_STATE);
    expect(f.steps).toEqual(["Interlock: Incomer 1 switched OFF", "Incomer 2 switched ON", "Bus coupler switched ON"]);
    expect(switchToFailover({ incomer1: "ON", incomer2: "ON", coupler: "ON" }).steps).toEqual(["Interlock: Incomer 1 switched OFF"]);
    expect(switchToFailover(FAILOVER_STATE).steps).toEqual(["Already in the Incomer 1 unhealthy state"]);
    const n = switchToNormal(FAILOVER_STATE);
    expect(n.state).toEqual(NORMAL_STATE);
    expect(n.steps).toEqual(["Incomer 2 switched OFF", "Bus coupler switched OFF", "Incomer 1 switched ON"]);
  });
  it("changing the philosophy needs a named authoriser", () => {
    const off = { failoverOnIncomer1Unhealthy: false };
    expect(validateOperatingState(NORMAL_STATE, off).errors.join(" ")).toMatch(/engineering change/);
    expect(validateOperatingState(NORMAL_STATE, { ...off, authorisedBy: "J. Doe" }).ok).toBe(true);
    expect(validateOperatingState(FAILOVER_STATE, { ...off, authorisedBy: "J. Doe" }).ok).toBe(false);
    expect(DEFAULT_RULE).toEqual({ failoverOnIncomer1Unhealthy: true });
  });
});

describe("the Automatic PMCC feeders switch: generated incomer feeders", () => {
  const set = () => designImccs(many(5)); // IMCCs of 2, 2, 1 loads
  it("two feeders per IMCC (one per incomer): Incomer 1 = W, Incomer 2 = S, both with the IMCC working current, group = IMCC tag", () => {
    const rows = imccFeederRows(set());
    expect(rows.map((r) => [r.description, r.status, r.group, r.amps])).toEqual([
      ["IMCC-01 Incomer 1", "W", "IMCC-01", 600], ["IMCC-01 Incomer 2", "S", "IMCC-01", 600],
      ["IMCC-02 Incomer 1", "W", "IMCC-02", 600], ["IMCC-02 Incomer 2", "S", "IMCC-02", 600],
      ["IMCC-03 Incomer 1", "W", "IMCC-03", 300], ["IMCC-03 Incomer 2", "S", "IMCC-03", 300],
    ]);
    expect(rows[0].kva).toBeCloseTo(kvaOf(600), 9);
    expect(rows[0].voltageV).toBe(433);
  });
  it("appends the rows to a PMCC list, originals unchanged; nothing to add -> unchanged; pure (no duplicates)", () => {
    const sheet = [["Description", "W/S", "Type (DOL/VFD/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)", "Duty/standby group"], ["Pump", "W", "DOL", null, 90, 433, null]];
    const rows = imccFeederRows(set());
    const out = appendFeedersToSheet(sheet, rows);
    expect(out.slice(0, 2)).toEqual(sheet);
    expect(out).toHaveLength(2 + 6);
    expect(out[2]).toEqual(["IMCC-01 Incomer 1", "W", "Feeder", 600, null, 433, "IMCC-01"]);
    expect(out[3][1]).toBe("S");
    expect(appendFeedersToSheet(sheet, [])).toEqual(sheet);
    expect(appendFeedersToSheet(sheet, rows)).toEqual(out);
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
  it("switch ON: the transformer sizing is recalculated with the working incomer feeder (number of transformers may change); the standby Incomer 2 feeder is not counted; OFF: unchanged", () => {
    const base = [["Description", "W/S", "Type (DOL/VFD/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)", "Duty/standby group"],
      ...Array.from({ length: 10 }, (_, n) => [`Fdr ${n + 1}`, "W", "Feeder", 467, null, 433, null])];
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const off = sizeTransformers(parseLoadSheet(base as never).rows, 433, opts);
    expect(off.recommended?.count).toBe(3); // 3502 kVA <= 3600 kVA
    const two = designImccs([L("a", "W", "Feeder", { a: 350 }), L("b", "W", "Feeder", { a: 350 })]); // design 875 A <= 900 A: one IMCC
    const rows = imccFeederRows(two);
    expect(rows.map((r) => [r.description, r.status, r.amps])).toEqual([["IMCC-01 Incomer 1", "W", 700], ["IMCC-01 Incomer 2", "S", 700]]);
    const on = sizeTransformers(parseLoadSheet(appendFeedersToSheet(base, rows) as never).rows, 433, opts);
    // only the working Incomer 1 feeder (700 A = 525 kVA) is counted; Incomer 2 (standby) is listed as excluded
    expect(on.totalKva).toBeCloseTo(off.totalKva + kvaOf(700), 6);
    expect(on.loads.some((l) => l.description === "IMCC-01 Incomer 1")).toBe(true);
    expect(on.loads.some((l) => l.description === "IMCC-01 Incomer 2")).toBe(false);
    expect(on.issues.some((i) => i.severity === "excluded" && i.description === "IMCC-01 Incomer 2" && /Standby \(S\)/.test(i.message))).toBe(true);
    expect(on.recommended!.count).toBeGreaterThan(3); // 3502 + 525 = 4027 kVA > 3600 kVA
    expect(sizeTransformers(parseLoadSheet(base as never).rows, 433, opts).totalKva).toBe(off.totalKva);
  });
});

describe("the IMCC input: same layout as the PMCC load list", () => {
  const sheet = [
    ["Description", "W/S", "Type (RDOL/DOL/VFD/SS/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)", "Duty/standby group"],
    ["Pump A", "w", "SS", 60, 30, 433, "G1"],
    ["Pump B", "Spare", "DOL", 60, 30, 433, null],
    ["Fan", "S", "VFD", null, 55, 433, null],
    ["Hoist", null, "Feeder", 59.7, 30, 433, null],
    ["Paste thickener", "W", "Feeder", "-", "-", 433, null],
    ["Valve", "W", "RDOL", null, 2.2, 690, null],
  ];
  it("reads W/w, S, Spare, a blank status, SS / RDOL / Feeder, amps and kW, group, voltage; bad numbers are flagged", () => {
    const r = parseImccLoadList(sheet);
    expect(r.errors).toEqual([]);
    expect(r.loads.map((l) => [l.description, l.status, l.method, l.ratedA, l.kw, l.voltageV])).toEqual([
      ["Pump A", "W", "SS", 60, 30, 433], ["Pump B", "SP", "DOL", 60, 30, 433], ["Fan", "S", "VFD", undefined, 55, 433],
      ["Hoist", undefined, "Feeder", 59.7, 30, 433], ["Paste thickener", "W", "Feeder", undefined, undefined, 433], ["Valve", "W", "RDOL", undefined, 2.2, 690],
    ]);
    expect(r.loads[0].group).toBe("G1");
    expect(r.loads[4].badFields).toEqual(["Amps", "kW"]);
  });
  it("a blank status is flagged and not counted; a row at another voltage is not placed; bad numbers are reported; entered amps beat kW", () => {
    const set = designImccs(parseImccLoadList(sheet).loads);
    const msgs = set.issues.map((x) => x.message).join(" | ");
    expect(msgs).toMatch(/"Hoist": status missing/);
    expect(msgs).toMatch(/"Valve": 690 V differs from the IMCC voltage 433 V/);
    expect(msgs).toMatch(/"Paste thickener": could not read a number in Amps, kW/);
    expect(set.imccs[0].bus.workingA).toBeCloseTo(60, 9);
  });
  it("reports a sheet without the expected columns", () => {
    expect(parseImccLoadList([["a", "b"], [1, 2]]).errors[0]).toMatch(/Missing column/);
  });
});

describe("the sample files served by the app", () => {
  const read = async (name: string) => {
    const f = readFileSync(new URL(`../../../public/${name}`, import.meta.url));
    return (await readExcelFile(f.buffer.slice(f.byteOffset, f.byteOffset + f.byteLength) as ArrayBuffer))[0].data as unknown[][];
  };
  it("IMCC_sample_load_list.xlsx: 51 rows (23 w, 1 S, 11 Spare, 16 blank); design from the working loads, hand-checked", async () => {
    const r = parseImccLoadList(await read("IMCC_sample_load_list.xlsx"));
    expect(r.errors).toEqual([]);
    expect(r.loads).toHaveLength(51);
    const count = (st: string | undefined) => r.loads.filter((l) => l.status === st).length;
    expect([count("W"), count("S"), count("SP"), count(undefined)]).toEqual([23, 1, 11, 16]);
    const workingA = r.loads.filter((l) => l.status === "W").map((l) => l.ratedA ?? 0);
    expect(workingA.reduce((a, b) => a + b, 0)).toBeCloseTo(1138.7802793889566, 9);
    for (const acb of [630, 800, 1000, 1250]) {
      const set = designImccs(r.loads, DEFAULT_MCC_ASSUMPTIONS, { maxIncomerA: acb });
      // greedy fill, independent: working current up to acb x 0.9 / 1.25
      const lim = (acb * 0.9) / 1.25;
      const split: number[][] = [[]];
      for (const a of workingA) { const cur = split[split.length - 1]; if (cur.reduce((x, y) => x + y, 0) + a > lim + 1e-9 && cur.length) split.push([a]); else cur.push(a); }
      expect(set.imccs.map((i) => i.bus.working), `ACB ${acb}`).toEqual(split.map((x) => x.length));
      for (const i of set.imccs.filter((x) => x.bus.working > 0)) expect(i.bus.designA).toBeLessThanOrEqual(acb * 0.9 + 1e-9);
    }
    const def = designImccs(r.loads); // 1000 A ACB: limit 720 A of working current: 1139 A -> 2 IMCCs
    expect(def.imccs).toHaveLength(2);
    expect(def.issues.filter((x) => /status missing/.test(x.message))).toHaveLength(16);
    expect(def.issues.some((x) => /Tailing Paste Thickener/.test(x.message) && /Amps, kW/.test(x.message))).toBe(true);
  });
  it("PMCC_sample_load_list.xlsx is read by the PMCC reader (40 rows)", async () => {
    const p = parseLoadSheet((await read("PMCC_sample_load_list.xlsx")) as never);
    expect(p.errors).toEqual([]);
    expect(p.rows).toHaveLength(40);
  });
});

describe("telling the two files apart by name", () => {
  it("the name starts with PMCC or IMCC, any case", () => {
    expect(["PMCC_sample.xlsx", "imcc list.xlsx", "  IMCC.xlsx", "sample_PMCC.xlsx", "load.xlsx"].map(classifyFile)).toEqual(["PMCC", "IMCC", "IMCC", null, null]);
  });
  it("one file: its own logic, no confirmation; both: both are processed after a confirmation", () => {
    expect(planIntake(["PMCC_a.xlsx"])).toEqual({ pmcc: 0, imcc: undefined, errors: [], needsConfirmation: false });
    expect(planIntake(["IMCC_a.xlsx"])).toEqual({ pmcc: undefined, imcc: 0, errors: [], needsConfirmation: false });
    expect(planIntake(["IMCC_a.xlsx", "PMCC_a.xlsx"])).toEqual({ pmcc: 1, imcc: 0, errors: [], needsConfirmation: true });
  });
  it("a name without the prefix, a non-xlsx file or two files of one kind is refused with a message", () => {
    expect(planIntake(["loads.xlsx"]).errors[0]).toMatch(/must start with PMCC or IMCC/);
    expect(planIntake(["PMCC_a.csv"]).errors[0]).toMatch(/\.xlsx/);
    expect(planIntake(["PMCC_a.xlsx", "PMCC_b.xlsx"]).errors[0]).toMatch(/More than one PMCC/);
    const bad = planIntake(["PMCC_a.xlsx", "loads.xlsx"]);
    expect(bad.pmcc).toBeUndefined();
    expect(bad.needsConfirmation).toBe(false);
  });
});

describe("Excel output", () => {
  const loads = () => [
    L("Screen", "W", "VFD", { kw: 55 }), L("Pump", "w", "DOL", { kw: 30 }), L("Spare pump", "Spare", "DOL", { kw: 30 }), L("Stby pump", "S", "SS", { kw: 30 }),
    ...many(2),
  ];
  const argbOf = (c: ExcelJS.Cell) => (c.fill as ExcelJS.FillPattern | undefined)?.fgColor?.argb;
  it("panel order INC 1 | loads | BC | INC 2 (bus coupler above Incomer 2); ON green, OFF light red; the app's colours; spare as SP; ratings equal the selected ACB; Feeder, not PWR FDR", async () => {
    for (const acb of [1000, 1250]) {
      const set = designImccs(loads(), DEFAULT_MCC_ASSUMPTIONS, { maxIncomerA: acb });
      for (const linked of [true, false]) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await imccWorkbookBuffer(set, DEFAULT_MCC_ASSUMPTIONS, { sourceName: "x.xlsx", linked }));
        expect(wb.worksheets.map((w) => w.name)).toEqual(["Summary", ...set.imccs.map((i) => i.tag), "Notes"]);
        const ws = wb.getWorksheet("IMCC-01")!;
        const texts: string[] = [];
        ws.eachRow((r) => r.eachCell((c) => { if (typeof c.value === "string") texts.push(c.value); }));
        expect(texts.some((t) => t.includes(linked ? 'PMCC outgoing feeder "IMCC-01 Incomer 1"' : "switch OFF"))).toBe(true);
        expect(texts.some((t) => /PWR FDR/.test(t))).toBe(false);
        const rowOf = (m: string) => { let n = 0; ws.eachRow((r) => { if (r.getCell(1).value === m) n = r.number; }); return n; };
        const [inc1, inc2, bc] = [rowOf("INC 1"), rowOf("INC 2"), rowOf("BC")];
        expect(inc1).toBeGreaterThan(0);
        expect(bc).toBeGreaterThan(inc1 + 1); // loads in between
        expect(inc2).toBe(bc + 1); // the bus coupler above Incomer 2, together at the end
        expect(argbOf(ws.getRow(inc1).getCell(3))).toBe("FF80DEEA");
        expect(argbOf(ws.getRow(inc2).getCell(3))).toBe("FF80DEEA");
        expect(argbOf(ws.getRow(bc).getCell(3))).toBe("FFFFF9C4");
        const state = (r: number) => [ws.getRow(r).getCell(5).value, argbOf(ws.getRow(r).getCell(5))];
        expect(state(inc1)).toEqual(["ON", "FFC6EFCE"]);
        expect(state(inc2)).toEqual(["OFF", "FFFFC7CE"]);
        expect(state(bc)).toEqual(["OFF", "FFFFC7CE"]);
        const fills: Record<string, string | undefined> = {};
        ws.eachRow((r) => { const v = r.getCell(5).value; if (r.number > inc1 && r.number < bc && typeof v === "string") fills[v] = argbOf(r.getCell(5)); });
        expect(fills).toEqual({ W: "FFDEB1FB", S: "FFFBB1DE", SP: "FFFFD966" });
        expect([inc1, inc2, bc].map((r) => ws.getRow(r).getCell(17).value)).toEqual([acb, acb, acb]);
        const sums: string[] = [];
        ws.eachRow((r) => { const v = r.getCell(11).value as { formula?: string } | null; if (r.number > 20 && v && typeof v === "object" && v.formula?.startsWith("SUM(")) sums.push(String(r.getCell(1).value)); });
        expect(sums).toEqual(["INC 1"]);
        // the loading block: Normal / Incomer 1 unhealthy columns
        const block: [unknown, unknown, unknown][] = [];
        ws.eachRow((r) => { const k = r.getCell(2).value; if (r.number < 20 && (k === "Incomer 1" || k === "Incomer 2" || k === "Bus coupler")) block.push([k, r.getCell(5).value, r.getCell(6).value]); });
        expect(block).toEqual([["Incomer 1", "ON", "OFF"], ["Bus coupler", "OFF", "ON"], ["Incomer 2", "OFF", "ON"]]);
      }
    }
  });
});
