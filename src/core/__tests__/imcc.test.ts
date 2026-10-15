import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_RULE, addPmcc, computeImcc, deleteImcc, deletePmcc, duplicateImcc, emptyDesignProject, migrateProject, newId,
  normalizeStatus, parseImccSheet, pmccExportSheet, pmccFeeders, pmccSummary, renamePmcc, saveImcc, syncLinks, validateOperatingState,
  type DesignProject, type Imcc, type ImccLoad, type LoadStatus,
} from "..";
import { imccWorkbookBuffer, pmccWorkbookBuffer } from "../imcc-workbook";
import { mccWorkbookBuffer } from "../mcc-workbook";

const SQ3 = Math.sqrt(3);
// Hand formulas (docs/imcc.md): I = kW x 1000 / (sqrt3 x 433 x pf x eff), pf VFD 0.9, DOL 0.7 (eff 0.93), feeder 0.7.
const iVfd = (kw: number) => (kw * 1000) / (SQ3 * 433 * 0.9);
const iDol = (kw: number) => (kw * 1000) / (SQ3 * 433 * 0.7 * 0.93);
const iFdr = (kw: number) => (kw * 1000) / (SQ3 * 433 * 0.7);
const kvaOf = (a: number) => (SQ3 * 433 * a) / 1000;

const load = (section: 1 | 2 | undefined, description: string, kw: number | undefined, status: string, method: ImccLoad["method"], group?: string): ImccLoad => ({
  id: newId("load"), section, description, kw, status: normalizeStatus(status), statusText: status, method, group,
});

const PMCC_SHEET = [
  ["Description", "W/S", "Type (DOL/VFD/Feeder)", "Amps rating (A)", "kW rating (kW)", "Voltage level (V)"],
  ["Pump A", "W", "DOL", null, 90, 433],
  ["Pump B", "S", "DOL", null, 90, 433],
  ["Lighting DB", "W", "Feeder", 100, null, 433],
];

function base(): { p: DesignProject; p1: string; p2: string } {
  let p = emptyDesignProject("T");
  const a = addPmcc(p, "PMCC-01", { fileName: "p1.xlsx", sheet: PMCC_SHEET.map((r) => [...r]) });
  if (!a.ok) throw new Error("setup");
  p = a.value;
  const b = addPmcc(p, "PMCC-02", { fileName: "p2.xlsx", sheet: PMCC_SHEET.map((r) => [...r]) });
  if (!b.ok) throw new Error("setup");
  p = b.value;
  return { p, p1: p.pmccs[0].id, p2: p.pmccs[1].id };
}

const imccOf = (p1: string, p2: string, loads: ImccLoad[], tag = "IMCC-01"): Imcc => ({
  id: newId("imcc"), tag, description: "Test IMCC", voltageV: 433,
  incomer1: { pmccId: p1, feederRef: "F-12" }, incomer2: { pmccId: p2, feederRef: "F-07" },
  normallyOn: 1, rule: { ...DEFAULT_RULE }, loads,
});

const STD_LOADS = () => [
  load(1, "Screen 1", 5.5, "W", "VFD"),
  load(1, "Spillage pump", 30, "w", "DOL", "SP1"),
  load(1, "Spillage pump spare", 30, "Spare", "DOL"),
  load(1, "Agitator standby", 30, "S", "DOL", "SP1"),
  load(2, "Feed pump", 55, "W", "VFD"),
  load(2, "Hoist", 30, "W", "PWR FDR"),
];

function created() {
  const { p, p1, p2 } = base();
  const imcc = imccOf(p1, p2, STD_LOADS());
  const r = saveImcc(p, imcc);
  if (!r.ok) throw new Error(r.errors.join(" "));
  return { project: r.value, imcc, p1, p2 };
}

describe("Test A: creating an IMCC links its incomers into both PMCCs", () => {
  it("creates one generated feeder per incomer, in the right PMCC, with the IMCC tag", () => {
    const { project, imcc, p1, p2 } = created();
    expect(project.generated.map((g) => [g.pmccId, g.description, g.destination, g.incomer])).toEqual([
      [p1, "IMCC-01 Incomer 1", "IMCC-01", 1], [p2, "IMCC-01 Incomer 2", "IMCC-01", 2],
    ]);
    expect(project.generated.map((g) => g.id)).toEqual([`imcc:${imcc.id}:1`, `imcc:${imcc.id}:2`]);
    expect(pmccFeeders(project, p1).filter((f) => f.origin === "imcc")).toHaveLength(1);
    expect(pmccFeeders(project, p2).filter((f) => f.origin === "imcc")).toHaveLength(1);
    expect(project.generated[0].feederRef).toBe("F-12");
  });

  it("refuses a PMCC or feeder reference that does not exist and changes nothing", () => {
    const { p, p1, p2 } = base();
    const bad = imccOf("pmcc_missing", p2, STD_LOADS());
    const r = saveImcc(p, bad);
    expect(r.ok).toBe(false);
    expect(!r.ok && r.errors.join(" ")).toMatch(/does not exist/);
    const noRef = imccOf(p1, p2, STD_LOADS());
    noRef.incomer2.feederRef = " ";
    const r2 = saveImcc(p, noRef);
    expect(!r2.ok && r2.errors.join(" ")).toMatch(/reference of the PMCC outgoing feeder/);
    expect(p.imccs).toHaveLength(0);
    expect(p.generated).toHaveLength(0);
  });
});

describe("Test B: a change in the IMCC load changes the IMCC and the PMCC numbers", () => {
  it("section working current and the linked PMCC feeder follow the loads (hand-calculated)", () => {
    const { project, imcc, p1, p2 } = created();
    const sec1 = iVfd(5.5) + iDol(30); // W and w; S and Spare excluded
    const sec2 = iVfd(55) + iFdr(30);
    const r = computeImcc(imcc);
    expect(r.sections[0].workingA).toBeCloseTo(sec1, 9);
    expect(r.sections[1].workingA).toBeCloseTo(sec2, 9);
    expect(r.sections[0].workingKva).toBeCloseTo(kvaOf(sec1), 9);
    expect(project.generated[0].ratedA).toBeCloseTo(sec1, 9);
    expect(project.generated[0].designA).toBeCloseTo(sec1 * 1.25, 9);
    expect(project.generated[0].rating).toMatchObject({ kind: "MCCB", ratingA: 100 }); // 1.25 x 69.6 A = 87 A -> 100 A

    const before = pmccSummary(pmccFeeders(project, p1)).workingKva;
    const changed: Imcc = { ...imcc, loads: imcc.loads.map((l) => (l.description === "Screen 1" ? { ...l, kw: 75 } : l)) };
    const r2 = saveImcc(project, changed);
    expect(r2.ok).toBe(true);
    if (!r2.ok) return;
    const sec1b = iVfd(75) + iDol(30);
    expect(r2.value.generated[0].ratedA).toBeCloseTo(sec1b, 9);
    const after = pmccSummary(pmccFeeders(r2.value, p1)).workingKva;
    expect(after - before).toBeCloseTo(kvaOf(sec1b) - kvaOf(sec1), 9);
    expect(r2.value.generated[0].rating).toMatchObject({ kind: "MCCB", ratingA
: 250 }); // 1.25 x 169.6 A = 212 A -> 250 A
    // the other PMCC is unaffected by a change in section 1
    expect(pmccSummary(pmccFeeders(r2.value, p2))).toEqual(pmccSummary(pmccFeeders(project, p2)));
  });
});

describe("Test C: the two incomers are alternative sources, never double-counted", () => {
  it("only the normally-ON incomer's feeder is working (W) demand; the other is S; each PMCC carries its own section only", () => {
    const { project, imcc, p1, p2 } = created();
    const g1 = project.generated[0], g2 = project.generated[1];
    expect([g1.status, g2.status]).toEqual(["W", "S"]);
    const sum1 = pmccSummary(pmccFeeders(project, p1)), sum2 = pmccSummary(pmccFeeders(project, p2));
    const sec = computeImcc(imcc).sections;
    // PMCC totals: manual rows (Pump A 90 kW DOL, Lighting DB 100 A) are the same in both
    const manualW = kvaOf(iDol(90)) + kvaOf(100);
    expect(sum1.workingKva).toBeCloseTo(manualW + sec[0].workingKva, 9);
    expect(sum2.workingKva).toBeCloseTo(manualW, 9); // incomer 2 is standby: not normal working demand
    expect(sum2.standbyKva).toBeCloseTo(kvaOf(iDol(90)) + sec[1].workingKva, 9);
    // the IMCC's total working demand across both PMCCs is counted once, as section 1 only
    expect(sum1.workingKva + sum2.workingKva - 2 * manualW).toBeCloseTo(sec[0].workingKva, 9);
    // switching which incomer is normally ON swaps W and S, still one W only
    const flipped = saveImcc(project, { ...imcc, normallyOn: 2 });
    expect(flipped.ok && flipped.value.generated.map((g) => g.status)).toEqual(["S", "W"]);
  });
});

describe("Test D: operating-state validation", () => {
  it("the only permitted normal state: one incomer ON, the other OFF, coupler OFF", () => {
    expect(validateOperatingState({ incomer1: "ON", incomer2: "OFF", coupler: "OFF" }).ok).toBe(true);
    expect(validateOperatingState({ incomer1: "OFF", incomer2: "ON", coupler: "OFF" }).ok).toBe(true);
  });
  it("both ON (coupler OFF), one ON with coupler ON, both ON with coupler ON and none ON are errors", () => {
    const both = validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "OFF" });
    expect(both.ok).toBe(false);
    expect(both.errors.join(" ")).toMatch(/Both incomers are ON/);
    const oneAndCoupler = validateOperatingState({ incomer1: "ON", incomer2: "OFF", coupler: "ON" });
    expect(oneAndCoupler.ok).toBe(false);
    expect(oneAndCoupler.errors.join(" ")).toMatch(/not assumed to supply both bus sections/);
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "ON" }).errors.length).toBeGreaterThanOrEqual(2);
    expect(validateOperatingState({ incomer1: "OFF", incomer2: "OFF", coupler: "OFF" }).ok).toBe(false);
  });
  it("changing the philosophy is an engineering change: refused without a named authoriser", () => {
    const loose = { exactlyOneIncomerOn: false, couplerMustBeOff: false };
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "OFF" }, loose).errors.join(" ")).toMatch(/engineering change/);
    expect(validateOperatingState({ incomer1: "ON", incomer2: "ON", coupler: "OFF" }, { ...loose, authorisedBy: "J. Doe" }).ok).toBe(true);
    const { p, p1, p2 } = base();
    const i = imccOf(p1, p2, STD_LOADS());
    i.rule = loose;
    expect(saveImcc(p, i).ok).toBe(false);
    i.rule = { ...loose, authorisedBy: "J. Doe" };
    expect(saveImcc(p, i).ok).toBe(true);
  });
});

describe("Test E: W, w, S and Spare", () => {
  it("W and w are the same, S and Spare are not counted, Spare is SP", () => {
    expect(normalizeStatus("W")).toBe("W");
    expect(normalizeStatus("w")).toBe("W");
    expect(normalizeStatus("S")).toBe("S");
    expect(normalizeStatus("Spare")).toBe("SP");
    expect(normalizeStatus("spare")).toBe("SP");
    expect(normalizeStatus("x")).toBeUndefined();
    expect(normalizeStatus("")).toBeUndefined();
    const { imcc } = created();
    const s1 = computeImcc(imcc).sections[0];
    expect([s1.working, s1.standby, s1.spare]).toEqual([2, 1, 1]);
    expect(s1.workingA).toBeCloseTo(iVfd(5.5) + iDol(30), 9);
    expect(s1.standbyKva).toBeCloseTo(kvaOf(iDol(30)), 9);
    expect(s1.connectedKva).toBeCloseTo(kvaOf(iVfd(5.5) + iDol(30) + iDol(30)), 9); // spare carries no load
  });
  it("a missing or invalid status is reported and the row is not counted", () => {
    const { p, p1, p2 } = base();
    const i = imccOf(p1, p2, [load(1, "Fan", 11, "W", "DOL"), load(1, "Mystery", 55, "maybe", "DOL"), load(1, "Blank", 55, "", "DOL")]);
    const r = computeImcc(i);
    expect(r.sections[0].workingA).toBeCloseTo(iDol(11), 9);
    expect(r.issues.filter((x) => x.severity === "error" && /status/.test(x.message))).toHaveLength(2);
    void p;
  });
  it("loads without a bus section are flagged, not allocated", () => {
    const { p1, p2 } = base();
    const r = computeImcc(imccOf(p1, p2, [load(undefined, "Orphan", 30, "W", "DOL")]));
    expect(r.sections[0].workingA + r.sections[1].workingA).toBe(0);
    expect(r.issues.some((x) => /no bus section/.test(x.message))).toBe(true);
  });
  it("spare is exported as SP in the IMCC workbook, and spare rows are kept", async () => {
    const { project, imcc } = created();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await imccWorkbookBuffer(project, imcc.id));
    const ws = wb.worksheets[0];
    const status: string[] = [];
    ws.eachRow((r) => { const v = r.getCell(5).value; if (typeof v === "string") status.push(v); });
    expect(status.filter((s) => s === "SP")).toHaveLength(1);
    expect(status.filter((s) => s === "W")).toHaveLength(4);
    expect(status.filter((s) => s === "S")).toHaveLength(1);
    expect(ws.getRow(1).getCell(1).value).toBe("IMCC-01 – IMCC load list");
  });
});

describe("Test F: Incomer 2 and the bus coupler follow Incomer 1", () => {
  it("always equal to the calculated Incomer 1 rating, also when section 2 is larger, and after recalculation", () => {
    const { imcc } = created();
    const r = computeImcc(imcc);
    expect(r.incomer1.ratingA).toBeDefined();
    expect(r.incomer2.ratingA).toBe(r.incomer1.ratingA);
    expect(r.coupler.ratingA).toBe(r.incomer1.ratingA);
    // section 2 is larger: Incomer 2 does NOT get its own, larger rating; a warning is raised
    expect(r.sections[1].workingA).toBeGreaterThan(r.sections[0].workingA);
    expect(r.issues.some((x) => /Incomer 2 section needs/.test(x.message))).toBe(true);
    const bigger = computeImcc({ ...imcc, loads: [...imcc.loads, load(1, "New VFD", 160, "W", "VFD")] });
    expect(bigger.incomer1.ratingA).toBeGreaterThan(r.incomer1.ratingA!);
    expect(bigger.incomer2.ratingA).toBe(bigger.incomer1.ratingA);
    expect(bigger.coupler.ratingA).toBe(bigger.incomer1.ratingA);
  });
  it("uses the existing breaker series: design 1.25 x current, next rating (hand check)", () => {
    const { imcc } = created();
    const sec1 = iVfd(5.5) + iDol(30);
    expect(computeImcc(imcc).sections[0].designA).toBeCloseTo(sec1 * 1.25, 9);
    expect(computeImcc(imcc).incomer1).toMatchObject({ kind: "MCCB", ratingA: 100 });
  });
});

describe("Test G: rename, change PMCC, delete", () => {
  it("renaming the IMCC renames the PMCC feeder descriptions; manual feeders stay intact", () => {
    const { project, imcc, p1 } = created();
    const manualBefore = pmccFeeders(project, p1).filter((f) => f.origin === "manual");
    const r = saveImcc(project, { ...imcc, tag: "IMCC-07" });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.value.generated.map((g) => g.description)).toEqual(["IMCC-07 Incomer 1", "IMCC-07 Incomer 2"]);
    expect(pmccFeeders(r.value, p1).filter((f) => f.origin === "manual")).toEqual(manualBefore);
  });
  it("changing the upstream PMCC moves the linked record; no stale copy is left behind", () => {
    const { project, imcc, p1, p2 } = created();
    const r = saveImcc(project, { ...imcc, incomer1: { pmccId: p2, feederRef: "F-99" } });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(pmccFeeders(r.value, p1).filter((f) => f.origin === "imcc")).toHaveLength(0);
    expect(pmccFeeders(r.value, p2).filter((f) => f.origin === "imcc")).toHaveLength(2);
    expect(r.value.generated[0].feederRef).toBe("F-99");
    expect(r.value.generated[0].id).toBe(project.generated[0].id); // same stable id
  });
  it("deleting asks for confirmation, lists the dependencies, removes only its own feeders", () => {
    const { project, imcc, p1, p2 } = created();
    const manual1 = pmccFeeders(project, p1).filter((f) => f.origin === "manual");
    const ask = deleteImcc(project, imcc.id);
    expect(ask.needsConfirmation).toBe(true);
    expect(ask.dependencies.map((d) => d.pmccId)).toEqual([p1, p2]);
    expect(ask.project).toBeUndefined();
    const done = deleteImcc(project, imcc.id, true);
    expect(done.project!.imccs).toHaveLength(0);
    expect(done.project!.generated).toHaveLength(0);
    expect(pmccFeeders(done.project!, p1)).toEqual(manual1);
    // a PMCC that still supplies an IMCC cannot be deleted
    expect(deletePmcc(project, p1).ok).toBe(false);
    expect(deletePmcc(done.project!, p1).ok).toBe(true);
  });
  it("a duplicate gets a new tag and its own links; the original is unchanged", () => {
    const { project, imcc } = created();
    const r = duplicateImcc(project, imcc.id);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.value.imccs.map((i) => i.tag)).toEqual(["IMCC-01", "IMCC-01-copy"]);
    expect(r.value.generated).toHaveLength(4);
    expect(new Set(r.value.generated.map((g) => g.id)).size).toBe(4);
  });
  it("PMCC tags and IMCC tags are unique; renaming a PMCC keeps the links", () => {
    const { project, p1 } = created();
    expect(addPmcc(project, "pmcc-01").ok).toBe(false);
    const r = renamePmcc(project, p1, "PMCC-A");
    expect(r.ok && r.value.generated[0].pmccId).toBe(p1);
  });
});

describe("Test H: projects without IMCCs behave as before", () => {
  it("an older project (no imccs, no generated) opens with safe defaults, user data kept", () => {
    const legacy = { name: "Old", pmccs: [{ id: "x1", tag: "PMCC-X", fileName: "a.xlsx", sheet: PMCC_SHEET }] };
    const p = migrateProject(legacy);
    expect(p).toMatchObject({ version: 1, name: "Old", imccs: [], generated: [] });
    expect(p.pmccs[0].sheet).toEqual(PMCC_SHEET);
    expect(migrateProject(undefined)).toMatchObject({ pmccs: [], imccs: [] });
  });
  it("the PMCC export sheet and workbook are identical to the existing MCC workbook when no IMCC is linked", async () => {
    const { p, p1 } = base();
    expect(pmccExportSheet(p, p1)).toEqual(PMCC_SHEET);
    const mine = new ExcelJS.Workbook(); await mine.xlsx.load(await pmccWorkbookBuffer(p, p1));
    const orig = new ExcelJS.Workbook(); await orig.xlsx.load((await mccWorkbookBuffer(PMCC_SHEET, "p1.xlsx")).buffer);
    expect(mine.worksheets.map((w) => w.name)).toEqual(orig.worksheets.map((w) => w.name));
    const dump = (wb: ExcelJS.Workbook) => wb.worksheets.map((w) => { const out: unknown[] = []; w.eachRow((r) => out.push(JSON.stringify(r.values))); return out; });
    expect(dump(mine)).toEqual(dump(orig));
  });
});

describe("Test I: Excel outputs are consistent with the linked design", () => {
  it("PMCC workbook: original rows kept, IMCC incomer rows added with source/destination, IMCC links sheet agrees", async () => {
    const { project, p1, p2 } = created();
    for (const [pid, n] of [[p1, 1], [p2, 2]] as const) {
      const sheet = pmccExportSheet(project, pid)!;
      expect(sheet.slice(0, PMCC_SHEET.length).map((r) => r.slice(0, 6))).toEqual(PMCC_SHEET); // originals unchanged
      const added = sheet[sheet.length - 1];
      expect(added[0]).toBe(`IMCC-01 Incomer ${n} (to IMCC-01)`);
      expect(added[1]).toBe(n === 1 ? "W" : "S");
      expect(added[3]).toBeCloseTo(project.generated[n - 1].ratedA, 9);
      expect(added.slice(6)).toEqual(["Auto from IMCC", `IMCC-01 Incomer ${n}`]);
      const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await pmccWorkbookBuffer(project, pid));
      expect(wb.worksheets.map((w) => w.name)).toEqual(["Load list", "Notes", "IMCC links"]);
      const links = wb.getWorksheet("IMCC links")!;
      expect(links.getRow(5).getCell(2).value).toBe(`IMCC-01 Incomer ${n}`);
      expect(links.getRow(5).getCell(6).value).toBeCloseTo(project.generated[n - 1].ratedA, 9);
      expect(links.getRow(5).getCell(5).value).toBe(n === 1 ? "W" : "S");
      const ll = wb.getWorksheet("Load list")!;
      let found = 0; ll.eachRow((r) => { if (String(r.getCell(1).value).startsWith(`IMCC-01 Incomer ${n}`)) found++; });
      expect(found).toBe(1);
    }
  });
  it("IMCC workbook: sections, totals as formulas, incomer 2 and coupler ratings equal to incomer 1", async () => {
    const { project, imcc } = created();
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await imccWorkbookBuffer(project, imcc.id));
    const ws = wb.worksheets[0];
    const texts: string[] = [];
    ws.eachRow((r) => r.eachCell((c) => { if (typeof c.value === "string") texts.push(c.value); }));
    expect(texts.some((t) => t.startsWith("INCOMER 1 – supplied from PMCC-01, feeder F-12"))).toBe(true);
    expect(texts.some((t) => t.startsWith("INCOMER 2 – supplied from PMCC-02, feeder F-07"))).toBe(true);
    const ratings: string[] = [];
    ws.eachRow((r) => { const k = r.getCell(2).value; if (k === "Incomer 1" || k === "Incomer 2" || k === "Bus coupler") ratings.push(String(r.getCell(14).value)); });
    expect(ratings).toHaveLength(3);
    expect(new Set(ratings).size).toBe(1);
    let totalFormulas = 0;
    ws.eachRow((r) => { const v = r.getCell(11).value as { formula?: string } | null; if (v && typeof v === "object" && v.formula?.startsWith("SUM(")) totalFormulas++; });
    expect(totalFormulas).toBe(2);
    expect(wb.worksheets.map((w) => w.name)).toEqual(["IMCC-01", "Notes"]);
  });
});

describe("Test J: no duplicates on repeated save / recalculation", () => {
  it("saving and re-syncing many times keeps exactly two generated feeders with the same ids", () => {
    let { project } = created();
    const ids = project.generated.map((g) => g.id);
    const imcc = project.imccs[0];
    for (let n = 0; n < 5; n++) {
      const r = saveImcc(project, imcc);
      expect(r.ok).toBe(true);
      if (r.ok) project = syncLinks(r.value);
    }
    expect(project.generated.map((g) => g.id)).toEqual(ids);
    expect(project.imccs).toHaveLength(1);
    // round trip through JSON (save / reopen) gives the same
    const reopened = migrateProject(JSON.parse(JSON.stringify(project)));
    expect(reopened.generated.map((g) => g.id)).toEqual(ids);
    expect(pmccFeeders(reopened, project.pmccs[0].id).filter((f) => f.origin === "imcc")).toHaveLength(1);
  });
});

describe("importing the reference layout", () => {
  const sheet = [
    [null, "ECR-X MCC FEEDERS"],
    [null, "S. NO.", null, null, "KW", "Current in Amps", "W/S/Spare", "Method of starting ", "KVA", "Group", "Working KVA"],
    [null, "INCOMER- 1"],
    [null, 1, "Linear screen - 1", null, 5.5, 8.5, "W", "VFD", 6.1, "Linear screen operation"],
    [null, 2, "Linear screen - Spare", null, 5.5, 8.5, "Spare", "VFD"],
    [null, 3, "Spillage pump", "Rating changed from 5.5kW to 30kW", 30, 60, "w", "DOL"],
    [null, null, null, null, null, null, null, null, 6.1],
    [null, "INCOMER- 2"],
    [null, 4, "Pump 2", null, 15, null, "S", "DOL"],
    [null, 5, "1-ph 6A Feeder-1"],
  ];
  it("reads sections, statuses (W, w, Spare, S, missing), methods, groups and notes; skips totals", () => {
    const r = parseImccSheet(sheet);
    expect(r.errors).toEqual([]);
    expect(r.incomerRows).toBe(2);
    expect(r.loads.map((l) => [l.section, l.description, l.kw, l.status, l.method])).toEqual([
      [1, "Linear screen - 1", 5.5, "W", "VFD"], [1, "Linear screen - Spare", 5.5, "SP", "VFD"], [1, "Spillage pump", 30, "W", "DOL"],
      [2, "Pump 2", 15, "S", "DOL"], [2, "1-ph 6A Feeder-1", undefined, undefined, undefined],
    ]);
    expect(r.loads[0].group).toBe("Linear screen operation");
    expect(r.loads[2].note).toBe("Rating changed from 5.5kW to 30kW");
    // the reference's own current column is not imported
    expect(r.loads.every((l) => l.ratedA === undefined)).toBe(true);
  });
  it("reports a sheet without the expected header", () => {
    expect(parseImccSheet([["a", "b"], [1, 2]]).errors[0]).toMatch(/S\. NO\./);
  });
});

export type { LoadStatus };
