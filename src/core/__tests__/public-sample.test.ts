import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import readExcelFile, { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import {
  transformerWise, DEFAULT_OPTIONS, STANDARD_RATINGS_KVA, buildReportBlob, displayCombos, parseLoadSheet, sizeTransformers,
  type SizedLoad,
} from "..";

/**
 * The file served at /sample_load_list.xlsx (the one users download), supplied by the project owner.
 * Expected numbers below are worked out here from the numbers in the file, not taken from the code.
 */
const file = readFileSync(new URL("../../../public/PMCC_sample_load_list.xlsx", import.meta.url));
const buf = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
const SQ3 = Math.sqrt(3);

// 433 V rows of the file (38 of the 40; the grinding mill and slurry pump are 690 V). Only the WORKING loads are counted:
// the 9 standby (S) rows are left out (185 and 160 kW DOL; 250 (ID fan), 200 (FD fan), 250 + 110 + 250 + 200 + 275 kW VFD).
//   DOL motors (kW): 185, 160, 132, 90, 315                                    -> 882 kW, pf 0.7
//   VFD motors (kW): 250 x 2 (ID fan), 200 x 2 (FD fan), 75, 250 x 6 (HGMS 1 and 2 feed pumps),
//                    200 x 3 (thickener 1), 275 x 3 (thickener 2)               -> 3900 kW, pf 0.9
//   amps only: dosing pump 180 A (DOL), booster fan 120 A (VFD), feeders 250, 800, 550, 700 and 900 A
const DOL_KW = 185 + 160 + 132 + 90 + 315;
const VFD_KW = 250 * 2 + 200 * 2 + 75 + 250 * 6 + 200 * 3 + 275 * 3;
const AMPS = 180 + 120 + 250 + 800 + 550 + 700 + 900;
const TOTAL_433 = DOL_KW / (0.7 * 0.95) + VFD_KW / (0.9 * 0.95) + (SQ3 * 433 * AMPS) / 1000; // 8,512.6 kVA
const TOTAL_690 = 450 / (0.7 * 0.95) + 355 / (0.9 * 0.95);

const load = async () => parseLoadSheet(await readSheet(buf));

describe("the downloadable sample file", () => {
  it("is read without problems: 40 rows, every one working (W) or standby (S)", async () => {
    const p = await load();
    expect(p.errors).toEqual([]);
    expect(p.rows).toHaveLength(40);
    expect(p.rows.every((r) => r.status === "W" || r.status === "S")).toBe(true);
    expect(p.rows.every((r) => !r.badFields && !r.statusText)).toBe(true);
    expect(p.rows.filter((r) => r.status === "S").map((r) => r.rowNumber)).toEqual([3, 5, 8, 11, 25, 26, 30, 34, 38]);
    expect(p.rows.map((r) => r.group).filter(Boolean)).toEqual([
      "BFW", "BFW", "CWP", "CWP", "Boiler", "Boiler", "Boiler", "ESP", "ESP", "ESP",
      ...Array<string>(4).fill("HGMS2 Feed Pump"), ...Array<string>(4).fill("HGMS1 Feed Pump"),
      ...Array<string>(4).fill("Thickener1 Feed Pump"), ...Array<string>(4).fill("Thickener2 Feed Pump"),
    ]);
    expect(p.rows.filter((r) => r.voltageV === 690)).toHaveLength(2);
  });

  it("has no author metadata and no typos in the group names", () => {
    const parts = unzipSync(new Uint8Array(buf));
    const core = strFromU8(parts["docProps/core.xml"]);
    expect(core).not.toMatch(/creator|lastModifiedBy/i);
    for (const [name, bytes] of Object.entries(parts)) {
      const text = strFromU8(bytes);
      expect(text, name).not.toMatch(/Puneet|openpyxl|Narayan/i);
      expect(text, name).not.toContain("Bolier");
    }
  });

  it("433 V: 29 working loads are counted; the 9 standby rows and the 2 rows at 690 V are listed as excluded", async () => {
    const r = sizeTransformers((await load()).rows, 433);
    expect(r.loads).toHaveLength(29);
    expect(r.loads.every((l) => l.status === "W")).toBe(true);
    const excluded = r.issues.filter((i) => i.severity === "excluded");
    expect(excluded.map((i) => i.rowNumber)).toEqual([3, 5, 8, 11, 20, 21, 25, 26, 30, 34, 38]);
    expect(excluded.filter((i) => /Standby \(S\)/.test(i.message))).toHaveLength(9);
    expect(r.issues.filter((i) => i.severity === "error")).toEqual([]);
    expect(r.totalKva).toBeCloseTo(TOTAL_433, 6); // 8,512.6 kVA
  });

  it("page defaults (433 V, 2000 kVA, start with 3): 3, 5 and 6 do not fit, 8 = 3 + 3 + 2 does", async () => {
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const r = sizeTransformers((await load()).rows, 433, opts);
    // carrying capacity (one out, 90 %): a set of 3 carries 3600 kVA, a set of 2 carries 1800 kVA
    // 3 = 3600, 5 = 5400, 6 = 7200 (all below 8,512.6), 8 = 9000 kVA
    expect(r.combos.map((c) => [c.count, c.layout, c.fits])).toEqual([
      [3, [3], false], [5, [3, 2], false], [6, [3, 3], false], [8, [3, 3, 2], true],
    ]);
    expect(r.recommended).toMatchObject({ count: 8, ratingKva: 2000 });
    expect(r.incomers).toBe(8);
    expect(r.couplers).toBe(5);
    expect(displayCombos(r).map((c) => c.count)).toEqual([3, 8]);
  });

  it("the placement on 8 transformers really is within every limit and keeps duty groups apart", async () => {
    const r = sizeTransformers((await load()).rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000] });
    const eight = r.combos.find((c) => c.count === 8)!;
    const placed: SizedLoad[] = eight.sections.flatMap((s) => s.loads);
    expect(placed).toHaveLength(29);
    expect(new Set(placed.map((l) => l.rowNumber)).size).toBe(29);
    for (const s of eight.sections) expect(s.kva).toBeLessThanOrEqual(0.8 * 2000 + 1e-6); // normal running <= 80 %
    for (const set of eight.sets) expect(set.kva).toBeLessThanOrEqual(0.9 * 2000 * (set.sections.length - 1) + 1e-6); // one out: the rest at <= 90 %
    expect(eight.sets.reduce((a, x) => a + x.kva, 0)).toBeCloseTo(TOTAL_433, 6);
    expect(eight.oneOut.every((o) => o.ok && o.loadingPct <= 90 + 1e-9)).toBe(true);
    for (const g of ["Boiler", "ESP", "HGMS2 Feed Pump", "HGMS1 Feed Pump", "Thickener1 Feed Pump", "Thickener2 Feed Pump"]) {
      const perSection = eight.sections.map((s) => s.loads.filter((l) => l.group === g).length);
      expect(Math.max(...perSection), g).toBe(1);
    }
  });

  it("690 V: 2 loads, 1,091.9 kVA; a set of 3 by default, a set of 2 when asked for", async () => {
    const rows = (await load()).rows;
    const r = sizeTransformers(rows, 690);
    expect(r.loads).toHaveLength(2);
    expect(r.totalKva).toBeCloseTo(TOTAL_690, 6);
    expect(r.recommended).toMatchObject({ count: 3, ratingKva: 2000 });
    const two = sizeTransformers(rows, 690, { ...DEFAULT_OPTIONS, counts: [2] });
    expect(two.recommended).toMatchObject({ count: 2, ratingKva: 2000 }); // 1,091.9 <= 1,800
    expect(two.couplers).toBe(1);
  });

  it("with all standard ratings on offer, 4000 kVA is not tried at 433 V (690 V only); 6 x 2500 kVA is the first that fits", async () => {
    // 3 = 2 x 2500 x 0.9 + ... : 3 x 2500 carries 4,500 kVA, 5 = 3 + 2 carries 4500 + 2250 = 6,750 kVA, 6 = 3 + 3 carries 9,000 kVA >= 8,512.6 kVA
    expect(TOTAL_433).toBeGreaterThan(6750);
    expect(TOTAL_433).toBeLessThan(9000);
    const r = sizeTransformers((await load()).rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: STANDARD_RATINGS_KVA });
    expect(r.combos.map((c) => [c.count, c.candidates.map((x) => [x.ratingKva, x.fits])])).toEqual([
      [3, [[2000, false], [2500, false]]],
      [5, [[2000, false], [2500, false]]],
      [6, [[2000, false], [2500, true]]],
    ]);
    expect(r.recommended).toMatchObject({ count: 6, ratingKva: 2500 });
  });

  it("amps-only loads are converted at their own voltage (433 V loads at 433 V, 690 V loads at 690 V)", async () => {
    const r = sizeTransformers((await load()).rows, 433);
    const dosing = r.loads.find((l) => l.description.startsWith("Dosing pump"))!;
    expect(dosing.kva).toBeCloseTo((SQ3 * 433 * 180) / 1000, 6); // 135.0 kVA
    expect(dosing.basis).toContain("√3 × 433 V × 180 A");
    expect(dosing.voltageV).toBe(433); // still a 433 V load
    const feeder = r.loads.find((l) => l.description.startsWith("MCC-2"))!;
    expect(feeder.kva).toBeCloseTo((SQ3 * 433 * 800) / 1000, 6); // 600.0 kVA
    // motors given in kW do not depend on the voltage
    const crusher = r.loads.find((l) => l.description === "Crusher")!;
    expect(crusher.kva).toBeCloseTo(315 / (0.7 * 0.95), 6);
  });

  it("one spare feeder of each type and rating, output only: 16 for this file (from the working loads only), and they change no number", async () => {
    const r = sizeTransformers((await load()).rows, 433);
    expect(r.spares.map((x) => x.description)).toEqual([
      "Spare DOL motor feeder, 90 kW (Spare)", "Spare DOL motor feeder, 132 kW (Spare)", "Spare DOL motor feeder, 160 kW (Spare)", "Spare DOL motor feeder, 185 kW (Spare)",
      "Spare DOL motor feeder, 315 kW (Spare)", "Spare DOL motor feeder, 180 A (Spare)",
      "Spare VFD motor feeder, 75 kW (Spare)", "Spare VFD motor feeder, 200 kW (Spare)", "Spare VFD motor feeder, 250 kW (Spare)",
      "Spare VFD motor feeder, 275 kW (Spare)", "Spare VFD motor feeder, 120 A (Spare)",
      "Spare feeder, 250 A (Spare)", "Spare feeder, 550 A (Spare)", "Spare feeder, 700 A (Spare)", "Spare feeder, 800 A (Spare)", "Spare feeder, 900 A (Spare)",
    ]);
    expect(r.spares.find((x) => x.rating === 200 && x.unit === "kW" && x.description.includes("VFD"))!.count).toBe(5); // two FD fans and three thickener 1 pumps (standby ones are not counted)
    expect(r.spares.find((x) => x.rating === 250 && x.unit === "kW")!.count).toBe(8); // two ID fans and six HGMS feed pumps (standby ones are not counted)
    expect(r.totalKva).toBeCloseTo(TOTAL_433, 6); // unchanged by the spares
    expect(r.loads).toHaveLength(29);
  });

  it("the Excel report for it: original sheet included unchanged, W/S counts, recommendation", async () => {
    const raw = await readSheet(buf);
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const result = sizeTransformers(parseLoadSheet(raw).rows, 433, opts);
    const blob = await buildReportBlob(result, { transformerLabel: "6600 / 433 V", options: opts, includePower: true, originalSheet: raw });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((s) => [s.sheet, s.data]));
    expect(Object.keys(sheets)).toEqual(["Summary", "Transformer-wise", "Original load list", "Loads", "3 Transformers", "8 Transformers", "Issues"]);
    const find = (rows: unknown[][], k: string) => rows.find((r) => r[0] === k);
    expect(find(sheets["Summary"], "Recommended")![1]).toBe("8 × 2000 kVA (sets of 3 + 3 + 2)");
    expect(find(sheets["Summary"], "Working (W) loads")![1]).toBe("29");
    expect(find(sheets["Summary"], "Standby (S) rows not counted")![1]).toBe(9);
    expect(find(sheets["Summary"], "Total kVA used for sizing")![1] as number).toBeCloseTo(TOTAL_433, 6);
    const orig = sheets["Original load list"];
    expect(orig).toHaveLength(raw.length);
    expect(orig[0]).toEqual(raw[0]);
    expect(orig[raw.length - 1]).toEqual(raw[raw.length - 1]); // MCC-5 incomer
    // layout of the 8-transformer sheet: 8 incomers, 5 couplers (2 + 2 + 1)
    const six = sheets["8 Transformers"];
    expect(six.filter((r) => r[0] === "Incomer" && r[1] !== "Bus coupler")).toHaveLength(8); // (the legend row has "Bus coupler" next to "Incomer")
    expect(six.filter((r) => r[0] === "Bus coupler")).toHaveLength(5);
    expect(six.filter((r) => r[0] === "Outgoing load")).toHaveLength(29);
  });

  it("the Transformer-wise sheet: per-transformer kW, kVA, current and % loading, the totals, and the tables that show 8 is not an excess", async () => {
    const raw = await readSheet(buf);
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const result = sizeTransformers(parseLoadSheet(raw).rows, 433, opts);
    const w = transformerWise(result, opts)!;
    // 8 transformers = 3 + 3 + 2, 2000 kVA each
    expect(w.rows).toHaveLength(8);
    expect(w.rows.map((r) => r.set)).toEqual([1, 1, 1, 2, 2, 2, 3, 3]);
    // independent arithmetic: kW (input) of every counted load, then the same total as kVA
    const kwDol = DOL_KW / 0.95, kwVfd = VFD_KW / 0.95;
    const kwAmps = (SQ3 * 433 * (180 * 0.7 + 120 * 0.9 + (250 + 800 + 550 + 700 + 900) * 0.7)) / 1000;
    expect(w.totals.kw).toBeCloseTo(kwDol + kwVfd + kwAmps, 6);
    expect(w.totals.kva).toBeCloseTo(TOTAL_433, 6);
    expect(w.totals.currentA).toBeCloseTo((TOTAL_433 * 1000) / (SQ3 * 433), 6); // 11,350 A
    expect(w.totals.installedKva).toBe(16000);
    expect(w.totals.loadingPct).toBeCloseTo((TOTAL_433 / 16000) * 100, 6); // 53.2 %
    expect(w.totals.loads).toBe(29);
    for (const r of w.rows) {
      expect(r.loadingPct).toBeCloseTo((r.kva / 2000) * 100, 9);
      expect(r.currentA).toBeCloseTo((r.kva * 1000) / (SQ3 * 433), 9);
      expect(r.ratedCurrentA).toBeCloseTo((2000 * 1000) / (SQ3 * 433), 9); // 2,667 A
      expect(r.headroomKva).toBeCloseTo(0.8 * 2000 - r.kva, 9);
      expect(r.within).toBe(true);
    }
    expect(w.rows.reduce((a, r) => a + r.kw, 0)).toBeCloseTo(w.totals.kw, 9);
    // numbers tried: 3, 5, 6 do not fit (short of the load), 8 does and is the one taken
    expect(w.quantity.map((q) => [q.count, q.fits, q.final])).toEqual([[3, false, false], [5, false, false], [6, false, false], [8, true, true]]);
    expect(w.quantity.map((q) => q.capacityKva)).toEqual([3600, 5400, 7200, 9000]);
    expect(w.quantity[2].marginKva).toBeCloseTo(7200 - TOTAL_433, 6); // short by 1,312.6 kVA
    expect(w.quantity[3].marginKva).toBeCloseTo(9000 - TOTAL_433, 6); // 487.4 kVA of margin, less than one transformer adds (1,125 kVA)
    expect(w.checks.join(" ")).toMatch(/short by 1,313 kVA/);
    expect(w.candidates.find((c) => c.final)).toMatchObject({ count: 8, ratingKva: 2000, fits: true });
    // the sheet in the Excel report carries the same numbers
    const blob = await buildReportBlob(result, { transformerLabel: "6600 / 433 V", options: opts, includePower: true, originalSheet: raw });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((x) => [x.sheet, x.data]));
    const sh = sheets["Transformer-wise"];
    const total = sh.find((r) => r[0] === "Total")!;
    expect(total[3] as number).toBeCloseTo(w.totals.kw, 6);
    expect(total.filter((v) => typeof v === "number")).toEqual(expect.arrayContaining([Math.round(w.totals.kva) === w.totals.kva ? w.totals.kva : total[5]]));
    expect(sh.filter((r) => /^S[123]-T\d$/.test(String(r[0])))).toHaveLength(8);
    expect(sh.some((r) => r[0] === "Why this number of transformers is not an excess")).toBe(true);
    expect(sh.filter((r) => r[10] === "◄ taken" || r[6] === "◄ taken").length).toBeGreaterThanOrEqual(2);
  });
});
