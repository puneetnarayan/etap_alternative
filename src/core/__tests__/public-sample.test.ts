import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import readExcelFile, { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_OPTIONS, STANDARD_RATINGS_KVA, buildReportBlob, displayCombos, parseLoadSheet, sizeTransformers,
  type SizedLoad,
} from "..";

/**
 * The file served at /sample_load_list.xlsx (the one users download), supplied by the project owner.
 * Expected numbers below are worked out here from the numbers in the file, not taken from the code.
 */
const file = readFileSync(new URL("../../../public/sample_load_list.xlsx", import.meta.url));
const buf = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
const SQ3 = Math.sqrt(3);

// 433 V rows of the file:
//   DOL motors (kW): 185, 185, 160, 160, 132, 90, 315          -> 1227 kW, pf 0.7
//   VFD motors (kW): 250 x 3 (ID fan), 200 x 3 (FD fan), 75,
//                    250 x 8 (HGMS 1 and 2 feed pumps), 110 (Warman slurry pump) -> 3535 kW, pf 0.9
//   amps only: dosing pump 180 A (DOL), booster fan 120 A (VFD), feeders 250 A and 800 A
const DOL_KW = 185 + 185 + 160 + 160 + 132 + 90 + 315;
const VFD_KW = 250 * 3 + 200 * 3 + 75 + 250 * 8 + 110;
const AMPS = 180 + 120 + 250 + 800;
const TOTAL_433 = DOL_KW / (0.7 * 0.95) + VFD_KW / (0.9 * 0.95) + (SQ3 * 433 * AMPS) / 1000;
const TOTAL_690 = 450 / (0.7 * 0.95) + 355 / (0.9 * 0.95);

const load = async () => parseLoadSheet(await readSheet(buf));

describe("the downloadable sample file", () => {
  it("is read without problems: 29 rows, every one working (W) or standby (S)", async () => {
    const p = await load();
    expect(p.errors).toEqual([]);
    expect(p.rows).toHaveLength(29);
    expect(p.rows.every((r) => r.status === "W" || r.status === "S")).toBe(true);
    expect(p.rows.every((r) => !r.badFields && !r.statusText)).toBe(true);
    expect(p.rows.filter((r) => r.status === "S").map((r) => r.rowNumber)).toEqual([3, 5, 8, 11, 25, 26, 30]);
    expect(p.rows.map((r) => r.group).filter(Boolean)).toEqual([
      "BFW", "BFW", "CWP", "CWP", "Boiler", "Boiler", "Boiler", "ESP", "ESP", "ESP",
      ...Array<string>(4).fill("HGMS2 Feed Pump"), ...Array<string>(4).fill("HGMS1 Feed Pump"),
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

  it("433 V: 27 loads (20 W, 7 S) and the 2 rows at 690 V are listed as excluded", async () => {
    const r = sizeTransformers((await load()).rows, 433);
    expect(r.loads).toHaveLength(27);
    expect(r.loads.filter((l) => l.status === "W")).toHaveLength(20);
    expect(r.loads.filter((l) => l.status === "S")).toHaveLength(7);
    expect(r.issues.filter((i) => i.severity === "excluded").map((i) => i.rowNumber)).toEqual([20, 21]);
    expect(r.issues.filter((i) => i.severity === "error")).toEqual([]);
    expect(r.totalKva).toBeCloseTo(TOTAL_433, 6); // 6,961.7 kVA
  });

  it("page defaults (433 V, 2000 kVA, start with 3): 3 and 5 = 3 + 2 do not fit, 6 = 3 + 3 does", async () => {
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const r = sizeTransformers((await load()).rows, 433, opts);
    // 3 x 2000 carries 3600 kVA, 5 = 3 + 2 carries 3600 + 1800 = 5400 kVA, both < 6962; 6 = 3 + 3 carries 7200 kVA
    expect(r.combos.map((c) => [c.count, c.layout, c.fits])).toEqual([[3, [3], false], [5, [3, 2], false], [6, [3, 3], true]]);
    expect(r.recommended).toMatchObject({ count: 6, ratingKva: 2000 });
    expect(r.incomers).toBe(6);
    expect(r.couplers).toBe(4);
    expect(displayCombos(r).map((c) => c.count)).toEqual([3, 6]);
  });

  it("the placement on 6 transformers really is within every limit and keeps standby groups apart", async () => {
    const r = sizeTransformers((await load()).rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: [2000] });
    const six = r.combos.find((c) => c.count === 6)!;
    const placed: SizedLoad[] = six.sections.flatMap((s) => s.loads);
    expect(placed).toHaveLength(27);
    expect(new Set(placed.map((l) => l.rowNumber)).size).toBe(27);
    for (const s of six.sections) expect(s.kva).toBeLessThanOrEqual(0.8 * 2000 + 1e-6); // normal running <= 80 %
    for (const set of six.sets) expect(set.kva).toBeLessThanOrEqual(3600 + 1e-6); // one out: remaining 2 x 2000 at <= 90 %
    expect(six.sets.reduce((a, x) => a + x.kva, 0)).toBeCloseTo(TOTAL_433, 6);
    expect(six.oneOut.every((o) => o.ok && o.loadingPct <= 90 + 1e-9)).toBe(true);
    // duty/standby groups: every member on its own section (the groups of 4 fit on 6 sections)
    for (const g of ["BFW", "CWP", "Boiler", "ESP", "HGMS2 Feed Pump", "HGMS1 Feed Pump"]) {
      const perSection = six.sections.map((s) => s.loads.filter((l) => l.group === g).length);
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

  it("with all three standard ratings on offer, only 4000 kVA fits 3 transformers (6,962 of 7,200 kVA)", async () => {
    expect(TOTAL_433).toBeGreaterThan(4500); // 3 x 2500 carries 0.9 x 2 x 2500 = 4,500 kVA
    expect(TOTAL_433).toBeLessThan(7200); // 3 x 4000 carries 0.9 x 2 x 4000 = 7,200 kVA
    const r = sizeTransformers((await load()).rows, 433, { ...DEFAULT_OPTIONS, ratingsKva: STANDARD_RATINGS_KVA });
    expect(r.combos[0].candidates.map((c) => [c.ratingKva, c.fits])).toEqual([[2000, false], [2500, false], [4000, true]]);
    expect(r.recommended).toMatchObject({ count: 3, ratingKva: 4000 });
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

  it("one spare feeder of each type and rating, output only: 13 for this file, and they change no number", async () => {
    const r = sizeTransformers((await load()).rows, 433);
    expect(r.spares.map((x) => x.description)).toEqual([
      "Spare DOL motor feeder, 90 kW (Spare)", "Spare DOL motor feeder, 132 kW (Spare)", "Spare DOL motor feeder, 160 kW (Spare)", "Spare DOL motor feeder, 185 kW (Spare)",
      "Spare DOL motor feeder, 315 kW (Spare)", "Spare DOL motor feeder, 180 A (Spare)",
      "Spare VFD motor feeder, 75 kW (Spare)", "Spare VFD motor feeder, 110 kW (Spare)", "Spare VFD motor feeder, 200 kW (Spare)", "Spare VFD motor feeder, 250 kW (Spare)", "Spare VFD motor feeder, 120 A (Spare)",
      "Spare feeder, 250 A (Spare)", "Spare feeder, 800 A (Spare)",
    ]);
    expect(r.spares.find((x) => x.rating === 250 && x.unit === "kW")!.count).toBe(11); // three ID fans and eight HGMS feed pumps
    expect(r.totalKva).toBeCloseTo(TOTAL_433, 6); // unchanged by the spares
    expect(r.loads).toHaveLength(27);
  });

  it("the Excel report for it: original sheet included unchanged, W/S counts, recommendation", async () => {
    const raw = await readSheet(buf);
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const result = sizeTransformers(parseLoadSheet(raw).rows, 433, opts);
    const blob = await buildReportBlob(result, { transformerLabel: "6600 / 433 V", options: opts, includePower: true, originalSheet: raw });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((s) => [s.sheet, s.data]));
    expect(Object.keys(sheets)).toEqual(["Summary", "Original load list", "Loads", "3 Transformers", "6 Transformers", "Issues"]);
    const find = (rows: unknown[][], k: string) => rows.find((r) => r[0] === k);
    expect(find(sheets["Summary"], "Recommended")![1]).toBe("6 × 2000 kVA (sets of 3 + 3)");
    expect(find(sheets["Summary"], "Working (W) / standby (S)")![1]).toBe("20 W / 7 S");
    expect(find(sheets["Summary"], "Total kVA used for sizing")![1] as number).toBeCloseTo(TOTAL_433, 6);
    const orig = sheets["Original load list"];
    expect(orig).toHaveLength(raw.length);
    expect(orig[0]).toEqual(raw[0]);
    expect(orig[raw.length - 1]).toEqual(raw[raw.length - 1]); // slurry pump
    // layout of the 6-transformer sheet: 3 + 3 incomers, 2 + 2 couplers
    const six = sheets["6 Transformers"];
    expect(six.filter((r) => r[0] === "Incomer" && r[1] !== "Bus coupler")).toHaveLength(6); // (the legend row has "Bus coupler" next to "Incomer")
    expect(six.filter((r) => r[0] === "Bus coupler")).toHaveLength(4);
    expect(six.filter((r) => r[0] === "Outgoing load")).toHaveLength(27);
  });
});
