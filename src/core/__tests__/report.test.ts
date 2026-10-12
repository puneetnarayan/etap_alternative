import { readFileSync } from "node:fs";
import readExcelFile, { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import { DEFAULT_OPTIONS, REPORT_COLORS, buildReportBlob, buildReportSheets, parseLoadSheet, sizeTransformers, type RawLoadRow, type ReportCell } from "..";

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

describe("Excel report formatting", () => {
  const bg = (c: ReportCell) => (c && typeof c === "object" ? c.backgroundColor : undefined);
  const val = (c: ReportCell) => (c && typeof c === "object" ? c.value : c);
  const sheetsFor = async (opts = DEFAULT_OPTIONS, volts = 433) => {
    const raw = await readSheet(buf);
    const result = sizeTransformers(parseLoadSheet(raw).rows, volts, opts);
    const sheets = buildReportSheets(result, { transformerLabel: `6600 / ${volts} V`, options: opts, includePower: true, originalSheet: raw });
    return { raw, result, byName: Object.fromEntries(sheets.map((x) => [x.sheet, x])), names: sheets.map((x) => x.sheet) };
  };

  it("colours are light orange headers, light green within limit, light red over", () => {
    expect(REPORT_COLORS.header).toBe("#FFE0B2");
    expect(REPORT_COLORS.within).toBe("#C6EFCE");
    expect(REPORT_COLORS.over).toBe("#FFC7CE");
  });

  it("includes the original load sheet, second, with the values as uploaded", async () => {
    const { raw, names, byName } = await sheetsFor();
    expect(names.slice(0, 3)).toEqual(["Summary", "Original load list", "Loads"]);
    const orig = byName["Original load list"].data;
    expect(orig).toHaveLength(raw.length);
    raw.forEach((r, i) => r.forEach((v, j) => {
      const got = val(orig[i][j]);
      const want = v === null || v === undefined || v === "" ? null : v;
      expect(got === "" ? null : got).toEqual(want);
    }));
  });

  it("the original sheet survives the round trip through a real .xlsx file", async () => {
    const raw = await readSheet(buf);
    const parsed = parseLoadSheet(raw);
    const result = sizeTransformers(parsed.rows, 433);
    const blob = await buildReportBlob(result, { transformerLabel: "6600 / 433 V", options: DEFAULT_OPTIONS, includePower: true, originalSheet: raw });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((x) => [x.sheet, x.data]));
    expect(Object.keys(sheets).slice(0, 3)).toEqual(["Summary", "Original load list", "Loads"]);
    const orig = sheets["Original load list"];
    expect(orig).toHaveLength(raw.length);
    expect(orig[0]).toEqual(raw[0]);
    expect(orig[1]).toEqual(raw[1].map((v) => (v === "" ? null : v)));
    expect(orig[16]).toEqual(raw[16]); // last row (slurry pump, 690 V) is there too
  });

  it("every table header is light orange; data cells are not", async () => {
    const { byName } = await sheetsFor();
    for (const name of ["Loads", "Original load list", "Issues"]) {
      const first = byName[name].data[0].filter((c) => val(c) !== null && val(c) !== "");
      expect(first.length).toBeGreaterThan(3);
      expect(first.every((c) => bg(c) === REPORT_COLORS.header)).toBe(true);
      expect(byName[name].stickyRowsCount).toBe(1);
    }
    expect(byName["Loads"].data[1].some((c) => bg(c) === REPORT_COLORS.header)).toBe(false);
    // headers inside the combination sheets
    const rows = byName["3 Transformers"].data;
    const headerRows = [
      rows.find((r) => val(r[0]) === "Rating (kVA)"),
      rows.find((r) => val(r[0]) === "Transformer" && val(r[1]) === "Loads"),
      rows.find((r) => val(r[0]) === "Transformer out"),
      rows.find((r) => val(r[0]) === "Transformer" && val(r[1]) === "Row"), // "Loads by transformer" table
    ];
    for (const row of headerRows) {
      expect(row).toBeDefined();
      expect(row!.every((c) => bg(c) === REPORT_COLORS.header)).toBe(true);
    }
    // Summary block headers and the tried table
    const sum = byName["Summary"].data;
    for (const h of ["Assumptions", "Total load", "Recommendation", "Transformers"]) {
      expect(bg(sum.find((r) => val(r[0]) === h)![0])).toBe(REPORT_COLORS.header);
    }
  });

  it("cells within limits are light green (sample: 3 x 2000 kVA is within 80 % / 90 %)", async () => {
    const { byName } = await sheetsFor();
    const rows = byName["3 Transformers"].data;
    // Ratings table: 3 x 2000 -> normal 57.9 %, one out 86.8 %, "Fits"
    const ratings = rows[rows.findIndex((r) => val(r[0]) === "Rating (kVA)") + 1];
    expect([ratings[2], ratings[3], ratings[4]].map(bg)).toEqual([REPORT_COLORS.within, REPORT_COLORS.within, REPORT_COLORS.within]);
    expect(val(ratings[4])).toBe("Fits");
    // Distribution rows T1..T3: loading and "Within 80 %" green
    for (const t of ["T1", "T2", "T3"]) {
      const r = rows.find((x) => val(x[0]) === t && x.length === 7)!;
      expect(bg(r[5])).toBe(REPORT_COLORS.within);
      expect(val(r[6])).toBe("Yes");
      expect(bg(r[6])).toBe(REPORT_COLORS.within);
    }
    // One-out rows: 86.8 % <= 90 % -> green OK
    const oneOutHead = rows.findIndex((r) => val(r[0]) === "Transformer out");
    for (const r of rows.slice(oneOutHead + 1, oneOutHead + 4)) {
      expect(bg(r[3])).toBe(REPORT_COLORS.within);
      expect(val(r[4])).toBe("OK");
      expect(bg(r[4])).toBe(REPORT_COLORS.within);
    }
    // Summary: recommended + the tried table are green
    const sum = byName["Summary"].data;
    expect(bg(sum.find((r) => val(r[0]) === "Recommended")![1])).toBe(REPORT_COLORS.within);
  });

  it("cells over a limit are light red, so the two states are never confused", async () => {
    // 2 x 2000 kVA forced: 3,471.6 kVA is 86.8 % normal (> 80) and 173.6 % one out (> 90); 3 is added and fits
    const { byName } = await sheetsFor({ ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [2] });
    const sum = byName["Summary"].data;
    const row2 = sum.find((r) => val(r[0]) === 2 && val(r[1]) === "2")!;
    expect([row2[3], row2[5], row2[6]].map(bg)).toEqual([REPORT_COLORS.over, REPORT_COLORS.over, REPORT_COLORS.over]);
    expect(val(row2[3])).toBe("No");
    const row3 = sum.find((r) => val(r[0]) === 3 && val(r[1]) === "3")!;
    expect([row3[3], row3[5], row3[6]].map(bg)).toEqual([REPORT_COLORS.within, REPORT_COLORS.within, REPORT_COLORS.within]);
    const two = byName["2 Transformers"].data;
    expect(bg(two[1][0])).toBe(REPORT_COLORS.over); // "DOES NOT FIT"
    const t1 = two.find((r) => val(r[0]) === "T1" && r.length === 7)!;
    expect(bg(t1[5])).toBe(REPORT_COLORS.over);
    expect(val(t1[6])).toBe("No");
    const oo = two[two.findIndex((r) => val(r[0]) === "Transformer out") + 1];
    expect(bg(oo[3])).toBe(REPORT_COLORS.over);
    expect(val(oo[4])).toBe("Over limit");
    const three = byName["3 Transformers"].data;
    expect(bg(three[1][0])).toBe(REPORT_COLORS.within);
  });

  it("a limit is inclusive: exactly 90 % is within", async () => {
    // 1800 kVA on 2 x 2000 = exactly the 90 % one-out limit
    const rows: RawLoadRow[] = [{ rowNumber: 2, description: "x", type: "Feeder", voltageV: 433, amps: (1800 * 1000) / (Math.sqrt(3) * 433) }];
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [2] };
    const sheets = buildReportSheets(sizeTransformers(rows, 433, opts), { transformerLabel: "x", options: opts, includePower: false });
    const two = sheets.find((x) => x.sheet === "2 Transformers")!.data;
    const oo = two[two.findIndex((r) => val(r[0]) === "Transformer out") + 1];
    expect(val(oo[4])).toBe("OK");
    expect(bg(oo[4])).toBe(REPORT_COLORS.within);
  });
});
