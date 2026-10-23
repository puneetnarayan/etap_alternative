import { readFileSync } from "node:fs";
import { strFromU8, unzipSync } from "fflate";
import readExcelFile, { readSheet } from "read-excel-file/universal";
import { describe, expect, it } from "vitest";
import { DEFAULT_OPTIONS, REPORT_COLORS, STANDARD_MCC_FEEDERS, buildReportBlob, buildReportSheets, kindColor, parseLoadSheet, sizeTransformers, spareKey, type RawLoadRow, type ReportCell } from "..";

/** Show both a set of 2 and a set of 3 (the page starts with 3 only). */
const BOTH = { ...DEFAULT_OPTIONS, counts: [2, 3] };
const file = readFileSync(new URL("./fixtures/small_plant_load_list.xlsx", import.meta.url));
const buf = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;

async function report(includePower: boolean, volts = 433) {
  const parsed = parseLoadSheet(await readSheet(buf));
  const result = sizeTransformers(parsed.rows, volts, BOTH);
  const blob = await buildReportBlob(result, {
    transformerLabel: `6600 / ${volts} V`, fileName: "sample.xlsx", options: BOTH, includePower,
  });
  const sheets = await readExcelFile(await blob.arrayBuffer());
  return { result, sheets: Object.fromEntries(sheets.map((s) => [s.sheet, s.data])) as Record<string, unknown[][]> };
}
const find = (rows: unknown[][], label: string) => rows.find((r) => r[0] === label);

describe("Excel report", () => {
  it("has summary, loads, one sheet per transformer combination, and issues", async () => {
    const { sheets } = await report(true);
    expect(Object.keys(sheets)).toEqual(["Summary", "Transformer-wise", "Justification", "Loads", "2 Transformers", "3 Transformers", "Issues"]);
  });

  it("summary numbers match the calculation", async () => {
    const { result, sheets } = await report(true);
    const s = sheets["Summary"];
    expect(find(s, "Transformer")![1]).toBe("6600 / 433 V");
    expect(find(s, "Recommended")![1]).toBe("3 × 2000 kVA");
    expect(find(s, "Total kVA used for sizing")![1] as number).toBeCloseTo(result.totalKva, 6);
    expect(find(s, "Total kW (input)")![1] as number).toBeCloseTo(result.totals.kw, 6);
    expect(find(s, "Total kvar")![1] as number).toBeCloseTo(result.totals.kvar, 6);
    expect(find(s, "Working (W) loads")![1]).toBe("12");
    expect(find(s, "Standby (S) rows not counted")![1]).toBe(2);
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
      const outgoing = rows.filter((r) => r[0] === "Outgoing load");
      expect(outgoing).toHaveLength(result.loads.length);
      expect([...new Set(outgoing.map((r) => r[1]))].sort()).toEqual(Array.from({ length: n }, (_, i) => `T${i + 1}`));
      expect(outgoing.map((r) => r[2]).sort()).toEqual(result.loads.map((l) => l.rowNumber).sort());
    }
  });

  it("2-transformer sheet shows 2 x 2500 (4000 kVA is 690 V only) with one coupler; one-out table is present", async () => {
    const { sheets } = await report(true);
    const rows = sheets["2 Transformers"];
    expect(rows[0][0]).toBe("2 × 2500 kVA transformers, 6600 / 433 V");
    expect(String(rows[2][0])).toContain("2 incomers, 1 bus coupler");
    expect(String(rows[1][0])).toContain("DOES NOT FIT"); // 2 x 2500 kVA carries 2,250 kVA, the working load is 2,952.8 kVA
    expect(rows.some((r) => r[0] === "T1")).toBe(true);
    expect(rows.some((r) => r[0] === "Transformer out")).toBe(true); // the one-out table is present
  });

  it("loads sheet carries W and S for each load", async () => {
    const { sheets } = await report(true);
    const rows = sheets["Loads"];
    const col = (rows[0] as string[]).indexOf("W/S");
    const values = rows.slice(1, -1).map((r) => r[col]);
    expect(values.filter((v) => v === "S")).toHaveLength(0); // standby rows are not counted, so they are not on the Loads sheet
    expect(values.filter((v) => v === "W")).toHaveLength(12);
  });

  it("issues sheet lists the excluded 690 V rows", async () => {
    const { sheets } = await report(true);
    const issues = sheets["Issues"];
    const excluded = issues.filter((r) => r[2] === "Excluded");
    expect(excluded).toHaveLength(4); // the two 690 V rows and the two standby rows
    expect(excluded.filter((r) => /Standby \(S\)/.test(String(r[3])))).toHaveLength(2);
    expect(excluded.filter((r) => /690 V/.test(String(r[3])))).toHaveLength(2);
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
    expect(Object.keys(sheets)).toEqual(["Summary", "Transformer-wise", "Justification", "Loads", "2 Transformers", "3 Transformers", "Issues"]);
    expect(String(sheets["3 Transformers"][2][0])).toContain("Added because the loads could not be balanced");
    expect(sheets["2 Transformers"][1][0]).toContain("DOES NOT FIT");
    expect(find(sheets["Summary"], "Recommended")![1]).toBe("3 × 2000 kVA");
  });
});

describe("Excel report with sets of transformers", () => {
  it("5 transformers = sets of 3 + 2: sets, couplers, names, and every load once", async () => {
    // 60 feeders of 83.33 kVA = 5,000 kVA; 3 x 2000 kVA (3,600) does not fit, 5 (3+2: 5,400) does
    const rows: RawLoadRow[] = Array.from({ length: 60 }, (_, i) => ({
      rowNumber: i + 2, description: `F${i}`, type: "Feeder", voltageV: 433, status: "W" as const,
      amps: (5000 / 60 * 1000) / (Math.sqrt(3) * 433),
    }));
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [3] };
    const result = sizeTransformers(rows, 433, opts);
    const blob = await buildReportBlob(result, { transformerLabel: "6600 / 433 V", options: opts, includePower: true });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((s) => [s.sheet, s.data]));
    // selected (3) and the final one (5) get a sheet; 4 is only in the summary table
    expect(Object.keys(sheets)).toEqual(["Summary", "Transformer-wise", "Justification", "Loads", "3 Transformers", "5 Transformers"]);
    const sum = sheets["Summary"];
    expect(find(sum, "Recommended")![1]).toBe("5 × 2000 kVA (1 set of 3 + 1 set of 2)");
    expect(find(sum, "Incomers")![1]).toBe(5);
    expect(find(sum, "Bus couplers")![1]).toBe(3);
    const tried = sum.filter((r) => typeof r[0] === "number" && typeof r[1] === "string" && /^\d+ sets? of [23]( \+ \d+ sets? of [23])*$/.test(r[1] as string));
    expect(tried.map((r) => [r[0], r[1], r[3]])).toEqual([[3, "1 set of 3", "No"], [5, "1 set of 3 + 1 set of 2", "Yes"]]); // 4 (2 + 2) adds nothing: skipped
    const five = sheets["5 Transformers"];
    expect(String(five[3][0])).toContain("5 incomers, 3 bus couplers, in 2 independent sets (3 + 2)");
    expect(five.some((r) => String(r[0]).startsWith("Set 1: 3 transformers"))).toBe(true);
    expect(five.some((r) => String(r[0]).startsWith("Set 2: 2 transformers"))).toBe(true);
    const body = five.filter((r) => r[0] === "Outgoing load");
    expect(body).toHaveLength(60);
    expect([...new Set(body.map((r) => r[1]))].sort()).toEqual(["S1-T1", "S1-T2", "S1-T3", "S2-T1", "S2-T2"]);
  });
});

describe("Excel report formatting", () => {
  const bg = (c: ReportCell) => (c && typeof c === "object" ? c.backgroundColor : undefined);
  const val = (c: ReportCell) => (c && typeof c === "object" ? c.value : c);
  const sheetsFor = async (opts = BOTH, volts = 433) => {
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
    expect(names.slice(0, 5)).toEqual(["Summary", "Transformer-wise", "Justification", "Original load list", "Loads"]);
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
    expect(Object.keys(sheets).slice(0, 5)).toEqual(["Summary", "Transformer-wise", "Justification", "Original load list", "Loads"]);
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
      rows.find((r) => val(r[0]) === "Element" && val(r[1]) === "Name"), // bus layout table
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
    // 2 x 2000 kVA forced: 2,952.8 kVA (working loads only) is 73.8 % normal (<= 80, within) but 147.6 % one out (> 90, over); 3 is added and fits
    const { byName } = await sheetsFor({ ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [2] });
    const sum = byName["Summary"].data;
    const row2 = sum.find((r) => val(r[0]) === 2 && val(r[1]) === "1 set of 2")!;
    expect([row2[3], row2[6], row2[7]].map(bg)).toEqual([REPORT_COLORS.over, REPORT_COLORS.within, REPORT_COLORS.over]);
    expect(val(row2[3])).toBe("No");
    const row3 = sum.find((r) => val(r[0]) === 3 && val(r[1]) === "1 set of 3")!;
    expect([row3[3], row3[6], row3[7]].map(bg)).toEqual([REPORT_COLORS.within, REPORT_COLORS.within, REPORT_COLORS.within]);
    const two = byName["2 Transformers"].data;
    expect(bg(two[1][0])).toBe(REPORT_COLORS.over); // "DOES NOT FIT"
    const t1 = two.find((r) => val(r[0]) === "T1" && r.length === 7)!;
    expect(bg(t1[5])).toBe(REPORT_COLORS.within); // each section ~73.8 % <= 80 %
    expect(val(t1[6])).toBe("Yes");
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

describe("Excel bus layout: incomers (cyan), outgoing loads, bus couplers (light yellow)", () => {
  const bg = (c: ReportCell) => (c && typeof c === "object" ? c.backgroundColor : undefined);
  const val = (c: ReportCell) => (c && typeof c === "object" ? c.value : c);
  const sheetFor = async (opts = BOTH, includePower = true) => {
    const raw = await readSheet(buf);
    const result = sizeTransformers(parseLoadSheet(raw).rows, 433, opts);
    const sheets = buildReportSheets(result, { transformerLabel: "6600 / 433 V", options: opts, includePower });
    return { result, sheets: Object.fromEntries(sheets.map((x) => [x.sheet, x])) };
  };
  /** The element column of the layout tables, in order, with the row cells. */
  const layout = (rows: ReportCell[][], withSpares = false) => {
    const heads = rows.map((r, i) => (val(r[0]) === "Element" && val(r[1]) === "Name" ? i : -1)).filter((i) => i >= 0);
    return heads.map((h) => {
      const body: ReportCell[][] = [];
      for (let i = h + 1; i < rows.length && rows[i].length > 0; i++) if (withSpares || val(rows[i][0]) !== "Spare feeder") body.push(rows[i]);
      return body;
    });
  };

  it("2 transformers: Incomer | loads | Bus coupler | loads | Incomer", async () => {
    const { result, sheets } = await sheetFor();
    const [table] = layout(sheets["2 Transformers"].data);
    const kinds = table.map((r) => val(r[0]));
    const nLoads = result.combos[0].sections.map((s) => s.loads.length);
    expect(kinds).toEqual([
      "Incomer", ...Array(nLoads[0]).fill("Outgoing load"), "Bus coupler", ...Array(nLoads[1]).fill("Outgoing load"), "Incomer",
    ]);
    expect(table.filter((r) => val(r[0]) === "Incomer").map((r) => val(r[1]))).toEqual(["T1", "T2"]);
    expect(table.filter((r) => val(r[0]) === "Bus coupler").map((r) => val(r[1]))).toEqual(["BC1"]);
  });

  it("3 transformers: Inc T1 | loads | BC1 | Inc T2 | loads | BC2 | loads | Inc T3 (one coupler between neighbours)", async () => {
    const { result, sheets } = await sheetFor();
    const [table] = layout(sheets["3 Transformers"].data);
    const compact: string[] = [];
    for (const r of table) {
      const k = String(val(r[0]));
      if (k === "Outgoing load") { if (compact[compact.length - 1] !== "loads") compact.push("loads"); }
      else compact.push(`${k} ${val(r[1])}`);
    }
    expect(compact).toEqual(["Incomer T1", "loads", "Bus coupler BC1", "Incomer T2", "loads", "Bus coupler BC2", "loads", "Incomer T3"]);
    expect(table.filter((r) => val(r[0]) === "Outgoing load")).toHaveLength(result.loads.length);
  });

  it("incomer rows are cyan end to end, bus coupler rows light yellow, outgoing loads are not coloured", async () => {
    expect(REPORT_COLORS.incomer).toBe("#80DEEA");
    expect(REPORT_COLORS.coupler).toBe("#FFF9C4");
    const { sheets } = await sheetFor();
    for (const name of ["2 Transformers", "3 Transformers"]) {
      const [table] = layout(sheets[name].data);
      for (const r of table) {
        const kind = val(r[0]);
        const colours = new Set(r.map(bg));
        if (kind === "Incomer") expect([...colours]).toEqual([REPORT_COLORS.incomer]);
        else if (kind === "Bus coupler") expect([...colours]).toEqual([REPORT_COLORS.coupler]);
        else expect(colours.has(REPORT_COLORS.incomer) || colours.has(REPORT_COLORS.coupler)).toBe(false);
      }
    }
  });

  it("every run of outgoing loads lies between an incomer and a bus coupler", async () => {
    const { sheets } = await sheetFor();
    for (const name of ["2 Transformers", "3 Transformers"]) {
      const kinds = layout(sheets[name].data)[0].map((r) => String(val(r[0])));
      let i = 0;
      while (i < kinds.length) {
        if (kinds[i] !== "Outgoing load") { i++; continue; }
        let j = i;
        while (j < kinds.length && kinds[j] === "Outgoing load") j++;
        const neighbours = [kinds[i - 1], kinds[j]].sort();
        expect(neighbours, `${name} loads ${i}..${j}`).toEqual(["Bus coupler", "Incomer"]);
        i = j;
      }
    }
  });

  it("an incomer row carries its section's load; the section's loads add up to it", async () => {
    const { sheets } = await sheetFor();
    const [table] = layout(sheets["3 Transformers"].data);
    const kvaCol = (table[0].length - 2); // kVA is the second-last column
    let section = "";
    const sums: Record<string, number> = {};
    const incomerKva: Record<string, number> = {};
    for (const r of table) {
      const kind = val(r[0]);
      if (kind === "Outgoing load") sums[String(val(r[1]))] = (sums[String(val(r[1]))] ?? 0) + Number(val(r[kvaCol]));
      if (kind === "Incomer") { section = String(val(r[1])); incomerKva[section] = Number(val(r[kvaCol])); }
    }
    for (const t of ["T1", "T2", "T3"]) expect(sums[t]).toBeCloseTo(incomerKva[t], 6);
  });

  it("rows line up with the header with and without the kW / kvar / power-factor columns", async () => {
    for (const includePower of [true, false]) {
      const { sheets } = await sheetFor(BOTH, includePower);
      const rows = sheets["3 Transformers"].data;
      const hIdx = rows.findIndex((r) => val(r[0]) === "Element" && val(r[1]) === "Name");
      const width = rows[hIdx].length;
      expect(width).toBe(includePower ? 14 : 11); // 8 base columns + 3 power columns + kVA + Current + Basis
      for (const r of layout(rows, true)[0]) expect(r).toHaveLength(width); // spare rows line up too
    }
  });

  it("several sets: one table per set; couplers only inside a set (named S1-BC1 ...), none between sets", async () => {
    const rows: RawLoadRow[] = Array.from({ length: 60 }, (_, i) => ({
      rowNumber: i + 2, description: `F${i}`, type: "Feeder", voltageV: 433, status: "W" as const,
      amps: (5000 / 60 * 1000) / (Math.sqrt(3) * 433),
    }));
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [5] };
    const result = sizeTransformers(rows, 433, opts);
    const sheet = buildReportSheets(result, { transformerLabel: "x", options: opts, includePower: false }).find((x) => x.sheet === "5 Transformers")!;
    const tables = layout(sheet.data);
    expect(tables).toHaveLength(2); // set of 3 and set of 2
    const names = (t: ReportCell[][], kind: string) => t.filter((r) => val(r[0]) === kind).map((r) => val(r[1]));
    expect(names(tables[0], "Incomer")).toEqual(["S1-T1", "S1-T2", "S1-T3"]);
    expect(names(tables[0], "Bus coupler")).toEqual(["S1-BC1", "S1-BC2"]);
    expect(names(tables[1], "Incomer")).toEqual(["S2-T1", "S2-T2"]);
    expect(names(tables[1], "Bus coupler")).toEqual(["S2-BC1"]);
    // 5 incomers, 3 couplers in total = what the summary says
    const total = (k: string) => tables.reduce((a, t) => a + t.filter((r) => val(r[0]) === k).length, 0);
    expect([total("Incomer"), total("Bus coupler")]).toEqual([5, 3]);
    expect(total("Outgoing load")).toBe(60);
  });

  it("the cyan / yellow / orange / green colours are in the real .xlsx, with the legend row", async () => {
    const raw = await readSheet(buf);
    const result = sizeTransformers(parseLoadSheet(raw).rows, 433);
    const blob = await buildReportBlob(result, { transformerLabel: "x", options: DEFAULT_OPTIONS, includePower: true });
    const sheets = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((x) => [x.sheet, x.data]));
    const rows = sheets["3 Transformers"];
    const legend = rows.findIndex((r) => r[0] === "Incomer" && r[1] === "Bus coupler" && r[2] === "Outgoing load");
    expect(legend).toBeGreaterThan(0);
    expect(rows.filter((r) => r[0] === "Incomer")).toHaveLength(4); // legend cell + 3 incomers
    expect(rows.filter((r) => r[0] === "Bus coupler")).toHaveLength(2); // 2 couplers (the legend has it in column 2)
  });
});

describe("Excel report when loads cannot be balanced / cannot be placed", () => {
  const kvaRows = (kvas: number[]): RawLoadRow[] =>
    kvas.map((kva, i) => ({ rowNumber: i + 2, description: `L${i}`, type: "Feeder", voltageV: 433, amps: (kva * 1000) / (Math.sqrt(3) * 433) }));
  const val = (c: ReportCell) => (c && typeof c === "object" ? c.value : c);

  it("states why a number of transformers does not fit (5 x 1000 kVA on 3 + 2)", async () => {
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [5] };
    const result = sizeTransformers(kvaRows([1000, 1000, 1000, 1000, 1000]), 433, opts);
    const sheets = buildReportSheets(result, { transformerLabel: "x", options: opts, includePower: false });
    const five = sheets.find((x) => x.sheet === "5 Transformers")!.data;
    expect(five.some((r) => val(r[4]) === "Does not fit: loads cannot be balanced within the limits")).toBe(true);
    expect(sheets.map((x) => x.sheet)).toContain("6 Transformers"); // 6 was added and fits
    expect(String(val(sheets.find((x) => x.sheet === "6 Transformers")!.data[2][0]))).toContain("could not be balanced");
  });

  it("lists loads that no transformer section can carry on the Summary", async () => {
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000] };
    const result = sizeTransformers(kvaRows([1700, 300]), 433, opts);
    const sum = buildReportSheets(result, { transformerLabel: "x", options: opts, includePower: false })[0].data;
    expect(sum.some((r) => String(val(r[1])).startsWith("None: a load is too big"))).toBe(true);
    const row = sum.find((r) => val(r[0]) === "Cannot be placed")!;
    expect(String(val(row[1]))).toContain("L0");
  });
});

describe("Excel report: pastel load kinds, number formats, centring, no author metadata", () => {
  type Styled = { value: string | number | boolean; backgroundColor?: string; format?: string; align?: string };
  const st = (c: ReportCell): Styled | null => (c !== null && typeof c === "object" ? (c as Styled) : null);
  const val = (c: ReportCell) => (c !== null && typeof c === "object" ? c.value : c);
  const build = async (opts = BOTH, includePower = true) => {
    const raw = await readSheet(buf);
    const result = sizeTransformers(parseLoadSheet(raw).rows, 433, opts);
    const o = { transformerLabel: "6600 / 433 V", options: opts, includePower, originalSheet: raw };
    return { raw, result, o, sheets: buildReportSheets(result, o) };
  };

  // --- colour science helpers (CIE76 distance between sRGB colours)
  const rgb = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const lab = (h: string) => {
    const [r, g, b] = rgb(h).map((c) => c / 255).map((c) => (c > 0.04045 ? ((c + 0.055) / 1.055) ** 2.4 : c / 12.92));
    const f = (t: number) => (t > 0.008856 ? Math.cbrt(t) : 7.787 * t + 16 / 116);
    const X = f((0.4124 * r + 0.3576 * g + 0.1805 * b) / 0.95047), Y = f(0.2126 * r + 0.7152 * g + 0.0722 * b), Z = f((0.0193 * r + 0.1192 * g + 0.9505 * b) / 1.08883);
    return [116 * Y - 16, 500 * (X - Y), 200 * (Y - Z)];
  };
  const dE = (a: string, b: string) => Math.hypot(...lab(a).map((v, i) => v - lab(b)[i]));
  const lightness = (h: string) => (Math.max(...rgb(h)) + Math.min(...rgb(h))) / 2 / 255;

  const NEW = { W: REPORT_COLORS.working, S: REPORT_COLORS.standby, DOL: REPORT_COLORS.dol, VFD: REPORT_COLORS.vfd, Feeder: REPORT_COLORS.feeder };
  const OLD = [REPORT_COLORS.header, REPORT_COLORS.within, REPORT_COLORS.over, REPORT_COLORS.incomer, REPORT_COLORS.coupler];

  it("W, S, DOL, VFD and Feeder each get their own pastel, clearly different from the existing shades and from each other", () => {
    const colours = Object.values(NEW);
    expect(new Set(colours).size).toBe(5);
    for (const c of colours) {
      expect(OLD).not.toContain(c);
      expect(lightness(c), c).toBeGreaterThanOrEqual(0.8); // pastel = light
      for (const o of OLD) expect(dE(c, o), `${c} vs ${o}`).toBeGreaterThan(18);
    }
    for (let i = 0; i < colours.length; i++) for (let j = i + 1; j < colours.length; j++) expect(dE(colours[i], colours[j]), `${colours[i]} vs ${colours[j]}`).toBeGreaterThan(18);
  });

  it("kindColor maps the values (case and space insensitive) and nothing else", () => {
    expect(kindColor("W")).toBe(NEW.W);
    expect(kindColor(" s ")).toBe(NEW.S);
    expect(kindColor("DOL")).toBe(NEW.DOL);
    expect(kindColor("DOL motor")).toBe(NEW.DOL);
    expect(kindColor("vfd")).toBe(NEW.VFD);
    expect(kindColor("VFD motor")).toBe(NEW.VFD);
    expect(kindColor("Feeder")).toBe(NEW.Feeder);
    expect(kindColor("Crusher")).toBeNull();
    expect(kindColor("")).toBeNull();
  });

  it("W/S and Treated-as cells are shaded on the Loads sheet, the bus layout and the original sheet", async () => {
    const { sheets, result, raw } = await build();
    const bgOf = (kind: string) => NEW[kind as keyof typeof NEW];
    // Loads sheet: W/S column and Treated as column
    const loads = sheets.find((x) => x.sheet === "Loads")!.data;
    const header = loads[0].map((c) => String(val(c)));
    const ws = header.indexOf("W/S"), treated = header.indexOf("Treated as");
    const body = loads.slice(1, loads.findIndex((r) => val(r[1]) === "Total"));
    expect(body).toHaveLength(result.loads.length);
    body.forEach((r, i) => {
      const l = result.loads[i];
      expect(st(r[ws])!.backgroundColor, `W/S row ${l.rowNumber}`).toBe(bgOf(l.status!));
      expect(st(r[ws])!.align).toBe("center");
      const kind = l.kind === "feeder" ? "Feeder" : l.type;
      expect(st(r[treated])!.backgroundColor, `treated as row ${l.rowNumber}`).toBe(bgOf(kind));
    });
    expect(new Set(body.map((r) => st(r[ws])!.backgroundColor))).toEqual(new Set([NEW.W])); // only working loads are counted
    expect(new Set(body.map((r) => st(r[treated])!.backgroundColor))).toEqual(new Set([NEW.DOL, NEW.VFD, NEW.Feeder]));
    // bus layout: outgoing-load rows carry the same shading
    const three = sheets.find((x) => x.sheet === "3 Transformers")!.data;
    const outgoing = three.filter((r) => val(r[0]) === "Outgoing load");
    expect(outgoing.length).toBe(result.loads.length);
    for (const r of outgoing) {
      expect(st(r[4])!.backgroundColor).toBe(NEW[String(val(r[4])) as "W" | "S"]); // W/S column
      expect(st(r[6])!.backgroundColor).toBeDefined(); // treated as
    }
    // legend row lists the five shades
    const legend = three.find((r) => val(r[0]) === "Incomer" && val(r[1]) === "Bus coupler")!;
    expect(legend.map((c) => st(c)?.backgroundColor)).toEqual([REPORT_COLORS.incomer, REPORT_COLORS.coupler, undefined, NEW.W, NEW.S, NEW.DOL, NEW.VFD, NEW.Feeder, REPORT_COLORS.spare]);
    // original sheet: the W / S / DOL / VFD / Feeder cells are shaded, everything else is not
    const orig = sheets.find((x) => x.sheet === "Original load list")!.data;
    let shaded = 0;
    orig.slice(1).forEach((r, ri) => r.forEach((c) => {
      const v = val(c);
      const want = typeof v === "string" ? kindColor(v) : null;
      expect(st(c)?.backgroundColor ?? null, `original ${ri + 2}`).toBe(want);
      if (want) shaded++;
    }));
    expect(shaded).toBe((raw.length - 1) * 2); // every data row has a W/S and a type value
  });

  /** Walk every numeric cell with the header (table) or label (key/value block) that describes it. */
  const numericCells = (sheet: { data: ReportCell[][] }) => {
    const out: { value: number; format?: string; align?: string; head: string }[] = [];
    let heads: string[] | null = null;
    for (const row of sheet.data) {
      const orange = row.length > 0 && row.every((c) => st(c)?.backgroundColor === REPORT_COLORS.header);
      if (orange) {
        const texts = row.map((c) => String(val(c) ?? ""));
        heads = texts.filter((t) => t !== "").length >= 3 ? texts : null; // 2-cell orange rows are block titles
        continue;
      }
      row.forEach((c, j) => {
        const sc = st(c);
        if (sc && typeof sc.value === "number") out.push({ value: sc.value, format: sc.format, align: sc.align, head: heads ? (heads[j] ?? "") : `${val(row[0])}` });
      });
    }
    return out;
  };

  it("only percentages have a decimal place; every other number is a whole number (power factor keeps two decimals)", async () => {
    const { sheets } = await build();
    let pctCells = 0, intCells = 0, pfCells = 0;
    for (const sh of sheets.filter((x) => x.sheet !== "Original load list")) {
      for (const c of numericCells(sh)) {
        const isPct = /%/.test(c.head);
        const isPf = /power factor|diversity factor/i.test(c.head);
        if (isPct) { expect(c.format, `${sh.sheet}: ${c.head}`).toBe("0.0"); pctCells++; }
        else if (isPf) { expect(c.format, `${sh.sheet}: ${c.head}`).toBe("0.00"); pfCells++; }
        else {
          // whole-number display: "#,##0", or no format at all for plain integers (counts, ratings, row numbers)
          if (c.format === undefined) expect(Number.isInteger(c.value), `${sh.sheet}: ${c.head} = ${c.value}`).toBe(true);
          else expect(c.format, `${sh.sheet}: ${c.head}`).toBe("#,##0");
          intCells++;
        }
      }
    }
    expect(pctCells).toBeGreaterThan(20);
    expect(intCells).toBeGreaterThan(100);
    expect(pfCells).toBeGreaterThan(5);
  });

  it("the Summary states the voltage used to convert amps to kVA (433 V for 433 V, 690 V for 690 V)", async () => {
    const { sheets } = await build();
    const row = sheets[0].data.find((r) => val(r[0]) === "Voltage used to convert amps to kVA (V)")!;
    expect(val(row[1])).toBe(433);
    const raw = await readSheet(buf);
    const r690 = sizeTransformers(parseLoadSheet(raw).rows, 690, BOTH);
    const s690 = buildReportSheets(r690, { transformerLabel: "6600 / 690 V", options: BOTH, includePower: true })[0].data;
    expect(val(s690.find((r) => val(r[0]) === "Voltage used to convert amps to kVA (V)")![1])).toBe(690);
  });

  it("the display is rounded but the values are not (kVA keeps its decimals underneath)", async () => {
    const { sheets, result } = await build();
    const summary = numericCells(sheets[0]);
    const total = summary.find((c) => c.head === "Total kVA used for sizing")!;
    expect(total.format).toBe("#,##0");
    expect(total.value).toBeCloseTo(result.totalKva, 9);
    expect(Number.isInteger(total.value)).toBe(false);
  });

  it("all numbers and percentages, and the W/S columns, are centred; text is left-aligned", async () => {
    const { sheets } = await build();
    let numbers = 0;
    for (const sh of sheets) {
      for (const row of sh.data) {
        row.forEach((c) => {
          const sc = st(c);
          if (!sc) return;
          if (typeof sc.value === "number") { expect(sc.align, `${sh.sheet}: ${sc.value}`).toBe("center"); numbers++; }
        });
      }
    }
    expect(numbers).toBeGreaterThan(300);
    // W/S column on Loads (header "W/S") is centred, including blank cells
    const loads = sheets.find((x) => x.sheet === "Loads")!.data;
    const ws = loads[0].map((c) => String(val(c))).indexOf("W/S");
    const upToTotal = loads.slice(1, loads.findIndex((r) => val(r[1]) === "Total"));
    for (const r of upToTotal) expect(st(r[ws])!.align).toBe("center");
    // descriptions stay left-aligned
    const desc = loads[0].map((c) => String(val(c))).indexOf("Description");
    for (const r of upToTotal) expect(st(r[desc])!.align).toBe("left");
  });

  it("the real .xlsx file carries the pastel fills, the number formats and centring, and no author metadata", async () => {
    const { result, o } = await build();
    const blob = await buildReportBlob(result, o);
    const parts = unzipSync(new Uint8Array(await blob.arrayBuffer()));
    const styles = strFromU8(parts["xl/styles.xml"]);
    for (const c of [...Object.values(NEW), ...OLD]) expect(styles.toUpperCase(), c).toContain(`FF${c.slice(1).toUpperCase()}`);
    const formats = [...styles.matchAll(/formatCode="([^"]*)"/g)].map((m) => m[1]);
    expect(formats).toEqual(expect.arrayContaining(["0.0", "#,##0", "0.00"]));
    expect(formats).not.toContain("#,##0.0");
    expect(styles).toContain('horizontal="center"');
    // no document properties at all (no creator / last modified by), and no author-like text anywhere
    expect(Object.keys(parts).filter((n) => n.startsWith("docProps"))).toEqual([]);
    for (const [name, bytes] of Object.entries(parts)) {
      expect(strFromU8(bytes), name).not.toMatch(/dc:creator|lastModifiedBy|Puneet|Narayan/i);
    }
  });
});

describe("Excel report: spare feeders (one of each type and rating, highlighted, not in the calculation)", () => {
  const bg = (c: ReportCell) => (c !== null && typeof c === "object" ? c.backgroundColor : undefined);
  const val = (c: ReportCell) => (c !== null && typeof c === "object" ? c.value : c);
  const build = async (opts = BOTH, includePower = true) => {
    const raw = await readSheet(buf);
    const result = sizeTransformers(parseLoadSheet(raw).rows, 433, opts);
    return { result, sheets: buildReportSheets(result, { transformerLabel: "6600 / 433 V", options: opts, includePower, originalSheet: raw }) };
  };
  const GOLD = REPORT_COLORS.spare;

  it("the spare colour is a new one, clearly different from every other shade in the report", () => {
    expect(GOLD).toBe("#FFD966");
    const rgb = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
    const lab = (h: string) => {
      const [r, g, b] = rgb(h).map((c) => c / 255).map((c) => (c > 0.04045 ? ((c + 0.055) / 1.055) ** 2.4 : c / 12.92));
      const f = (t: number) => (t > 0.008856 ? Math.cbrt(t) : 7.787 * t + 16 / 116);
      const X = f((0.4124 * r + 0.3576 * g + 0.1805 * b) / 0.95047), Y = f(0.2126 * r + 0.7152 * g + 0.0722 * b), Z = f((0.0193 * r + 0.1192 * g + 0.9505 * b) / 1.08883);
      return [116 * Y - 16, 500 * (X - Y), 200 * (Y - Z)];
    };
    const dE = (a: string, b: string) => Math.hypot(...lab(a).map((v, i) => v - lab(b)[i]));
    const others = Object.entries(REPORT_COLORS).filter(([k]) => !["spare", "border"].includes(k)).map(([, v]) => v);
    expect(others).toHaveLength(10);
    for (const o of others) expect(dE(GOLD, o), `vs ${o}`).toBeGreaterThan(30);
  });

  it("Loads sheet: below the Total row, one gold spare row per type and rating, with no kW / kvar / kVA figures", async () => {
    const { result, sheets } = await build();
    const loads = sheets.find((x) => x.sheet === "Loads")!.data;
    const totalIdx = loads.findIndex((r) => val(r[1]) === "Total");
    const after = loads.slice(totalIdx + 1);
    expect(after[0]).toEqual([]);
    expect(String(val(after[1][0]))).toContain("Spare feeders: one of each type and rating (not included in the calculation)");
    const header = after[2].map((c) => String(val(c)));
    expect(header).toEqual(loads[0].map((c) => String(val(c)))); // same columns as the loads
    const spareRows = after.slice(3);
    expect(spareRows).toHaveLength(result.spares.length);
    expect(spareRows).toHaveLength(12);
    spareRows.forEach((r, i) => {
      const ws = header.indexOf("W/S");
      expect(r.map((c, k) => (bg(c) === GOLD ? k : -1)).filter((k) => k >= 0), `row ${i}`).toEqual([ws]); // only the SP cell is gold
      expect(val(r[header.indexOf("Description")])).toBe(result.spares[i].description);
      expect(val(r[header.indexOf("Basis")])).toBe("Spare: not included in the calculation");
      for (const col of ["Row", "Duty/standby group", "kW (input)", "kvar", "Power factor", "kVA"]) expect(val(r[header.indexOf(col)]), col).toBe("");
      expect(val(r[header.indexOf("W/S")])).toBe("SP");
      expect(String(val(r[header.indexOf("Description")]))).toMatch(/ \(Spare\)$/);
      const sp = result.spares[i];
      expect(val(r[header.indexOf("Rated kW (motor)")])).toBe(sp.unit === "kW" ? sp.rating : "");
    });
    // the Total row is the total of the loads only
    const kva = header.indexOf("kVA");
    const loadRows = loads.slice(1, totalIdx);
    expect(Number(val(loads[totalIdx][kva]))).toBeCloseTo(result.totalKva, 6);
    expect(loadRows.reduce((a, r) => a + Number(val(r[kva])), 0)).toBeCloseTo(result.totalKva, 6);
  });

  it("Summary says how many spares were added; the totals are those of the loads", async () => {
    const { result, sheets } = await build();
    const sum = sheets[0].data;
    expect(val(sum.find((r) => val(r[0]) === "Spare feeders added (not in the calculation)")![1])).toBe(12);
    expect(val(sum.find((r) => val(r[0]) === "Total kVA used for sizing")![1])).toBeCloseTo(result.totalKva, 9);
    expect(val(sum.find((r) => val(r[0]) === "Loads included")![1])).toBe(result.loads.length);
  });

  it("bus layout: each spare sits right after a load of the same type and rating, one of each in total, gold end to end", async () => {
    const { result, sheets } = await build();
    for (const name of ["2 Transformers", "3 Transformers"]) {
      const rows = sheets.find((x) => x.sheet === name)!.data;
      const spareRows = rows.filter((r) => val(r[0]) === "Spare feeder");
      expect(spareRows, name).toHaveLength(result.spares.length);
      expect(new Set(spareRows.map((r) => val(r[3]))).size).toBe(result.spares.length); // each description once
      rows.forEach((r, i) => {
        if (val(r[0]) !== "Spare feeder") return;
        expect(r.map((c, k) => (bg(c) === GOLD ? k : -1)).filter((k) => k >= 0)).toEqual([4]); // only the SP cell (W/S column)
        const prev = rows[i - 1];
        expect(val(prev[0]), `${name}: spare ${val(r[3])} follows a load`).toBe("Outgoing load");
        expect(val(prev[1])).toBe(val(r[1])); // same section
        // same type and rating as the load before it
        const sp = result.spares.find((x) => x.description === val(r[3]))!;
        const twin = result.loads.find((l) => l.rowNumber === Number(val(prev[2])))!;
        expect(spareKey(twin)).toBe(sp.key);
      });
    }
  });

  it("the spares do not change the bus layout figures: incomer rows still equal the sum of their outgoing loads", async () => {
    const { sheets, result } = await build();
    const rows = sheets.find((x) => x.sheet === "3 Transformers")!.data;
    const h = rows.findIndex((r) => val(r[0]) === "Element" && val(r[1]) === "Name");
    const body: ReportCell[][] = [];
    for (let i = h + 1; i < rows.length && rows[i].length > 0; i++) body.push(rows[i]);
    const kvaCol = body[0].length - 3;
    const sums: Record<string, number> = {};
    const incomer: Record<string, number> = {};
    for (const r of body) {
      const kind = val(r[0]);
      if (kind === "Outgoing load") sums[String(val(r[1]))] = (sums[String(val(r[1]))] ?? 0) + Number(val(r[kvaCol]));
      if (kind === "Incomer") incomer[String(val(r[1]))] = Number(val(r[kvaCol]));
      if (kind === "Spare feeder") expect(val(r[kvaCol])).toBe(""); // nothing to add up
    }
    for (const t of Object.keys(incomer)) expect(sums[t]).toBeCloseTo(incomer[t], 6);
    expect(Object.values(incomer).reduce((a, b) => a + b, 0)).toBeCloseTo(result.totalKva, 6);
  });

  it("several sets: still one spare of each, shared across the sets", async () => {
    const rows: RawLoadRow[] = Array.from({ length: 60 }, (_, i) => ({
      rowNumber: i + 2, description: `F${i}`, type: i % 2 ? "DOL" : "Feeder", voltageV: 433, status: "W" as const,
      ...(i % 2 ? { kw: 100 + (i % 6) * 10 } : { amps: 250 + (i % 3) * 50 }),
    }));
    const opts = { ...DEFAULT_OPTIONS, ratingsKva: [2000], counts: [5] };
    const result = sizeTransformers(rows, 433, opts);
    const sheet = buildReportSheets(result, { transformerLabel: "x", options: opts, includePower: false }).find((x) => x.sheet === "5 Transformers")!;
    const spareRows = sheet.data.filter((r) => val(r[0]) === "Spare feeder");
    expect(spareRows).toHaveLength(result.spares.length);
    expect(result.spares.length).toBe(3 + 3); // DOL 110 / 130 / 150 kW (odd rows) + feeders 250 / 300 / 350 A (even rows)
  });

  it("the real .xlsx carries the gold fill and the spare rows", async () => {
    const { result, sheets } = await build();
    const blob = await buildReportBlob(result, { transformerLabel: "x", options: BOTH, includePower: true, originalSheet: await readSheet(buf) });
    const parts = unzipSync(new Uint8Array(await blob.arrayBuffer()));
    expect(strFromU8(parts["xl/styles.xml"]).toUpperCase()).toContain("FFFFD966");
    const read = Object.fromEntries((await readExcelFile(await blob.arrayBuffer())).map((x) => [x.sheet, x.data]));
    expect(read["3 Transformers"].filter((r) => r[0] === "Spare feeder")).toHaveLength(12);
    expect(read["Loads"].some((r) => String(r[0]).startsWith("Spare feeders: one of each type and rating"))).toBe(true);
    expect(sheets.length).toBeGreaterThan(3);
  });

  it("lists them in the bus layout without a row number, and shows the combination total", async () => {
    const { sheets } = await build({ ...DEFAULT_OPTIONS, ratingsKva: [2000], mccFeeders: STANDARD_MCC_FEEDERS });
    const byName = Object.fromEntries(sheets.map((x) => [x.sheet, x]));
    const sum = byName["Summary"].data;
    expect(sum.some((r) => String(val(r[0])).startsWith("PMCC/IMCC feeders added"))).toBe(true);
    const combo = Object.entries(byName).find(([k]) => /Transformers$/.test(k))![1].data;
    const mcc = combo.filter((r) => String(val(r[3]) ?? "").startsWith("PMCC/IMCC feeder"));
    expect(mcc.length).toBeGreaterThanOrEqual(28);
    expect(mcc.every((r) => val(r[2]) === "")).toBe(true);
    expect(combo.some((r) => String(val(r[0])).startsWith("Total load "))).toBe(true);
  });

  it("a Current (A) column follows kVA on the Loads sheet and in the bus layout (incomers too, whole numbers, centred)", async () => {
    const { sheets, result } = await build();
    const loads = sheets.find((x) => x.sheet === "Loads")!.data;
    const h = loads[0].map((c) => String(val(c)));
    expect(h.indexOf("Current (A)")).toBe(h.indexOf("kVA") + 1);
    expect(h.indexOf("Basis")).toBe(h.indexOf("Current (A)") + 1);
    result.loads.forEach((l, i) => {
      const c = loads[i + 1][h.indexOf("Current (A)")] as { value: number; format: string; align: string };
      expect(c.value).toBeCloseTo(l.current, 9);
      expect([c.format, c.align]).toEqual(["#,##0", "center"]);
    });
    const rows = sheets.find((x) => x.sheet === "3 Transformers")!.data;
    const lh = rows.findIndex((r) => val(r[0]) === "Element" && val(r[1]) === "Name");
    const lay = rows[lh].map((c) => String(val(c)));
    const ci = lay.indexOf("Current (A)");
    expect(ci).toBe(lay.indexOf("kVA") + 1);
    const inc = rows.slice(lh + 1).find((r) => val(r[0]) === "Incomer")!;
    expect(Number(val(inc[ci]))).toBeCloseTo((Number(val(inc[lay.indexOf("kVA")])) * 1000) / (Math.sqrt(3) * 433), 6);
  });
});

