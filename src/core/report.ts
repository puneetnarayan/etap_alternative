import writeExcelFile from "write-excel-file/universal";
import type { ComboResult, SizingOptions, SizingResult } from "./transformer-sizing";

type Cell = null | string | number | { value: string | number; fontWeight?: "bold"; format?: string };
type Row = Cell[];

export interface ReportSheet {
  sheet: string;
  data: Row[];
  columns?: { width: number }[];
  stickyRowsCount?: number;
}

export interface ReportOptions {
  /** e.g. "6600 / 433 V" */
  transformerLabel: string;
  fileName?: string;
  /** The assumptions the result was computed with. */
  options: SizingOptions;
  /** Include input kW, kvar and power factor columns. */
  includePower: boolean;
}

const b = (v: string): Cell => ({ value: v, fontWeight: "bold" });
const n = (v: number, format = "#,##0.0"): Cell => ({ value: v, format });
const heads = (...h: string[]): Row => h.map(b);

function loadHeaders(includePower: boolean, withSection = false): Row {
  return heads(
    ...(withSection ? ["Transformer"] : []),
    "Row", "Description", "W/S", "Duty/standby group", "Treated as", "Rated kW (motor)",
    ...(includePower ? ["kW (input)", "kvar", "Power factor"] : []),
    "kVA", "Basis",
  );
}

function loadCells(l: SizingResult["loads"][number], includePower: boolean, section?: string): Row {
  return [
    ...(section ? [section] : []),
    l.rowNumber, l.description, l.status ?? "", l.group ?? "", l.kind === "feeder" ? "Feeder" : `${l.type} motor`,
    l.ratedKw === undefined ? "" : n(l.ratedKw),
    ...(includePower ? [n(l.kw), n(l.kvar), n(l.powerFactor, "0.00")] : []),
    n(l.kva), l.basis,
  ];
}

function summarySheet(r: SizingResult, o: ReportOptions): ReportSheet {
  const op = o.options;
  const rec = r.recommended;
  const data: Row[] = [
    [b("Transformer sizing report")],
    ["Transformer", o.transformerLabel],
    ...(o.fileName ? [["Source file", o.fileName] as Row] : []),
    [],
    [b("Assumptions")],
    ["Motor efficiency (%)", n(op.efficiency * 100, "0.0")],
    ["Power factor, DOL motors", n(op.powerFactor.DOL, "0.00")],
    ["Power factor, VFD motors", n(op.powerFactor.VFD, "0.00")],
    ["Power factor, feeders", n(op.powerFactor.FEEDER, "0.00")],
    ["Max transformer loading, normal running (%)", n(op.normalLimit * 100, "0")],
    ["Max transformer loading, one transformer out (%)", n(op.contingencyLimit * 100, "0")],
    ["Ratings considered (kVA)", [...op.ratingsKva].sort((x, y) => x - y).join(", ")],
    ["Total kVA basis", "Arithmetic sum of each load's kVA"],
    ["Load factor", "All loads at 100 % (working W and standby S)"],
    [],
    [b("Total load")],
    ["Loads included", r.loads.length],
    ["Working (W) / standby (S)", `${r.loads.filter((l) => l.status === "W").length} W / ${r.loads.filter((l) => l.status === "S").length} S`],
    ["Total kVA used for sizing", n(r.totalKva)],
    ...(o.includePower
      ? ([
          ["Total kW (input)", n(r.totals.kw)],
          ["Total kvar", n(r.totals.kvar)],
          ["Overall power factor (kW ÷ √(kW² + kvar²))", n(r.totals.powerFactor, "0.00")],
        ] as Row[])
      : []),
    [],
    [b("Recommendation")],
    rec
      ? ["Recommended", `${rec.count} × ${rec.ratingKva} kVA`]
      : ["Recommended", "No configuration fits the ratings considered"],
    ...(rec ? ([["Incomers", rec.count], ["Bus couplers", rec.count - 1]] as Row[]) : []),
    [],
    [b("By number of transformers")],
    heads("Transformers", "Rating (kVA)", "Fits", "Normal loading (%)", "One out loading (%)"),
    ...r.combos.map((c): Row => [
      c.count, c.chosen.ratingKva, c.fits ? "Yes" : "No", n(c.chosen.normalPct), n(c.chosen.contingencyPct),
    ]),
  ];
  return { sheet: "Summary", data, columns: [{ width: 48 }, { width: 34 }, { width: 14 }, { width: 20 }, { width: 20 }] };
}

function loadsSheet(r: SizingResult, o: ReportOptions): ReportSheet {
  const rows: Row[] = [loadHeaders(o.includePower), ...r.loads.map((l) => loadCells(l, o.includePower))];
  rows.push([null, b("Total"), null, null, null, null, ...(o.includePower ? [n(r.totals.kw), n(r.totals.kvar), n(r.totals.powerFactor, "0.00")] : []), n(r.totalKva)]);
  return {
    sheet: "Loads",
    data: rows,
    columns: [{ width: 7 }, { width: 38 }, { width: 6 }, { width: 18 }, { width: 14 }, { width: 16 }, ...(o.includePower ? [{ width: 12 }, { width: 12 }, { width: 12 }] : []), { width: 12 }, { width: 44 }],
    stickyRowsCount: 1,
  };
}

function comboSheet(c: ComboResult, r: SizingResult, o: ReportOptions): ReportSheet {
  const op = o.options;
  const data: Row[] = [
    [b(`${c.count} × ${c.chosen.ratingKva} kVA transformers, ${o.transformerLabel}`)],
    [c.fits ? "Fits within both limits" : b("DOES NOT FIT within both limits")],
    ...(c.added ? ([["Added because the selected number of transformers did not fit."]] as Row[]) : []),
    [`${c.incomers} incomers, ${c.couplers} bus coupler${c.couplers === 1 ? "" : "s"}; bus sections in a line: ${c.sections.map((s) => s.name).join(" – coupler – ")}`],
    [],
    [b("Ratings considered")],
    heads("Rating (kVA)", "Installed kVA", "Normal loading (%)", "One out loading (%)", "Result"),
    ...c.candidates.map((x): Row => [
      x.ratingKva, x.installedKva, n(x.normalPct), n(x.contingencyPct),
      x.fits ? "Fits" : `Over limit (${[!x.normalOk && "normal", !x.contingencyOk && "one out"].filter(Boolean).join(" + ")})`,
    ]),
    [],
    [b("Load distribution (normal running)")],
    heads("Transformer", "Loads", ...(o.includePower ? ["kW (input)", "kvar"] : []), "kVA", "Loading (%)", `Within ${op.normalLimit * 100} %`),
    ...c.sections.map((s): Row => [
      s.name, s.loads.length, ...(o.includePower ? [n(s.kw), n(s.kvar)] : []), n(s.kva), n(s.loadingPct), s.withinNormalLimit ? "Yes" : "No",
    ]),
    [],
    [b(`One transformer out (each remaining unit limited to ${op.contingencyLimit * 100} %)`)],
    heads("Transformer out", "Load moved (kVA)", "Each remaining (kVA)", "Loading (%)", "Result"),
    ...c.oneOut.map((s): Row => [s.outName, n(s.movedKva), n(s.eachKva), n(s.loadingPct), s.ok ? "OK" : "Over limit"]),
    [],
    ...(c.notes.length ? ([[b("Notes")], ...c.notes.map((t): Row => [t]), []] as Row[]) : []),
    [b("Loads by transformer")],
    loadHeaders(o.includePower, true),
    ...c.sections.flatMap((s) => s.loads.map((l) => loadCells(l, o.includePower, s.name))),
  ];
  return {
    sheet: `${c.count} Transformers`,
    data,
    columns: [{ width: 20 }, { width: 22 }, { width: 38 }, { width: 8 }, { width: 22 }, { width: 18 }, { width: 18 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 44 }],
  };
}

function issuesSheet(r: SizingResult): ReportSheet {
  const label = { error: "Error (not included)", excluded: "Excluded", warning: "Warning" } as const;
  return {
    sheet: "Issues",
    data: [
      heads("Row", "Description", "Type", "Message"),
      ...r.issues.map((i): Row => [i.rowNumber, i.description, label[i.severity], i.message]),
    ],
    columns: [{ width: 7 }, { width: 38 }, { width: 22 }, { width: 80 }],
    stickyRowsCount: 1,
  };
}

export function buildReportSheets(result: SizingResult, o: ReportOptions): ReportSheet[] {
  return [
    summarySheet(result, o),
    loadsSheet(result, o),
    ...result.combos.map((c) => comboSheet(c, result, o)),
    ...(result.issues.length > 0 ? [issuesSheet(result)] : []),
  ];
}

export async function buildReportBlob(result: SizingResult, o: ReportOptions): Promise<Blob> {
  return writeExcelFile(buildReportSheets(result, o)).toBlob();
}
