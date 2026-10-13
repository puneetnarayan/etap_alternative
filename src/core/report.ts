import writeExcelFile from "write-excel-file/universal";
import { fmtInt, fmtPct } from "./format";
import {
  displayCombos, failureReason, spareKey,
  type ComboResult, type SetResult, type SizingOptions, type SizingResult, type SpareFeeder,
} from "./transformer-sizing";

/** Colours used in the Excel report. */
export const REPORT_COLORS = {
  header: "#FFE0B2", // light orange
  within: "#C6EFCE", // light green
  over: "#FFC7CE", // light red
  incomer: "#80DEEA", // cyan
  coupler: "#FFF9C4", // light yellow
  // Pastels for the load kinds, each clearly different from the shades above and from each other.
  working: "#DEB1FB", // W: lavender
  standby: "#FBB1DE", // S: pink
  dol: "#BACEF3", // DOL: periwinkle blue
  vfd: "#CBFBB1", // VFD: lime
  feeder: "#D8D4CA", // Feeder: sand
  // Spare feeders (one of each type and rating, not in the calculation): a stronger gold so they stand out.
  spare: "#FFD966",
  border: "#BFBFBF",
} as const;

type Style = {
  fontWeight?: "bold";
  fontSize?: number;
  backgroundColor?: string;
  textColor?: string;
  align?: "left" | "center" | "right";
  alignVertical?: "top" | "center" | "bottom";
  wrap?: boolean;
  borderColor?: string;
  borderStyle?: "thin";
  format?: string;
  /** Merge this cell with the next `columnSpan - 1` cells in the row (which must be null). */
  columnSpan?: number;
};
export type ReportCell = null | string | number | boolean | ({ value: string | number | boolean } & Style);
type Row = ReportCell[];

export interface ReportSheet {
  sheet: string;
  data: Row[];
  columns?: { width: number }[];
  stickyRowsCount?: number;
  /** Print orientation. */
  orientation?: "landscape";
}

export interface ReportOptions {
  /** e.g. "6600 / 433 V" */
  transformerLabel: string;
  fileName?: string;
  /** The assumptions the result was computed with. */
  options: SizingOptions;
  /** Include input kW, kvar and power factor columns. */
  includePower: boolean;
  /** The load list exactly as read from the uploaded file (rows of cell values); included as its own sheet. */
  originalSheet?: readonly (readonly unknown[])[];
}

/* ------------------------------ cell helpers ------------------------------ */

const BORDER: Style = { borderStyle: "thin", borderColor: REPORT_COLORS.border };

/** Sheet title. */
const title = (v: string): ReportCell => ({ value: v, fontWeight: "bold", fontSize: 14 });
/** Bold label above a table. */
const label = (v: string): ReportCell => ({ value: v, fontWeight: "bold", fontSize: 12 });
/** Table header cell: light orange. */
const head = (v: string): ReportCell => ({
  value: v, fontWeight: "bold", backgroundColor: REPORT_COLORS.header, align: "center", alignVertical: "center", ...BORDER,
});
const heads = (...h: string[]): Row => h.map(head);
/** Plain table cell: numbers are centred, text is left-aligned. */
const cell = (v: string | number, align: "left" | "right" | "center" = typeof v === "number" ? "center" : "left"): ReportCell =>
  typeof v === "number" ? { value: v, format: "#,##0", align, ...BORDER } : { value: v, align, ...BORDER };
/**
 * Number cell, centred. Whole numbers are shown without decimals ("#,##0"); only percentages
 * (see `pct`, `limited`) have one decimal. Power factor is a ratio below 1 and keeps two decimals ("0.00").
 */
const n = (v: number, format = "#,##0"): ReportCell => ({ value: v, format, align: "center", ...BORDER });
/** A percentage, centred, one decimal. */
const pct = (v: number): ReportCell => n(v, "0.0");
/** Table cell shaded light green when within its limit, light red when over. */
const fill = (within: boolean) => (within ? REPORT_COLORS.within : REPORT_COLORS.over);
const limited = (v: number, limitPct: number): ReportCell =>
  ({ value: v, format: "0.0", align: "center", backgroundColor: fill(v <= limitPct + 1e-9), ...BORDER });
const verdict = (text: string, within: boolean): ReportCell =>
  ({ value: text, align: "center", backgroundColor: fill(within), ...BORDER });

/** Pastel shade for a W / S / DOL / VFD / Feeder value (null for anything else). */
export function kindColor(value: string): string | null {
  switch (value.trim().toLowerCase()) {
    case "w": return REPORT_COLORS.working;
    case "s": return REPORT_COLORS.standby;
    case "dol": case "dol motor": return REPORT_COLORS.dol;
    case "vfd": case "vfd motor": return REPORT_COLORS.vfd;
    case "feeder": return REPORT_COLORS.feeder;
    default: return null;
  }
}
/** W / S cell: centred and shaded. */
const statusCell = (v: string): ReportCell => {
  const bg = kindColor(v);
  return { value: v, align: "center", ...(bg ? { backgroundColor: bg } : {}), ...BORDER };
};
/** "Treated as" cell (DOL motor / VFD motor / Feeder), shaded by kind. */
const treatedAs = (l: { kind: string; type: string }): ReportCell => {
  const text = l.kind === "feeder" ? "Feeder" : `${l.type} motor`;
  const bg = kindColor(text);
  return { value: text, align: "left", ...(bg ? { backgroundColor: bg } : {}), ...BORDER };
};

/** A short banner (green when within limits, red otherwise) merged across `span` columns. */
const banner = (text: string, within: boolean, span = 5): Row => [
  { value: text, fontWeight: "bold", align: "left", backgroundColor: fill(within), columnSpan: span, ...BORDER },
  ...Array<ReportCell>(span - 1).fill(null),
];
/** A short neutral note merged across `span` columns. */
const noteBanner = (text: string, span = 5): Row => [
  { value: text, align: "left", columnSpan: span, ...BORDER },
  ...Array<ReportCell>(span - 1).fill(null),
];

/** A table-style header spanning `span` columns for a key/value block on the Summary sheet. */
const blockHeader = (v: string, span = 2): Row => [head(v), ...Array.from({ length: span - 1 }, () => head(""))];

const textOf = (c: ReportCell): string => (c === null || c === undefined ? "" : typeof c === "object" ? String(c.value) : String(c));

/** Column widths from the content. Rows with a single filled cell are free text and are left to overflow. */
function autoColumns(rows: Row[], min = 8, max = 60): { width: number }[] {
  const cols = Math.max(0, ...rows.map((r) => r.length));
  const widths = Array<number>(cols).fill(min);
  for (const r of rows) {
    if (r.filter((c) => textOf(c) !== "").length <= 1) continue;
    r.forEach((c, i) => {
      const raw = typeof c === "object" && c !== null ? c.value : c;
      const len = typeof raw === "number" ? 11 : textOf(c).length + 2;
      widths[i] = Math.min(max, Math.max(widths[i], len));
    });
  }
  return widths.map((width) => ({ width }));
}

/* ------------------------------- load tables ------------------------------ */

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
    ...(section ? [cell(section)] : []),
    cell(l.rowNumber), cell(l.description), statusCell(l.status ?? ""), cell(l.group ?? ""),
    treatedAs(l),
    l.ratedKw === undefined ? cell("") : n(l.ratedKw),
    ...(includePower ? [n(l.kw), n(l.kvar), n(l.powerFactor, "0.00")] : []),
    n(l.kva), cell(l.basis),
  ];
}

/**
 * A spare feeder row: same columns as a load row, shaded gold end to end. It shows the type and
 * rating only; there are no kW / kvar / kVA values because a spare is not part of any total.
 */
function spareCells(sp: SpareFeeder, includePower: boolean, lead: Row): Row {
  return tint([
    ...lead,
    cell(""), cell(sp.description), cell(""), cell(""), cell(sp.kind === "feeder" ? "Feeder" : `${sp.type} motor`),
    sp.unit === "kW" ? n(sp.rating) : cell(""),
    ...(includePower ? [cell(""), cell(""), cell("")] : []),
    cell(""), cell("Spare: not included in the calculation"),
  ], REPORT_COLORS.spare);
}

/* --------------------------------- sheets --------------------------------- */

function summarySheet(r: SizingResult, o: ReportOptions): ReportSheet {
  const op = o.options;
  const rec = r.recommended;
  const normalPct = op.normalLimit * 100;
  const contPct = op.contingencyLimit * 100;
  const kv = (k: string, v: ReportCell): Row => [cell(k), v];
  const data: Row[] = [
    [title("Transformer sizing report")],
    [],
    kv("Transformer", cell(o.transformerLabel)),
    ...(o.fileName ? [kv("Source file", cell(o.fileName))] : []),
    [],
    blockHeader("Assumptions"),
    kv("Motor efficiency (%)", pct(op.efficiency * 100)),
    kv("Power factor, DOL motors", n(op.powerFactor.DOL, "0.00")),
    kv("Power factor, VFD motors", n(op.powerFactor.VFD, "0.00")),
    kv("Power factor, feeders", n(op.powerFactor.FEEDER, "0.00")),
    kv("Max transformer loading, normal running (%)", pct(normalPct)),
    kv("Max transformer loading, one transformer out (%)", pct(contPct)),
    kv("Ratings considered (kVA)", cell([...op.ratingsKva].sort((x, y) => x - y).join(", "))),
    kv("Voltage used to convert amps to kVA (V)", n(op.calculationVoltageV[r.selectedVoltageV] ?? r.selectedVoltageV)),
    kv("Total kVA basis", cell("Arithmetic sum of each load's kVA")),
    kv("Load factor", cell("All loads at 100 % (working W and standby S)")),
    [],
    blockHeader("Total load"),
    kv("Loads included", cell(r.loads.length)),
    kv("Working (W) / standby (S)", cell(`${r.loads.filter((l) => l.status === "W").length} W / ${r.loads.filter((l) => l.status === "S").length} S`)),
    kv("Spare feeders added (not in the calculation)", cell(r.spares.length)),
    kv("Total kVA used for sizing", n(r.totalKva)),
    ...(o.includePower
      ? [
          kv("Total kW (input)", n(r.totals.kw)),
          kv("Total kvar", n(r.totals.kvar)),
          kv("Overall power factor (kW ÷ √(kW² + kvar²))", n(r.totals.powerFactor, "0.00")),
        ]
      : []),
    [],
    blockHeader("Recommendation"),
    rec
      ? kv("Recommended", verdict(`${rec.count} × ${rec.ratingKva} kVA${rec.layout.length > 1 ? ` (sets of ${rec.layout.join(" + ")})` : ""}`, true))
      : kv("Recommended", verdict(r.blockers.length > 0 ? "None: a load is too big for any transformer section" : "No configuration fits the ratings considered", false)),
    ...r.blockers.map((m): Row => [cell("Cannot be placed"), cell(m)]),
    ...(rec ? [kv("Incomers", cell(rec.count)), kv("Bus couplers", cell(rec.count - rec.layout.length))] : []),
    [],
    [label("Numbers of transformers tried")],
    heads("Transformers", "Sets", "Rating (kVA)", "Fits", "Max load (kVA)", "Average normal loading (%)", "One out loading (%)"),
    ...r.combos.map((c): Row => [
      cell(c.count), cell(c.layout.join(" + "), "center"), cell(c.chosen.ratingKva), verdict(c.fits ? "Yes" : "No", c.fits),
      n(c.chosen.maxLoadKva), limited(c.chosen.normalPct, normalPct), limited(c.chosen.contingencyPct, contPct),
    ]),
  ];
  return { sheet: "Summary", data, columns: autoColumns(data) };
}

/** The uploaded load list, as read (values only), with the header row in light orange. */
function originalSheet(raw: NonNullable<ReportOptions["originalSheet"]>): ReportSheet {
  const clean = (v: unknown): string | number | boolean | null => {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
    return String(v);
  };
  const blank = (v: unknown) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  const headerIdx = raw.findIndex((r) => r.some((v) => !blank(v)));
  const width = headerIdx >= 0 ? Math.max(...raw.map((r) => r.length)) : 0;
  const data: Row[] = raw.map((r, ri) =>
    Array.from({ length: ri >= headerIdx && headerIdx >= 0 ? width : r.length }, (_, ci) => r[ci]).map((v): ReportCell => {
      const c = clean(v);
      if (c === null) {
        if (ri < headerIdx) return null; // blank rows above the header stay empty
        return ri === headerIdx ? head("") : { value: "", ...BORDER };
      }
      if (ri === headerIdx) return head(String(c));
      if (typeof c === "number") return { value: c, align: "center", ...BORDER }; // as uploaded: no number format applied
      const bg = typeof c === "string" ? kindColor(c) : null;
      const isStatus = typeof c === "string" && ["w", "s"].includes(c.trim().toLowerCase());
      return { value: c, align: isStatus ? "center" : "left", ...(bg ? { backgroundColor: bg } : {}), ...BORDER };
    }),
  );
  return {
    sheet: "Original load list",
    data,
    columns: autoColumns(data, 8, 50),
    stickyRowsCount: headerIdx >= 0 ? headerIdx + 1 : undefined,
  };
}

function loadsSheet(r: SizingResult, o: ReportOptions): ReportSheet {
  const power: Row = o.includePower ? [n(r.totals.kw), n(r.totals.kvar), n(r.totals.powerFactor, "0.00")] : [];
  const bold = (v: string): ReportCell => ({ value: v, fontWeight: "bold", ...BORDER });
  const rows: Row[] = [
    loadHeaders(o.includePower),
    ...r.loads.map((l) => loadCells(l, o.includePower)),
    [cell(""), bold("Total"), cell(""), cell(""), cell(""), cell(""), ...power, n(r.totalKva), cell("")],
  ];
  if (r.spares.length > 0) {
    rows.push(
      [],
      [label("Spare feeders: one of each type and rating (not included in the calculation)")],
      loadHeaders(o.includePower),
      ...r.spares.map((sp) => spareCells(sp, o.includePower, [])),
    );
  }
  return { sheet: "Loads", data: rows, columns: autoColumns(rows), stickyRowsCount: 1 };
}

/** Re-colour every cell of a row (keeps each cell's own style). */
function tint(row: Row, color: string): Row {
  return row.map((c): ReportCell => {
    if (c !== null && typeof c === "object") return { ...c, backgroundColor: color };
    return { value: c ?? "", backgroundColor: color, ...BORDER };
  });
}

function layoutHeaders(includePower: boolean): Row {
  return heads(
    "Element", "Name", "Row", "Description", "W/S", "Duty/standby group", "Treated as", "Rated kW (motor)",
    ...(includePower ? ["kW (input)", "kvar", "Power factor"] : []),
    "kVA", "Basis",
  );
}

/**
 * The board of one set in order along the bus, as rows: incomers (cyan) and bus couplers (light yellow)
 * with the outgoing loads between them. Each section's loads sit next to its incomer and a coupler:
 *   set of 2:  Incomer T1 | loads T1 | Bus coupler BC1 | loads T2 | Incomer T2
 *   set of 3:  Incomer T1 | loads T1 | BC1 | Incomer T2 | loads T2 | BC2 | loads T3 | Incomer T3
 */
export function busLayoutRows(set: SetResult, c: ComboResult, o: ReportOptions, spares: SpareFeeder[] = [], placed: Set<string> = new Set()): Row[] {
  const multi = c.sets.length > 1;
  const R = c.chosen.ratingKva;
  const k = set.sections.length;
  const pf = (kw: number, kvar: number) => (kw === 0 && kvar === 0 ? 0 : kw / Math.hypot(kw, kvar));
  const empty = (): ReportCell => cell("");
  const incomer = (i: number): Row => {
    const s = set.sections[i];
    return tint([
      cell("Incomer"), cell(s.name), empty(), cell(`Incomer from ${s.name}: ${R} kVA transformer, ${o.transformerLabel}`), empty(), empty(), empty(), empty(),
      ...(o.includePower ? [n(s.kw), n(s.kvar), n(pf(s.kw, s.kvar), "0.00")] : []),
      n(s.kva), cell(`Section load: ${fmtPct(s.loadingPct)} % of ${fmtInt(R)} kVA (normal limit ${fmtPct(o.options.normalLimit * 100)} %)`),
    ], REPORT_COLORS.incomer);
  };
  const coupler = (i: number): Row => {
    const a = set.sections[i].name;
    const b = set.sections[i + 1].name;
    return tint([
      cell("Bus coupler"), cell(multi ? `S${set.index}-BC${i + 1}` : `BC${i + 1}`), empty(), cell(`Bus coupler between ${a} and ${b}`), empty(), empty(), empty(), empty(),
      ...(o.includePower ? [empty(), empty(), empty()] : []),
      empty(), cell(`Ties bus sections ${a} and ${b}`),
    ], REPORT_COLORS.coupler);
  };
  // Each spare feeder sits right after the first load of the same type and rating (one spare in total, not one per section).
  const loads = (i: number): Row[] =>
    set.sections[i].loads.flatMap((l): Row[] => {
      const row: Row = [cell("Outgoing load"), ...loadCells(l, o.includePower, set.sections[i].name)];
      const key = spareKey(l);
      const spare = key !== null && !placed.has(key) ? spares.find((x) => x.key === key) : undefined;
      if (!spare) return [row];
      placed.add(spare.key);
      return [row, spareCells(spare, o.includePower, [cell("Spare feeder"), cell(set.sections[i].name)])];
    });

  const rows: Row[] = [];
  for (let i = 0; i < k; i++) {
    if (i < k - 1) rows.push(incomer(i), ...loads(i), coupler(i));
    else rows.push(...loads(i), incomer(i)); // last section mirrored so its loads sit between a coupler and its incomer
  }
  return rows;
}

function comboSheet(c: ComboResult, o: ReportOptions, spares: SpareFeeder[]): ReportSheet {
  const placed = new Set<string>();
  const op = o.options;
  const normalPct = op.normalLimit * 100;
  const contPct = op.contingencyLimit * 100;
  const multi = c.sets.length > 1;
  const chain = (names: string[]) => names.join(" – coupler – ");
  const layoutLine = multi
    ? `${c.incomers} incomers, ${c.couplers} bus couplers, in ${c.sets.length} independent sets (${c.layout.join(" + ")}); one coupler between neighbouring incomers in a set, none between sets. ${c.sets.map((s) => `Set ${s.index}: ${chain(s.sections.map((x) => x.name))}`).join("; ")}`
    : `${c.incomers} incomers, ${c.couplers} bus coupler${c.couplers === 1 ? "" : "s"}; bus sections in a line: ${chain(c.sections.map((s) => s.name))}`;
  const data: Row[] = [
    [title(`${c.count} × ${c.chosen.ratingKva} kVA transformers, ${o.transformerLabel}`)],
    banner(c.fits ? "Fits within both limits" : "DOES NOT FIT within both limits", c.fits),
    ...(c.added ? [noteBanner("Added because the loads could not be balanced within the limits with fewer transformers.")] : []),
    [layoutLine],
    [],
    [label("Ratings considered")],
    heads("Rating (kVA)", "Installed kVA", "Average normal loading (%)", "One out loading (%)", "Result"),
    ...c.candidates.map((x): Row => [
      cell(x.ratingKva), cell(x.installedKva), limited(x.normalPct, normalPct), limited(x.contingencyPct, contPct),
      verdict(x.fits ? "Fits" : `Does not fit: ${failureReason(x)}`, x.fits),
    ]),
    [],
    ...(multi
      ? ([
          [label("Load per set")],
          heads("Set", "Transformers", "Load (kVA)", "Limit (kVA)", "Result"),
          ...c.sets.map((s): Row => [
            cell(s.index), cell(s.size), n(s.kva), n(s.capacityKva), verdict(s.kva <= s.capacityKva + 1e-9 ? "Within limit" : "Over limit", s.kva <= s.capacityKva + 1e-9),
          ]),
          [],
        ] as Row[])
      : []),
    ...c.sets.flatMap((set): Row[] => [
      [label(multi ? `Set ${set.index}: ${set.size} transformers – load distribution (normal running)` : "Load distribution (normal running)")],
      heads("Transformer", "Loads", ...(o.includePower ? ["kW (input)", "kvar"] : []), "kVA", "Loading (%)", `Within ${fmtPct(normalPct)} %`),
      ...set.sections.map((s): Row => [
        cell(s.name), cell(s.loads.length), ...(o.includePower ? [n(s.kw), n(s.kvar)] : []), n(s.kva),
        limited(s.loadingPct, normalPct), verdict(s.withinNormalLimit ? "Yes" : "No", s.withinNormalLimit),
      ]),
      [],
      [label(`One transformer out${multi ? ` in set ${set.index}` : ""} (each remaining unit limited to ${fmtPct(contPct)} %)`)],
      heads("Transformer out", "Load moved (kVA)", "Each remaining (kVA)", "Loading (%)", "Result"),
      ...set.oneOut.map((s): Row => [
        cell(s.outName), n(s.movedKva), n(s.eachKva), limited(s.loadingPct, contPct), verdict(s.ok ? "OK" : "Over limit", s.ok),
      ]),
      [],
    ]),
    ...(c.notes.length ? ([[label("Notes")], ...c.notes.map((t): Row => [t]), []] as Row[]) : []),
    [label(multi ? "Bus layout per set: incomers, outgoing loads and bus couplers" : "Bus layout: incomers, outgoing loads and bus couplers")],
    [
      { value: "Incomer", fontWeight: "bold", align: "center", backgroundColor: REPORT_COLORS.incomer, ...BORDER },
      { value: "Bus coupler", fontWeight: "bold", align: "center", backgroundColor: REPORT_COLORS.coupler, ...BORDER },
      { value: "Outgoing load", align: "center", ...BORDER },
      statusCell("W"), statusCell("S"),
      { value: "DOL", align: "center", backgroundColor: REPORT_COLORS.dol, ...BORDER },
      { value: "VFD", align: "center", backgroundColor: REPORT_COLORS.vfd, ...BORDER },
      { value: "Feeder", align: "center", backgroundColor: REPORT_COLORS.feeder, ...BORDER },
      { value: "Spare feeder", fontWeight: "bold", align: "center", backgroundColor: REPORT_COLORS.spare, ...BORDER },
    ],
    [],
    ...c.sets.flatMap((set): Row[] => [
      ...(multi ? ([[label(`Set ${set.index}`)]] as Row[]) : []),
      layoutHeaders(o.includePower),
      ...busLayoutRows(set, c, o, spares, placed),
      [],
    ]),
  ];
  return { sheet: `${c.count} Transformers`, data, columns: autoColumns(data) };
}

function issuesSheet(r: SizingResult): ReportSheet {
  const kind = { error: "Error (not included)", excluded: "Excluded", warning: "Warning" } as const;
  const data: Row[] = [
    heads("Row", "Description", "Type", "Message"),
    ...r.issues.map((i): Row => [cell(i.rowNumber), cell(i.description), cell(kind[i.severity]), cell(i.message)]),
  ];
  return { sheet: "Issues", data, columns: autoColumns(data), stickyRowsCount: 1 };
}

export function buildReportSheets(result: SizingResult, o: ReportOptions): ReportSheet[] {
  const sheets: ReportSheet[] = [
    summarySheet(result, o),
    ...(o.originalSheet ? [originalSheet(o.originalSheet)] : []),
    loadsSheet(result, o),
    ...displayCombos(result).map((c) => comboSheet(c, o, result.spares)),
    ...(result.issues.length > 0 ? [issuesSheet(result)] : []),
  ];
  // Wide tables print better in landscape.
  return sheets.map((sh) => ({ ...sh, orientation: "landscape" as const }));
}

export async function buildReportBlob(result: SizingResult, o: ReportOptions): Promise<Blob> {
  return writeExcelFile(buildReportSheets(result, o)).toBlob();
}
