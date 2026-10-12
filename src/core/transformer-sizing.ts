/**
 * Transformer count / rating selection for a load list (see docs/transformer-sizing.md).
 *
 * Units: volts (V), amps (A), kW, kVA. All loads are counted at 100 % (standby included).
 */

export type LoadType = "DOL" | "VFD" | "FEEDER";

export interface SizingOptions {
  /** Motor efficiency used when converting motor kW to kVA. */
  efficiency: number;
  /** Power factors by load type (motors: from kW; amps-only loads: used to report kW). */
  powerFactor: Record<LoadType, number>;
  /** Max loading of each transformer in normal running (fraction of rated kVA). */
  normalLimit: number;
  /** Max loading of each remaining transformer with one transformer out (fraction). */
  contingencyLimit: number;
  /** Standard transformer ratings in kVA, ascending. */
  ratingsKva: number[];
  /** Candidate counts to consider. 2 = two incomers + 1 coupler; 3 = three incomers + 2 couplers. */
  counts: number[];
  /**
   * How to order candidates before picking the first that fits:
   *  - "min-installed": ascending total installed kVA (2x2000, 2x2500, 3x2000, 3x2500, 2x4000, 3x4000)
   *  - "count-first":   add a transformer before increasing the rating (2x2000, 3x2000, 2x2500, 3x2500, ...)
   */
  stepUp: "min-installed" | "count-first";
  /**
   * How the total load in kVA is formed:
   *  - "arithmetic": sum of each load's kVA (conservative; ignores that loads have different power factors)
   *  - "vector":     sqrt((sum kW)^2 + (sum kvar)^2)
   */
  totalBasis: "arithmetic" | "vector";
}

/** Standard transformer ratings offered in the UI (kVA). */
export const STANDARD_RATINGS_KVA = [2000, 2500, 4000];

export const DEFAULT_OPTIONS: SizingOptions = {
  efficiency: 0.95,
  powerFactor: { DOL: 0.7, VFD: 0.9, FEEDER: 0.7 },
  normalLimit: 0.8,
  contingencyLimit: 0.9,
  ratingsKva: STANDARD_RATINGS_KVA,
  counts: [2, 3],
  stepUp: "min-installed",
  totalBasis: "arithmetic",
};

/** Transformer secondary voltage options offered in the UI (primary is 6600 V). */
export const TRANSFORMER_OPTIONS = [
  { id: "6600/433", primaryV: 6600, secondaryV: 433, label: "6600 / 433 V" },
  { id: "6600/690", primaryV: 6600, secondaryV: 690, label: "6600 / 690 V" },
] as const;

export interface RawLoadRow {
  /** 1-based row number in the source sheet, for traceability. */
  rowNumber: number;
  description: string;
  /** Free text from the Type column. */
  type: string;
  amps?: number;
  kw?: number;
  /** Volts. */
  voltageV?: number;
  /** Working (W) or standby (S), from the optional Working/Standby column. */
  status?: "W" | "S";
  /** Raw text of a Working/Standby cell that could not be read as W or S. */
  statusText?: string;
  /** Optional duty/standby group: loads with the same group name are put on different transformer sections. */
  group?: string;
  /** Set when a numeric cell could not be read (e.g. text in the Amps column). */
  badFields?: string[];
}

export interface SizedLoad {
  rowNumber: number;
  description: string;
  kind: "motor" | "feeder";
  type: LoadType;
  voltageV: number;
  status?: "W" | "S";
  /** Motor rating as given in the load list (shaft kW). Undefined for amps-only loads. */
  ratedKw?: number;
  /** Electrical input kW (motors: rated kW ÷ efficiency), so that kVA² = kW² + kvar². */
  kw: number;
  kvar: number;
  kva: number;
  powerFactor: number;
  basis: string;
  group?: string;
  warnings: string[];
}

export interface RowIssue {
  rowNumber: number;
  description: string;
  severity: "error" | "excluded" | "warning";
  message: string;
}

const SQRT3 = Math.sqrt(3);
const kvarOf = (kw: number, kva: number) => Math.sqrt(Math.max(kva * kva - kw * kw, 0));
const VOLTAGE_MATCH_TOLERANCE = 0.01;

function normaliseType(raw: string): LoadType | "UNKNOWN" {
  const t = raw.toUpperCase().replace(/[^A-Z]/g, "");
  if (t.includes("VFD")) return "VFD";
  if (t.includes("DOL")) return "DOL";
  if (t.includes("FEEDER")) return "FEEDER";
  return "UNKNOWN";
}

/**
 * Classify one load row and convert it to kVA.
 *  - kW given            -> motor; kVA = kW / (pf * efficiency); pf by type (VFD 0.9, otherwise DOL 0.7)
 *  - only amps, DOL/VFD  -> motor; kVA = sqrt(3) * V * I
 *  - only amps, otherwise-> outgoing feeder; kVA = sqrt(3) * V * I
 */
export function sizeLoad(
  row: RawLoadRow,
  selectedVoltageV: number,
  opts: SizingOptions = DEFAULT_OPTIONS,
): { load: SizedLoad } | { issue: RowIssue } {
  const fail = (severity: RowIssue["severity"], message: string) => ({
    issue: { rowNumber: row.rowNumber, description: row.description, severity, message },
  });
  const warnings: string[] = [];
  if (row.statusText) warnings.push(`Working/standby value "${row.statusText}" not recognised (use W or S).`);

  if (row.badFields?.length) return fail("error", `Could not read a number in: ${row.badFields.join(", ")}.`);
  if (row.kw !== undefined && !(row.kw > 0)) return fail("error", "kW must be greater than 0.");
  if (row.amps !== undefined && !(row.amps > 0)) return fail("error", "Amps must be greater than 0.");
  if (row.kw === undefined && row.amps === undefined) return fail("error", "Neither kW nor amps is given.");

  let voltageV = row.voltageV;
  if (voltageV === undefined) {
    voltageV = selectedVoltageV;
    warnings.push(`No voltage given; assumed ${selectedVoltageV} V.`);
  } else if (Math.abs(voltageV - selectedVoltageV) / selectedVoltageV > VOLTAGE_MATCH_TOLERANCE) {
    return fail("excluded", `Voltage ${voltageV} V does not match the selected ${selectedVoltageV} V transformer.`);
  }

  const type = normaliseType(row.type);

  if (row.kw !== undefined) {
    let pfType: LoadType = "DOL";
    if (type === "VFD") pfType = "VFD";
    else if (type !== "DOL") {
      warnings.push(
        type === "FEEDER"
          ? "Type is Feeder but kW is given; treated as a DOL motor."
          : "Type not recognised; treated as a DOL motor.",
      );
    }
    const pf = opts.powerFactor[pfType];
    const kwIn = row.kw / opts.efficiency;
    const kva = kwIn / pf;
    return {
      load: {
        rowNumber: row.rowNumber,
        description: row.description,
        kind: "motor",
        type: pfType,
        voltageV,
        ratedKw: row.kw,
        kw: kwIn,
        kvar: kvarOf(kwIn, kva),
        kva,
        powerFactor: pf,
        basis: `kW ÷ (pf ${pf} × η ${opts.efficiency})`,
        group: row.group,
        status: row.status,
        warnings,
      },
    };
  }

  const amps = row.amps as number;
  const kva = (SQRT3 * voltageV * amps) / 1000;
  const isMotor = type === "DOL" || type === "VFD";
  const loadType: LoadType = isMotor ? type : "FEEDER";
  const pf = opts.powerFactor[loadType];
  if (!isMotor && type === "UNKNOWN" && row.type.trim() !== "") {
    warnings.push(`Type "${row.type}" not recognised; treated as an outgoing feeder.`);
  }
  return {
    load: {
      rowNumber: row.rowNumber,
      description: row.description,
      kind: isMotor ? "motor" : "feeder",
      type: loadType,
      voltageV,
      kw: kva * pf,
      kvar: kvarOf(kva * pf, kva),
      kva,
      powerFactor: pf,
      basis: `√3 × ${voltageV} V × ${amps} A (kW shown at pf ${pf})`,
      group: row.group,
      status: row.status,
      warnings,
    },
  };
}

export interface Candidate {
  count: number;
  ratingKva: number;
  installedKva: number;
  /** Loading of each transformer in normal running, % of rating (load shared equally). */
  normalPct: number;
  /** Loading of each remaining transformer with one out, % of rating (load shared equally). */
  contingencyPct: number;
  normalOk: boolean;
  contingencyOk: boolean;
  fits: boolean;
  /** Largest total load (kVA) this configuration can carry within both limits. */
  maxLoadKva: number;
}

export function evaluateCandidate(
  totalKva: number,
  count: number,
  ratingKva: number,
  opts: SizingOptions = DEFAULT_OPTIONS,
): Candidate {
  const normalPct = (totalKva / (count * ratingKva)) * 100;
  const contingencyPct = (totalKva / ((count - 1) * ratingKva)) * 100;
  const normalOk = normalPct <= opts.normalLimit * 100 + 1e-9;
  const contingencyOk = contingencyPct <= opts.contingencyLimit * 100 + 1e-9;
  return {
    count,
    ratingKva,
    installedKva: count * ratingKva,
    normalPct,
    contingencyPct,
    normalOk,
    contingencyOk,
    fits: normalOk && contingencyOk,
    maxLoadKva: Math.min(count * ratingKva * opts.normalLimit, (count - 1) * ratingKva * opts.contingencyLimit),
  };
}

export function candidatesInOrder(totalKva: number, opts: SizingOptions = DEFAULT_OPTIONS): Candidate[] {
  const all: Candidate[] = [];
  for (const rating of opts.ratingsKva) {
    for (const count of opts.counts) all.push(evaluateCandidate(totalKva, count, rating, opts));
  }
  if (opts.stepUp === "min-installed") {
    // Stable sort keeps the rating/count order for equal installed kVA.
    return all.sort((a, b) => a.installedKva - b.installedKva);
  }
  return all; // already rating-major, count-minor = "count-first"
}

export interface Power { kw: number; kvar: number; kva: number }

/** Combine loads into one kW / kvar / kVA figure using the chosen total basis. */
export function combine(loads: Pick<SizedLoad, "kw" | "kvar" | "kva">[], basis: SizingOptions["totalBasis"]): Power {
  const kw = loads.reduce((a, l) => a + l.kw, 0);
  const kvar = loads.reduce((a, l) => a + l.kvar, 0);
  const kva = basis === "vector" ? Math.hypot(kw, kvar) : loads.reduce((a, l) => a + l.kva, 0);
  return { kw, kvar, kva };
}

export interface Section extends Power {
  /** "T1", "T2", ... — also names the bus section fed by that transformer. */
  name: string;
  loads: SizedLoad[];
  /** % of rated kVA in normal running. */
  loadingPct: number;
  withinNormalLimit: boolean;
}

const groupKey = (g?: string) => (g ?? "").trim().toLowerCase();

/**
 * Spread whole loads over the bus sections to balance kVA (largest load first, onto the
 * least-loaded section). Loads with the same duty/standby group go on different sections
 * whenever the group is no larger than the number of sections. Loads are never split, so
 * sections will not be exactly equal.
 */
export function distributeLoads(
  loads: SizedLoad[],
  count: number,
  ratingKva: number,
  opts: SizingOptions = DEFAULT_OPTIONS,
): { sections: Section[]; notes: string[] } {
  const buckets: SizedLoad[][] = Array.from({ length: count }, () => []);
  const kvaOf = (i: number) => buckets[i].reduce((a, l) => a + l.kva, 0);
  const usedByGroup = new Map<string, Set<number>>();
  const notes: string[] = [];
  const sorted = [...loads].sort((a, b) => b.kva - a.kva || a.rowNumber - b.rowNumber);

  for (const load of sorted) {
    const key = groupKey(load.group);
    const used = key ? (usedByGroup.get(key) ?? new Set<number>()) : new Set<number>();
    let candidates = buckets.map((_, i) => i).filter((i) => !used.has(i));
    if (candidates.length === 0) {
      candidates = buckets.map((_, i) => i);
      notes.push(
        `Group "${load.group}" has more loads than the ${count} transformers, so "${load.description || `row ${load.rowNumber}`}" shares a section with another member.`,
      );
    }
    const target = candidates.reduce((best, i) => (kvaOf(i) < kvaOf(best) ? i : best), candidates[0]);
    buckets[target].push(load);
    if (key) usedByGroup.set(key, used.add(target));
  }

  const sections: Section[] = buckets.map((l, i) => {
    const p = combine(l, opts.totalBasis);
    const loadingPct = (p.kva / ratingKva) * 100;
    return {
      name: `T${i + 1}`, loads: l, ...p, loadingPct,
      withinNormalLimit: loadingPct <= opts.normalLimit * 100 + 1e-9,
    };
  });
  return { sections, notes };
}

export interface OneOutScenario {
  /** Transformer that is out, e.g. "T2". */
  outName: string;
  /** kVA of that transformer's own bus section that has to be picked up by the others. */
  movedKva: number;
  /** kVA each remaining transformer carries (total shared equally through the closed bus couplers). */
  eachKva: number;
  loadingPct: number;
  ok: boolean;
}

export interface ComboResult {
  count: number;
  incomers: number;
  couplers: number;
  /** One entry per rating considered, ascending. */
  candidates: Candidate[];
  /** Smallest rating that fits; if none fits, the largest rating considered (see `fits`). */
  chosen: Candidate;
  fits: boolean;
  sections: Section[];
  oneOut: OneOutScenario[];
  notes: string[];
}

export interface SizingResult {
  selectedVoltageV: number;
  loads: SizedLoad[];
  issues: RowIssue[];
  /** Totals by every basis, for display. */
  totals: Power & { kvaArithmetic: number; kvaVector: number; powerFactor: number };
  /** Total load used for sizing (per `totalBasis`). */
  totalKva: number;
  totalKw: number;
  /** All counts × ratings, in order of preference. */
  candidates: Candidate[];
  recommended?: Candidate;
  /** One result per transformer count considered (e.g. 2 and 3). */
  combos: ComboResult[];
  // Convenience for the recommended combination:
  incomers?: number;
  couplers?: number;
  sections?: Section[];
}

function buildCombo(loads: SizedLoad[], totalKva: number, count: number, opts: SizingOptions): ComboResult {
  const candidates = [...opts.ratingsKva]
    .sort((a, b) => a - b)
    .map((r) => evaluateCandidate(totalKva, count, r, opts));
  const fitting = candidates.find((c) => c.fits);
  const chosen = fitting ?? candidates[candidates.length - 1];
  const { sections, notes } = distributeLoads(loads, count, chosen.ratingKva, opts);
  const eachKva = totalKva / (count - 1);
  const oneOut = sections.map((s) => {
    const loadingPct = (eachKva / chosen.ratingKva) * 100;
    return {
      outName: s.name,
      movedKva: s.kva,
      eachKva,
      loadingPct,
      ok: loadingPct <= opts.contingencyLimit * 100 + 1e-9,
    };
  });
  const overloaded = sections.filter((s) => !s.withinNormalLimit).map((s) => s.name);
  if (fitting && overloaded.length > 0) {
    notes.push(
      `The total fits, but the balanced distribution puts ${overloaded.join(", ")} above the ${opts.normalLimit * 100} % normal limit (loads cannot be split).`,
    );
  }
  return { count, incomers: count, couplers: count - 1, candidates, chosen, fits: !!fitting, sections, oneOut, notes };
}

export function sizeTransformers(
  rows: RawLoadRow[],
  selectedVoltageV: number,
  opts: SizingOptions = DEFAULT_OPTIONS,
): SizingResult {
  const loads: SizedLoad[] = [];
  const issues: RowIssue[] = [];
  for (const row of rows) {
    const r = sizeLoad(row, selectedVoltageV, opts);
    if ("issue" in r) issues.push(r.issue);
    else {
      loads.push(r.load);
      for (const w of r.load.warnings) {
        issues.push({ rowNumber: row.rowNumber, description: row.description, severity: "warning", message: w });
      }
    }
  }
  const arithmetic = combine(loads, "arithmetic");
  const vector = combine(loads, "vector");
  const used = opts.totalBasis === "vector" ? vector : arithmetic;
  const totals = {
    kw: used.kw,
    kvar: used.kvar,
    kva: used.kva,
    kvaArithmetic: arithmetic.kva,
    kvaVector: vector.kva,
    powerFactor: vector.kva > 0 ? vector.kw / vector.kva : 0,
  };
  const totalKva = used.kva;
  const candidates = candidatesInOrder(totalKva, opts);
  const recommended = loads.length > 0 ? candidates.find((c) => c.fits) : undefined;
  const combos = loads.length > 0 ? opts.counts.map((n) => buildCombo(loads, totalKva, n, opts)) : [];
  const result: SizingResult = {
    selectedVoltageV, loads, issues, totals, totalKva, totalKw: used.kw, candidates, recommended, combos,
  };
  if (recommended) {
    const combo = combos.find((c) => c.count === recommended.count)!;
    result.incomers = combo.incomers;
    result.couplers = combo.couplers;
    result.sections = combo.sections;
  }
  return result;
}

/* ---------------------------- spreadsheet reading ---------------------------- */

/** Cell values as read from a sheet (string, number, boolean, Date, null); typed loosely on purpose. */
type Cell = unknown;

const blank = (c: Cell) => c === null || c === undefined || (typeof c === "string" && c.trim() === "");

function parseNumber(c: Cell): { value?: number; bad: boolean } {
  if (blank(c)) return { bad: false };
  if (typeof c === "number") return Number.isFinite(c) ? { value: c, bad: false } : { bad: true };
  if (typeof c === "string") {
    const m = c.replace(/,/g, "").match(/^\s*(-?\d+(?:\.\d+)?)\s*[a-zA-Z]*\s*$/);
    if (m) return { value: Number(m[1]), bad: false };
  }
  return { bad: true };
}

/** "433", 433, "433 V", 0.433 or "0.433 kV" -> 433 (volts). */
export function parseVoltage(c: Cell): { value?: number; bad: boolean } {
  if (blank(c)) return { bad: false };
  const isKv = typeof c === "string" && /kv/i.test(c);
  const n = parseNumber(c);
  if (n.bad || n.value === undefined) return { bad: true };
  return { value: isKv || n.value < 10 ? n.value * 1000 : n.value, bad: false };
}

/** "W", "Working", "Duty" -> W; "S", "Standby", "Stand-by" -> S. */
export function parseStatus(c: Cell): { status?: "W" | "S"; unrecognised?: string } {
  if (blank(c)) return {};
  const t = String(c).toLowerCase().replace(/[^a-z]/g, "");
  if (["w", "working", "work", "duty", "run", "running"].includes(t)) return { status: "W" };
  if (["s", "standby", "stby", "sb"].includes(t)) return { status: "S" };
  return { unrecognised: String(c).trim() };
}

const norm = (h: Cell) => String(h ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");

export interface ParsedSheet {
  rows: RawLoadRow[];
  /** Fatal problems (e.g. missing columns). When non-empty, `rows` is empty. */
  errors: string[];
}

/** Turn sheet data (array of rows, first non-empty row = headers) into load rows. */
export function parseLoadSheet(data: Cell[][]): ParsedSheet {
  const headerIdx = data.findIndex((r) => r.some((c) => !blank(c)));
  if (headerIdx < 0) return { rows: [], errors: ["The sheet is empty."] };

  const cols: Partial<Record<"description" | "type" | "voltage" | "amps" | "kw" | "group" | "status", number>> = {};
  data[headerIdx].forEach((h, i) => {
    const n = norm(h);
    if (!n) return;
    if (n.startsWith("desc")) cols.description ??= i;
    else if (n.startsWith("type")) cols.type ??= i;
    else if (n.includes("volt")) cols.voltage ??= i;
    else if (n.includes("amp")) cols.amps ??= i;
    else if (n.includes("kw")) cols.kw ??= i;
    else if (n.includes("group") || n.includes("pair")) cols.group ??= i;
    else if (n.includes("working") || n.includes("standby") || n.includes("duty") || n === "ws") cols.status ??= i;
  });

  const missing = (["description", "type", "amps", "kw", "voltage"] as const).filter((k) => cols[k] === undefined);
  if (missing.length > 0) {
    return {
      rows: [],
      errors: [
        `Missing column(s): ${missing.join(", ")}. Expected headers: Description, Type, Amps rating, kW rating, Voltage level.`,
      ],
    };
  }

  const rows: RawLoadRow[] = [];
  for (let r = headerIdx + 1; r < data.length; r++) {
    const line = data[r];
    if (line.every(blank)) continue;
    const badFields: string[] = [];
    const amps = parseNumber(line[cols.amps!]);
    const kw = parseNumber(line[cols.kw!]);
    const volt = parseVoltage(line[cols.voltage!]);
    if (amps.bad) badFields.push("Amps");
    if (kw.bad) badFields.push("kW");
    if (volt.bad) badFields.push("Voltage");
    const status = cols.status !== undefined ? parseStatus(line[cols.status]) : {};
    rows.push({
      rowNumber: r + 1,
      description: String(line[cols.description!] ?? "").trim(),
      type: String(line[cols.type!] ?? "").trim(),
      amps: amps.value,
      kw: kw.value,
      voltageV: volt.value,
      status: status.status,
      statusText: status.unrecognised,
      group: cols.group !== undefined && !blank(line[cols.group]) ? String(line[cols.group]).trim() : undefined,
      badFields: badFields.length ? badFields : undefined,
    });
  }
  return { rows, errors: [] };
}
