/**
 * DesignDesign-project data model for PMCC boards and IMCCs, and the rules that keep them consistent (docs/imcc.md).
 *
 * One source of truth: `DesignProject.imccs` (the IMCC designs) and `DesignProject.pmccs` (each with its own load list).
 * The outgoing feeder an IMCC incomer occupies in a PMCC is DERIVED from the IMCC (`syncLinks`), keyed by a stable
 * id (`imcc:<imccId>:<incomer>`), so it can never be duplicated or drift from the IMCC. Manually entered PMCC
 * feeders (the rows of the PMCC load list) are never touched.
 *
 * All operations are pure: they return a new DesignProject or a list of errors, never a half-updated one.
 * Current and breaker sizing reuse `currentFor` / `selectBreaker` (mcc-breaker.ts); there is no second algorithm.
 */
import { DEFAULT_MCC_ASSUMPTIONS, breakerLabel, currentFor, selectBreaker, type BreakerSelection, type MccAssumptions, type MccLoadType } from "./mcc-breaker";
import { parseLoadSheet } from "./transformer-sizing";

/* --------------------------------- status --------------------------------- */

/** W = working, S = standby, SP = full spare feeder ("Spare" in the input). */
export type LoadStatus = "W" | "S" | "SP";

/** `W` and `w` are the same; `Spare` (or `SP`) is a spare feeder; anything else is invalid (undefined). */
export function normalizeStatus(v: unknown): LoadStatus | undefined {
  const t = String(v ?? "").trim().toLowerCase();
  if (t === "w" || t === "working") return "W";
  if (t === "s" || t === "standby") return "S";
  if (t === "spare" || t === "sp") return "SP";
  return undefined;
}

export type StartMethod = "DOL" | "RDOL" | "VFD" | "PWR FDR";

export function normalizeMethod(v: unknown): StartMethod | undefined {
  const t = String(v ?? "").trim().toUpperCase().replace(/\s+/g, " ");
  if (t === "DOL" || t === "RDOL" || t === "VFD") return t;
  if (t === "PWR FDR" || t === "POWER FEEDER" || t === "FEEDER" || t === "PWR FEEDER") return "PWR FDR";
  return undefined;
}

export const loadTypeOf = (m: StartMethod | undefined): MccLoadType => (m === "VFD" ? "VFD" : m === "DOL" || m === "RDOL" ? "DOL" : "FEEDER");

/* ---------------------------------- model --------------------------------- */

export type BusSection = 1 | 2;

export interface ImccLoad {
  id: string;
  /** Bus section (incomer 1 or 2). Missing = flagged, never guessed. */
  section?: BusSection;
  sNo?: string;
  description: string;
  kw?: number;
  ratedA?: number;
  /** Normalised status; undefined = missing or invalid (see `statusText`). */
  status?: LoadStatus;
  statusText?: string;
  method?: StartMethod;
  group?: string;
  /** Free text from the source (e.g. "Rating changed from 5.5 kW to 30 kW", feeder description). */
  note?: string;
}

/** Operating philosophy as a rule. The default is the project's: exactly one incomer ON, coupler OFF. */
export interface OperatingRule {
  exactlyOneIncomerOn: boolean;
  couplerMustBeOff: boolean;
  /** A rule that differs from the default is an engineering change and needs a named authoriser. */
  authorisedBy?: string;
}
export const DEFAULT_RULE: OperatingRule = { exactlyOneIncomerOn: true, couplerMustBeOff: true };

export interface IncomerSource {
  pmccId: string;
  /** Reference of the PMCC outgoing feeder supplying this incomer (e.g. "Feeder 12"). */
  feederRef: string;
}

export interface Imcc {
  id: string;
  tag: string;
  description: string;
  voltageV: number;
  incomer1: IncomerSource;
  incomer2: IncomerSource;
  /** Which incomer is ON in normal operation (the other and the coupler are OFF). */
  normallyOn: BusSection;
  rule: OperatingRule;
  loads: ImccLoad[];
}

export interface Pmcc {
  id: string;
  tag: string;
  /** The PMCC load list as uploaded (header row first): the manual outgoing feeders. */
  fileName?: string;
  sheet?: unknown[][];
}

/** The PMCC outgoing feeder generated for an IMCC incomer. Derived data: rebuilt by `syncLinks`. */
export interface GeneratedFeeder {
  id: string;
  pmccId: string;
  imccId: string;
  incomer: BusSection;
  description: string;
  destination: string;
  feederRef: string;
  /** W when this incomer is the one normally ON, S for the one normally OFF. */
  status: LoadStatus;
  /** Working current / kVA of the bus section this incomer feeds. */
  ratedA: number;
  kva: number;
  designA?: number;
  rating?: BreakerSelection;
}

export interface DesignProject {
  version: 1;
  name: string;
  pmccs: Pmcc[];
  imccs: Imcc[];
  generated: GeneratedFeeder[];
}

export const emptyDesignProject = (name = "DesignProject"): DesignProject => ({ version: 1, name, pmccs: [], imccs: [], generated: [] });

/** Accepts a project of this or an earlier shape (no IMCCs, no generated feeders) and fills in the defaults. */
export function migrateProject(raw: unknown): DesignProject {
  const r = (raw && typeof raw === "object" ? raw : {}) as Partial<DesignProject>;
  const p: DesignProject = {
    version: 1,
    name: typeof r.name === "string" ? r.name : "DesignProject",
    pmccs: Array.isArray(r.pmccs) ? r.pmccs : [],
    imccs: Array.isArray(r.imccs) ? r.imccs.map((i) => ({ ...i, rule: i.rule ?? { ...DEFAULT_RULE }, normallyOn: i.normallyOn ?? 1, loads: i.loads ?? [] })) : [],
    generated: [],
  };
  return syncLinks(p);
}

let counter = 0;
export const newId = (prefix: string): string =>
  `${prefix}_${Date.now().toString(36)}${(counter++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;

/* ------------------------------ operating state ---------------------------- */

export type SwitchState = "ON" | "OFF";
export interface OperatingState { incomer1: SwitchState; incomer2: SwitchState; coupler: SwitchState }

export interface Validation { ok: boolean; errors: string[]; warnings: string[] }

/**
 * Checks a state against the rule. With the default rule the only permitted normal state is one incomer ON,
 * the other incomer OFF and the bus coupler OFF.
 */
export function validateOperatingState(s: OperatingState, rule: OperatingRule = DEFAULT_RULE): Validation {
  const errors: string[] = [];
  const warnings: string[] = [];
  const on = (s.incomer1 === "ON" ? 1 : 0) + (s.incomer2 === "ON" ? 1 : 0);
  const changed = rule.exactlyOneIncomerOn !== DEFAULT_RULE.exactlyOneIncomerOn || rule.couplerMustBeOff !== DEFAULT_RULE.couplerMustBeOff;
  if (changed && !rule.authorisedBy?.trim()) {
    errors.push("The operating philosophy differs from the default (one incomer ON, the other incomer and the bus coupler OFF). That is an engineering change and needs a named authoriser.");
  }
  if (rule.exactlyOneIncomerOn) {
    if (on === 2) errors.push("Both incomers are ON: not permitted. Exactly one incomer may be ON at a time.");
    if (on === 0) errors.push("No incomer is ON: the IMCC is not supplied. Exactly one incomer must be ON.");
  }
  if (rule.couplerMustBeOff && s.coupler === "ON") {
    errors.push(
      on === 1
        ? "Bus coupler is ON with one incomer ON: not permitted. One incomer is not assumed to supply both bus sections through the coupler."
        : "Bus coupler is ON: it must remain OFF.",
    );
  }
  if (errors.length === 0 && on === 1 && s.coupler === "OFF") warnings.push("Permitted normal state: one incomer ON, the other incomer OFF, bus coupler OFF.");
  return { ok: errors.length === 0, errors, warnings };
}

export const normalState = (i: Pick<Imcc, "normallyOn">): OperatingState => ({
  incomer1: i.normallyOn === 1 ? "ON" : "OFF",
  incomer2: i.normallyOn === 2 ? "ON" : "OFF",
  coupler: "OFF",
});

/* ------------------------------ IMCC calculation --------------------------- */

export interface LoadResult {
  load: ImccLoad;
  usedA?: number;
  designA?: number;
  kva?: number;
  /** Counted in the normal working load: status W only. */
  working: boolean;
  workingKva: number;
  workingA: number;
  breaker?: BreakerSelection;
}

export interface SectionResult {
  section: BusSection;
  loads: LoadResult[];
  working: number;
  standby: number;
  spare: number;
  /** Working + standby (spare feeders carry no load). */
  connectedKva: number;
  workingKva: number;
  workingA: number;
  standbyKva: number;
  /** Reference workbook's "LCE KVA" = working kVA x 0.8. Meaning not confirmed (OPEN_ITEMS.md #8): reported, never used for sizing. */
  lceKva: number;
  designA: number;
  breaker: BreakerSelection;
}

export interface ImccResult {
  sections: [SectionResult, SectionResult];
  /** Calculated design rating of Incomer 1 (from its own section's working load). */
  incomer1: BreakerSelection;
  /** Incomer 2 and the bus coupler follow Incomer 1. */
  incomer2: BreakerSelection;
  coupler: BreakerSelection;
  issues: { severity: "error" | "warning"; message: string }[];
  operating: Validation;
}

export const LCE_FACTOR = 0.8;
const SQ3 = Math.sqrt(3);
const calcV = (v: number, a: MccAssumptions) => (Math.round(v) === 433 ? a.calcVoltage433 : v);

export function computeImcc(imcc: Imcc, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): ImccResult {
  const issues: ImccResult["issues"] = [];
  const mk = (section: BusSection): SectionResult => {
    const loads: LoadResult[] = imcc.loads.filter((l) => l.section === section).map((l) => {
      const cur = currentFor({ type: loadTypeOf(l.method), kw: l.kw, ratedA: l.ratedA, voltageV: imcc.voltageV }, a);
      const kva = cur.usedA === undefined ? undefined : (SQ3 * calcV(imcc.voltageV, a) * cur.usedA) / 1000;
      const working = l.status === "W";
      return {
        load: l, usedA: cur.usedA, designA: cur.designA, kva, working,
        workingKva: working && kva !== undefined ? kva : 0,
        workingA: working && cur.usedA !== undefined ? cur.usedA : 0,
        breaker: cur.designA === undefined ? undefined : selectBreaker(cur.designA, { mccb630: a.mccb630 }),
      };
    });
    const sum = (f: (x: LoadResult) => number) => loads.reduce((t, x) => t + f(x), 0);
    const workingA = sum((x) => x.workingA);
    const designA = workingA * a.designMultiplier;
    const workingKva = sum((x) => x.workingKva);
    const st = (s: LoadStatus) => loads.filter((x) => x.load.status === s);
    return {
      section, loads,
      working: st("W").length, standby: st("S").length, spare: st("SP").length,
      workingKva, workingA, designA,
      standbyKva: st("S").reduce((t, x) => t + (x.kva ?? 0), 0),
      connectedKva: [...st("W"), ...st("S")].reduce((t, x) => t + (x.kva ?? 0), 0),
      lceKva: workingKva * LCE_FACTOR,
      breaker: selectBreaker(workingA > 0 ? designA : undefined, { mccb630: a.mccb630 }),
    };
  };
  const s1 = mk(1);
  const s2 = mk(2);
  const incomer1 = s1.breaker;
  // Incomer 2 and the bus coupler are the same design rating as Incomer 1 (rule), always.
  const incomer2 = { ...incomer1 };
  const coupler = { ...incomer1 };

  const noSection = imcc.loads.filter((l) => l.section === undefined);
  if (noSection.length) issues.push({ severity: "error", message: `${noSection.length} load(s) have no bus section (incomer 1 or 2): not allocated, not counted. Allocation is never guessed.` });
  for (const l of imcc.loads) {
    if (l.status === undefined) issues.push({ severity: "error", message: `"${l.description}": status ${l.statusText ? `"${l.statusText}"` : "missing"}; use W, w, S or Spare. Not counted.` });
    else if (l.status !== "SP" && l.kw === undefined && l.ratedA === undefined) issues.push({ severity: "warning", message: `"${l.description}": no kW or current, so it carries no load in the calculation.` });
    if (l.status !== undefined && l.status !== "SP" && !l.method) issues.push({ severity: "warning", message: `"${l.description}": method of starting missing or not recognised (DOL, RDOL, VFD, PWR FDR); treated as a power feeder.` });
  }
  for (const sec of [s1, s2]) {
    const groups = new Map<string, { w: number; s: number }>();
    for (const x of sec.loads) {
      const g = x.load.group?.trim();
      if (!g || x.load.status === undefined || x.load.status === "SP") continue;
      const e = groups.get(g) ?? { w: 0, s: 0 };
      if (x.load.status === "W") e.w++; else e.s++;
      groups.set(g, e);
    }
    for (const [g, e] of groups) if (e.w === 0 && e.s > 0) issues.push({ severity: "warning", message: `Incomer ${sec.section} section: duty/standby group "${g}" has standby but no working load.` });
  }
  if (incomer1.kind === "REVIEW") issues.push({ severity: "warning", message: `Incomer 1: ${incomer1.message}` });
  if (s2.breaker.ratingA !== undefined && incomer1.ratingA !== undefined && s2.breaker.ratingA > incomer1.ratingA) {
    issues.push({ severity: "warning", message: `Incomer 2 section needs ${breakerLabel(s2.breaker)}, more than Incomer 1's ${breakerLabel(incomer1)} which Incomer 2 and the bus coupler follow: review.` });
  }
  if (s2.workingA > 0 && s1.workingA === 0) issues.push({ severity: "warning", message: "Incomer 1 section has no working load, so no Incomer 1 rating can be calculated." });
  return { sections: [s1, s2], incomer1, incomer2, coupler, issues, operating: validateOperatingState(normalState(imcc), imcc.rule) };
}

/* --------------------------------- links ---------------------------------- */

export const linkId = (imccId: string, incomer: BusSection) => `imcc:${imccId}:${incomer}`;

/** Rebuilds the generated PMCC feeders from the IMCCs. Idempotent: same input, same output, no duplicates. */
export function syncLinks(p: DesignProject, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): DesignProject {
  const generated: GeneratedFeeder[] = [];
  for (const i of p.imccs) {
    const r = computeImcc(i, a);
    for (const n of [1, 2] as const) {
      const src = n === 1 ? i.incomer1 : i.incomer2;
      const sec = r.sections[n - 1];
      const pmcc = p.pmccs.find((x) => x.id === src.pmccId);
      if (!pmcc) continue; // reported by validateImcc; nothing is generated for a missing PMCC
      const designA = sec.workingA > 0 ? sec.workingA * a.designMultiplier : undefined;
      generated.push({
        id: linkId(i.id, n), pmccId: pmcc.id, imccId: i.id, incomer: n,
        description: `${i.tag} Incomer ${n}`, destination: i.tag, feederRef: src.feederRef,
        status: n === i.normallyOn ? "W" : "S",
        ratedA: sec.workingA, kva: sec.workingKva, designA,
        // PMCC feeder sized with the existing rules on its own current: NOT inherited from the IMCC incomer rating.
        rating: selectBreaker(designA, { mccb630: a.mccb630 }),
      });
    }
  }
  return { ...p, generated };
}

export type Result<T> = { ok: true; value: T } | { ok: false; errors: string[] };

export function validateImcc(p: DesignProject, i: Imcc): string[] {
  const errors: string[] = [];
  if (!i.tag.trim()) errors.push("The IMCC needs a tag / name.");
  if (p.imccs.some((x) => x.id !== i.id && x.tag.trim().toLowerCase() === i.tag.trim().toLowerCase())) errors.push(`IMCC tag "${i.tag}" is already used.`);
  for (const n of [1, 2] as const) {
    const s = n === 1 ? i.incomer1 : i.incomer2;
    if (!s.pmccId) errors.push(`Incomer ${n}: select the upstream PMCC.`);
    else if (!p.pmccs.some((x) => x.id === s.pmccId)) errors.push(`Incomer ${n}: the selected PMCC does not exist.`);
    if (!s.feederRef.trim()) errors.push(`Incomer ${n}: give the reference of the PMCC outgoing feeder that supplies it.`);
  }
  if (!(i.voltageV > 0)) errors.push("Give the IMCC voltage.");
  const op = validateOperatingState(normalState(i), i.rule);
  errors.push(...op.errors);
  return errors;
}

/** Creates or updates an IMCC together with its PMCC feeder links, or changes nothing and says why. */
export function saveImcc(p: DesignProject, i: Imcc, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): Result<DesignProject> {
  const errors = validateImcc(p, i);
  if (errors.length) return { ok: false, errors };
  const exists = p.imccs.some((x) => x.id === i.id);
  const imccs = exists ? p.imccs.map((x) => (x.id === i.id ? i : x)) : [...p.imccs, i];
  return { ok: true, value: syncLinks({ ...p, imccs }, a) };
}

export function duplicateImcc(p: DesignProject, id: string, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): Result<DesignProject> {
  const src = p.imccs.find((x) => x.id === id);
  if (!src) return { ok: false, errors: ["The IMCC to duplicate does not exist."] };
  let tag = `${src.tag}-copy`;
  for (let n = 2; p.imccs.some((x) => x.tag.toLowerCase() === tag.toLowerCase()); n++) tag = `${src.tag}-copy${n}`;
  const copy: Imcc = {
    ...src, id: newId("imcc"), tag,
    incomer1: { ...src.incomer1 }, incomer2: { ...src.incomer2 }, rule: { ...src.rule },
    loads: src.loads.map((l) => ({ ...l, id: newId("load") })),
  };
  return saveImcc(p, copy, a);
}

export interface DeleteOutcome { project?: DesignProject; needsConfirmation?: boolean; dependencies: GeneratedFeeder[]; errors: string[] }

/** Deleting removes the IMCC and ONLY the feeders it generated; it asks for confirmation first when there are any. */
export function deleteImcc(p: DesignProject, id: string, confirmed = false, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): DeleteOutcome {
  if (!p.imccs.some((x) => x.id === id)) return { dependencies: [], errors: ["The IMCC does not exist."] };
  const dependencies = p.generated.filter((g) => g.imccId === id);
  if (dependencies.length && !confirmed) return { needsConfirmation: true, dependencies, errors: [] };
  return { project: syncLinks({ ...p, imccs: p.imccs.filter((x) => x.id !== id) }, a), dependencies, errors: [] };
}

export function addPmcc(p: DesignProject, tag: string, sheet?: { fileName: string; sheet: unknown[][] }): Result<DesignProject> {
  if (!tag.trim()) return { ok: false, errors: ["The PMCC needs a tag."] };
  if (p.pmccs.some((x) => x.tag.trim().toLowerCase() === tag.trim().toLowerCase())) return { ok: false, errors: [`PMCC tag "${tag}" is already used.`] };
  return { ok: true, value: { ...p, pmccs: [...p.pmccs, { id: newId("pmcc"), tag: tag.trim(), fileName: sheet?.fileName, sheet: sheet?.sheet }] } };
}

export function deletePmcc(p: DesignProject, id: string): Result<DesignProject> {
  const users = p.imccs.filter((i) => i.incomer1.pmccId === id || i.incomer2.pmccId === id);
  if (users.length) return { ok: false, errors: [`This PMCC supplies ${users.map((u) => u.tag).join(", ")}: change or delete those IMCCs first.`] };
  return { ok: true, value: { ...p, pmccs: p.pmccs.filter((x) => x.id !== id) } };
}

/** Renaming a PMCC changes nothing else: links use ids, descriptions use the IMCC tag. */
export function renamePmcc(p: DesignProject, id: string, tag: string): Result<DesignProject> {
  if (!tag.trim()) return { ok: false, errors: ["The PMCC needs a tag."] };
  if (p.pmccs.some((x) => x.id !== id && x.tag.trim().toLowerCase() === tag.trim().toLowerCase())) return { ok: false, errors: [`PMCC tag "${tag}" is already used.`] };
  return { ok: true, value: { ...p, pmccs: p.pmccs.map((x) => (x.id === id ? { ...x, tag: tag.trim() } : x)) } };
}

/* ------------------------------ PMCC feeder list --------------------------- */

export interface PmccFeederRow {
  id: string;
  origin: "manual" | "imcc";
  description: string;
  destination: string;
  status?: LoadStatus;
  ratedA?: number;
  kw?: number;
  kva?: number;
  rating?: BreakerSelection;
  feederRef?: string;
  imccId?: string;
  incomer?: BusSection;
}

/** Manual feeders of a PMCC (its load-list rows, unchanged) followed by the generated IMCC incomer feeders. */
export function pmccFeeders(p: DesignProject, pmccId: string, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): PmccFeederRow[] {
  const pmcc = p.pmccs.find((x) => x.id === pmccId);
  if (!pmcc) return [];
  const manual: PmccFeederRow[] = [];
  if (pmcc.sheet) {
    for (const r of parseLoadSheet(pmcc.sheet as never).rows) {
      const type = r.type.toUpperCase().includes("VFD") ? "VFD" : r.type.toUpperCase().includes("DOL") ? "DOL" : "FEEDER";
      const cur = currentFor({ type, kw: r.kw, ratedA: r.amps, voltageV: r.voltageV }, a);
      manual.push({
        id: `row:${pmccId}:${r.rowNumber}`, origin: "manual", description: r.description, destination: "",
        status: r.status, ratedA: cur.usedA, kw: r.kw,
        kva: cur.usedA !== undefined && r.voltageV ? (SQ3 * calcV(r.voltageV, a) * cur.usedA) / 1000 : undefined,
        rating: cur.designA === undefined ? undefined : selectBreaker(cur.designA, { mccb630: a.mccb630 }),
      });
    }
  }
  const auto = p.generated.filter((g) => g.pmccId === pmccId).map((g): PmccFeederRow => ({
    id: g.id, origin: "imcc", description: g.description, destination: g.destination, status: g.status, ratedA: g.ratedA,
    kva: g.kva, rating: g.rating, feederRef: g.feederRef, imccId: g.imccId, incomer: g.incomer,
  }));
  return [...manual, ...auto];
}

export interface PmccSummary {
  feeders: number;
  manual: number;
  generated: number;
  /** Working (W) kVA only: standby (S) and spare (SP) are not normal working demand. */
  workingKva: number;
  standbyKva: number;
  spare: number;
}

export function pmccSummary(rows: PmccFeederRow[]): PmccSummary {
  const sum = (s: LoadStatus) => rows.filter((r) => r.status === s).reduce((t, r) => t + (r.kva ?? 0), 0);
  return {
    feeders: rows.length, manual: rows.filter((r) => r.origin === "manual").length, generated: rows.filter((r) => r.origin === "imcc").length,
    workingKva: sum("W"), standbyKva: sum("S"), spare: rows.filter((r) => r.status === "SP").length,
  };
}

/* ---------------------------- PMCC export sheet ---------------------------- */

const norm = (h: unknown) => String(h ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");

/**
 * The PMCC load list as exported: the original rows unchanged, then one row per linked IMCC incomer, with
 * "Source" and "Destination" columns. A PMCC with no IMCC link is returned exactly as uploaded.
 */
export function pmccExportSheet(p: DesignProject, pmccId: string, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): unknown[][] | undefined {
  const pmcc = p.pmccs.find((x) => x.id === pmccId);
  if (!pmcc?.sheet) return undefined;
  const sheet = pmcc.sheet.map((r) => [...r]);
  const links = p.generated.filter((g) => g.pmccId === pmccId);
  if (links.length === 0) return sheet;
  const headerIdx = sheet.findIndex((r) => r.some((c) => String(c ?? "").trim() !== ""));
  const header = sheet[headerIdx];
  const col = (test: (n: string) => boolean) => header.findIndex((h) => test(norm(h)));
  const c = {
    desc: col((n) => n.startsWith("desc")), type: col((n) => n.startsWith("type")), volt: col((n) => n.includes("volt")),
    amps: col((n) => n.includes("amp")), status: col((n) => n.includes("working") || n.includes("standby") || n.includes("duty") || n === "ws"),
  };
  const width = Math.max(...sheet.slice(headerIdx).map((r) => r.length));
  const srcCol = width, dstCol = width + 1;
  header[srcCol] = "Source"; header[dstCol] = "Destination";
  for (let r = headerIdx + 1; r < sheet.length; r++) { sheet[r][srcCol] = "Manual"; sheet[r][dstCol] = ""; }
  for (const g of links) {
    const row: unknown[] = Array(dstCol + 1).fill(null);
    const imcc = p.imccs.find((i) => i.id === g.imccId)!;
    if (c.desc >= 0) row[c.desc] = `${g.description} (to ${g.destination})`;
    if (c.type >= 0) row[c.type] = "Feeder";
    if (c.volt >= 0) row[c.volt] = imcc.voltageV;
    if (c.amps >= 0) row[c.amps] = g.ratedA > 0 ? g.ratedA : null;
    if (c.status >= 0) row[c.status] = g.status;
    row[srcCol] = "Auto from IMCC";
    row[dstCol] = `${g.destination} Incomer ${g.incomer}`;
    sheet.push(row);
  }
  void a;
  return sheet;
}
