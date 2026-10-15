/**
 * IMCC design (docs/imcc.md). Stand-alone: it works from its own IMCC load list. The only link to the rest of the
 * application is the "Automatic PMCC feeders" switch: when ON, each generated IMCC incomer is also an outgoing feeder
 * of the PMCC load list (`imccFeederRows`, `appendFeedersToSheet`); when OFF nothing outside the IMCC tab changes.
 *
 * Script / logic (`designImccs`): only working (W / w) loads are used for the design. They are taken in list order and
 * fill bus sections one after the other; a new section starts when adding the next load would push the section's design
 * current (working current x design multiplier) above the incomer limit (1000 A). Two sections make one IMCC, so the
 * number of IMCCs follows from the loads. Standby (S) and spare (SP) rows are placed with their group's (or preceding
 * working load's) section and are not counted. Current and breaker sizing reuse `currentFor` / `selectBreaker`.
 */
import { DEFAULT_MCC_ASSUMPTIONS, breakerLabel, currentFor, selectBreaker, type BreakerSelection, type MccAssumptions, type MccLoadType } from "./mcc-breaker";

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

/** SS = soft starter. */
export type StartMethod = "DOL" | "RDOL" | "SS" | "VFD" | "PWR FDR";

export function normalizeMethod(v: unknown): StartMethod | undefined {
  const t = String(v ?? "").trim().toUpperCase().replace(/\s+/g, " ");
  if (t === "DOL" || t === "RDOL" || t === "VFD" || t === "SS") return t;
  if (t === "SOFT STARTER" || t === "SOFTSTARTER") return "SS";
  if (t === "PWR FDR" || t === "POWER FEEDER" || t === "FEEDER" || t === "PWR FEEDER") return "PWR FDR";
  return undefined;
}

/** A soft-starter motor is calculated like a DOL motor (power factor and efficiency): assumption, OPEN_ITEMS.md #12. */
export const loadTypeOf = (m: StartMethod | undefined): MccLoadType => (m === "VFD" ? "VFD" : m === "DOL" || m === "RDOL" || m === "SS" ? "DOL" : "FEEDER");

export interface ImccLoad {
  id: string;
  /** Row in the source sheet (1-based), for traceability. */
  srcRow?: number;
  sNo?: string;
  description: string;
  kw?: number;
  ratedA?: number;
  /** Normalised status; undefined = missing or invalid (see `statusText`). */
  status?: LoadStatus;
  statusText?: string;
  method?: StartMethod;
  group?: string;
  note?: string;
}

let counter = 0;
export const newId = (prefix: string): string => `${prefix}_${Date.now().toString(36)}${(counter++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;

/* ------------------------------ operating state ---------------------------- */

export type SwitchState = "ON" | "OFF";
export interface OperatingState { incomer1: SwitchState; incomer2: SwitchState; coupler: SwitchState }
export interface Validation { ok: boolean; errors: string[]; warnings: string[] }

/** Operating philosophy as a rule. The default is the project's: exactly one incomer ON, coupler OFF. */
export interface OperatingRule {
  exactlyOneIncomerOn: boolean;
  couplerMustBeOff: boolean;
  /** A rule that differs from the default is an engineering change and needs a named authoriser. */
  authorisedBy?: string;
}
export const DEFAULT_RULE: OperatingRule = { exactlyOneIncomerOn: true, couplerMustBeOff: true };
export const NORMAL_STATE: OperatingState = { incomer1: "ON", incomer2: "OFF", coupler: "OFF" };

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
    errors.push(on === 1
      ? "Bus coupler is ON with one incomer ON: not permitted. One incomer is not assumed to supply both bus sections through the coupler."
      : "Bus coupler is ON: it must remain OFF.");
  }
  if (errors.length === 0 && on === 1 && s.coupler === "OFF") warnings.push("Permitted normal state: one incomer ON, the other incomer OFF, bus coupler OFF.");
  return { ok: errors.length === 0, errors, warnings };
}

/* --------------------------------- results -------------------------------- */

export interface LoadResult {
  load: ImccLoad;
  usedA?: number;
  designA?: number;
  kva?: number;
  /** Counted in the working load: status W only. */
  working: boolean;
  workingKva: number;
  workingA: number;
  breaker?: BreakerSelection;
}

export interface SectionResult {
  /** 1 or 2 within its IMCC. */
  section: 1 | 2;
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

export interface ImccDesign {
  tag: string;
  index: number;
  sections: [SectionResult, SectionResult];
  /** Calculated design rating of Incomer 1 (from its own section's working load). */
  incomer1: BreakerSelection;
  /** Incomer 2 and the bus coupler follow Incomer 1. */
  incomer2: BreakerSelection;
  coupler: BreakerSelection;
  issues: { severity: "error" | "warning"; message: string }[];
}

export interface ImccDesignSet {
  imccs: ImccDesign[];
  /** Problems that belong to no single IMCC (missing status, a load larger than one incomer ...). */
  issues: { severity: "error" | "warning"; message: string }[];
  maxIncomerA: number;
  voltageV: number;
}

export const LCE_FACTOR = 0.8;
export const MAX_INCOMER_A = 1000;
const SQ3 = Math.sqrt(3);
const calcV = (v: number, a: MccAssumptions) => (Math.round(v) === 433 ? a.calcVoltage433 : v);

function emptySection(section: 1 | 2): SectionResult {
  return {
    section, loads: [], working: 0, standby: 0, spare: 0, connectedKva: 0, workingKva: 0, workingA: 0, standbyKva: 0, lceKva: 0, designA: 0,
    breaker: { kind: "REVIEW", status: "review", message: "No design current available, so no breaker is selected." },
  };
}

function sectionOf(section: 1 | 2, loads: LoadResult[], a: MccAssumptions): SectionResult {
  const sum = (f: (x: LoadResult) => number) => loads.reduce((t, x) => t + f(x), 0);
  const st = (s: LoadStatus) => loads.filter((x) => x.load.status === s);
  const workingA = sum((x) => x.workingA);
  const designA = workingA * a.designMultiplier;
  const workingKva = sum((x) => x.workingKva);
  return {
    section, loads, working: st("W").length, standby: st("S").length, spare: st("SP").length,
    workingKva, workingA, designA, standbyKva: st("S").reduce((t, x) => t + (x.kva ?? 0), 0),
    connectedKva: [...st("W"), ...st("S")].reduce((t, x) => t + (x.kva ?? 0), 0),
    lceKva: workingKva * LCE_FACTOR,
    breaker: selectBreaker(workingA > 0 ? designA : undefined, { mccb630: a.mccb630 }),
  };
}

/** Current / kVA of one load with the application's conventions. */
export function loadResult(l: ImccLoad, voltageV: number, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): LoadResult {
  const cur = currentFor({ type: loadTypeOf(l.method), kw: l.kw, ratedA: l.ratedA, voltageV }, a);
  const kva = cur.usedA === undefined ? undefined : (SQ3 * calcV(voltageV, a) * cur.usedA) / 1000;
  const working = l.status === "W";
  return {
    load: l, usedA: cur.usedA, designA: cur.designA, kva, working,
    workingKva: working && kva !== undefined ? kva : 0,
    workingA: working && cur.usedA !== undefined ? cur.usedA : 0,
    breaker: cur.designA === undefined ? undefined : selectBreaker(cur.designA, { mccb630: a.mccb630 }),
  };
}

/**
 * The design script. Working loads in list order fill sections up to `maxIncomerA` of design current; two sections
 * per IMCC. Returns as many IMCCs as the loads need (none when there is no working load).
 */
export function designImccs(
  loads: readonly ImccLoad[],
  a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS,
  opts: { maxIncomerA?: number; voltageV?: number } = {},
): ImccDesignSet {
  const maxA = opts.maxIncomerA ?? MAX_INCOMER_A;
  const voltageV = opts.voltageV ?? 433;
  const issues: ImccDesignSet["issues"] = [];
  const valid = loads.filter((l) => l.status !== undefined);
  for (const l of loads) {
    if (l.status === undefined) issues.push({ severity: "error", message: `"${l.description}": status ${l.statusText ? `"${l.statusText}"` : "missing"}; use W, w, S or Spare. Not counted, not placed.` });
    else if (l.status !== "SP" && l.kw === undefined && l.ratedA === undefined) issues.push({ severity: "warning", message: `"${l.description}": no kW or current, so it carries no load in the calculation.` });
    if (l.status !== undefined && l.status !== "SP" && !l.method) issues.push({ severity: "warning", message: `"${l.description}": method of starting missing or not recognised (DOL, RDOL, SS, VFD, PWR FDR); treated as a power feeder.` });
  }
  const results = valid.map((l) => loadResult(l, voltageV, a));

  // 1. Working loads fill sections in list order.
  const secOf = new Map<string, number>();
  let k = 0;
  let acc = 0;
  for (const r of results) {
    if (!r.working) continue;
    const designIfAdded = (acc + r.workingA) * a.designMultiplier;
    if (designIfAdded > maxA + 1e-9 && acc > 0) { k++; acc = 0; }
    if (r.workingA * a.designMultiplier > maxA + 1e-9) {
      issues.push({ severity: "error", message: `"${r.load.description}": design current ${(r.workingA * a.designMultiplier).toFixed(0)} A alone exceeds the ${maxA} A incomer limit. It is placed alone in its own section: engineering review required.` });
    }
    acc += r.workingA;
    secOf.set(r.load.id, k);
  }
  const hasWorking = secOf.size > 0;
  // 2. Standby and spare rows go with their group's working load, else the nearest preceding working load.
  const groupSec = new Map<string, number>();
  for (const r of results) if (r.working && r.load.group?.trim()) { const g = r.load.group.trim().toLowerCase(); if (!groupSec.has(g)) groupSec.set(g, secOf.get(r.load.id)!); }
  let lastSec: number | undefined;
  const firstWorkingSec = hasWorking ? secOf.get(results.find((r) => r.working)!.load.id)! : 0;
  for (const r of results) {
    if (r.working) { lastSec = secOf.get(r.load.id)!; continue; }
    const g = r.load.group?.trim().toLowerCase();
    secOf.set(r.load.id, (g ? groupSec.get(g) : undefined) ?? lastSec ?? firstWorkingSec);
  }
  if (!hasWorking) {
    issues.push({ severity: "error", message: "No working (W / w) load: there is nothing to design an IMCC for." });
    return { imccs: [], issues, maxIncomerA: maxA, voltageV };
  }

  const nSections = Math.max(...secOf.values()) + 1;
  const bySection: LoadResult[][] = Array.from({ length: nSections + (nSections % 2) }, () => []);
  for (const r of results) bySection[secOf.get(r.load.id)!].push(r);

  const imccs: ImccDesign[] = [];
  for (let i = 0; i < bySection.length; i += 2) {
    const s1 = sectionOf(1, bySection[i], a);
    const s2 = bySection[i + 1].length ? sectionOf(2, bySection[i + 1], a) : emptySection(2);
    const incomer1 = s1.breaker;
    const tag = `IMCC-${String(i / 2 + 1).padStart(2, "0")}`;
    const imccIssues: ImccDesign["issues"] = [];
    if (incomer1.kind === "REVIEW") imccIssues.push({ severity: "warning", message: `Incomer 1: ${incomer1.message}` });
    if (s2.breaker.ratingA !== undefined && incomer1.ratingA !== undefined && s2.breaker.ratingA > incomer1.ratingA) {
      imccIssues.push({ severity: "warning", message: `Incomer 2 section needs ${breakerLabel(s2.breaker)}, more than Incomer 1's ${breakerLabel(incomer1)} which Incomer 2 and the bus coupler follow: review.` });
    }
    if (s2.workingA === 0) imccIssues.push({ severity: "warning", message: "Incomer 2 section has no working load: no PMCC feeder is generated for Incomer 2." });
    // Duty/standby group with standby but no working load in the section.
    for (const sec of [s1, s2]) {
      const groups = new Map<string, { w: number; s: number }>();
      for (const x of sec.loads) {
        const g = x.load.group?.trim();
        if (!g || x.load.status === "SP") continue;
        const e = groups.get(g) ?? { w: 0, s: 0 };
        if (x.load.status === "W") e.w++; else e.s++;
        groups.set(g, e);
      }
      for (const [g, e] of groups) if (e.w === 0 && e.s > 0) imccIssues.push({ severity: "warning", message: `${tag} Incomer ${sec.section}: duty/standby group "${g}" has standby but no working load.` });
    }
    imccs.push({ tag, index: i / 2 + 1, sections: [s1, s2], incomer1, incomer2: { ...incomer1 }, coupler: { ...incomer1 }, issues: imccIssues });
  }
  return { imccs, issues, maxIncomerA: maxA, voltageV };
}

/* --------------------- the "Automatic PMCC feeders" switch ------------------ */

/** The PMCC outgoing feeder generated for one IMCC incomer. */
export interface ImccFeederRow {
  description: string;
  imccTag: string;
  incomer: 1 | 2;
  /** Working current of the incomer's section (A). */
  amps: number;
  kva: number;
  voltageV: number;
  /** W for Incomer 1 (normally ON), S for Incomer 2 (alternative source, normally OFF). */
  status: "W" | "S";
  /** The IMCC tag: the two incomers go to different transformer sections (existing duty/standby group rule). */
  group: string;
}

/**
 * One outgoing feeder per IMCC incomer that has a working load. The two incomers are alternative sources: Incomer 1
 * is working (W), Incomer 2 standby (S), and both carry the group = IMCC tag so the transformer sizing places them on
 * different bus sections. Sized from the section's working current, never from the incomer breaker rating.
 */
export function imccFeederRows(set: ImccDesignSet): ImccFeederRow[] {
  const out: ImccFeederRow[] = [];
  for (const i of set.imccs) {
    for (const n of [1, 2] as const) {
      const s = i.sections[n - 1];
      if (s.workingA <= 0) continue;
      out.push({ description: `${i.tag} Incomer ${n}`, imccTag: i.tag, incomer: n, amps: s.workingA, kva: s.workingKva, voltageV: set.voltageV, status: n === 1 ? "W" : "S", group: i.tag });
    }
  }
  return out;
}

const norm = (h: unknown) => String(h ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");

/**
 * The PMCC load list with the IMCC incomer feeders added as rows (Type "Feeder", Amps, voltage, W/S, group). The
 * original rows are unchanged. Columns for W/S and group are added at the end when the sheet has none. A sheet whose
 * required columns are missing, or with no feeders to add, is returned unchanged.
 */
export function appendFeedersToSheet(sheet: readonly (readonly unknown[])[], feeders: readonly ImccFeederRow[]): unknown[][] {
  const out = sheet.map((r) => [...r]);
  if (feeders.length === 0) return out;
  const headerIdx = out.findIndex((r) => r.some((c) => String(c ?? "").trim() !== ""));
  if (headerIdx < 0) return out;
  const header = out[headerIdx];
  const col = (test: (n: string) => boolean) => header.findIndex((h) => test(norm(h)));
  const c = {
    desc: col((n) => n.startsWith("desc")), type: col((n) => n.startsWith("type")), volt: col((n) => n.includes("volt")),
    amps: col((n) => n.includes("amp")),
    status: col((n) => !n.includes("group") && !n.includes("pair") && (n.includes("working") || n.includes("standby") || n.includes("duty") || n === "ws")),
    group: col((n) => n.includes("group") || n.includes("pair")),
  };
  if ([c.desc, c.type, c.volt, c.amps].some((i) => i < 0)) return out;
  let width = Math.max(...out.slice(headerIdx).map((r) => r.length));
  if (c.status < 0) { c.status = width++; header[c.status] = "W/S"; }
  if (c.group < 0) { c.group = width++; header[c.group] = "Duty/standby group"; }
  for (const f of feeders) {
    const row: unknown[] = Array(width).fill(null);
    row[c.desc] = f.description;
    row[c.type] = "Feeder";
    row[c.amps] = f.amps;
    row[c.volt] = f.voltageV;
    row[c.status] = f.status;
    row[c.group] = f.group;
    out.push(row);
  }
  return out;
}
