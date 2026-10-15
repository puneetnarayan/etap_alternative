/**
 * IMCC design (docs/imcc.md). Stand-alone: it works from its own IMCC load list. The only link to the rest of the
 * application is the "Automatic PMCC feeders" switch: when ON, each IMCC incomer is also an outgoing feeder of the PMCC
 * load list (`imccFeederRows`, `appendFeedersToSheet`); when OFF nothing outside the IMCC tab changes.
 *
 * Panel: Incomer 1 at the beginning; Incomer 2 and the bus coupler side by side at the other end. One bus carries all
 * the IMCC's loads. Normal state: Incomer 1 ON, Incomer 2 and bus coupler OFF. If Incomer 1 is unhealthy, Incomer 2 and the
 * bus coupler are both switched ON, with an interlock that switches Incomer 1 OFF first if it is ON.
 * So Incomer 1, Incomer 2 and the bus coupler each carry the whole IMCC and have the same calculated rating.
 *
 * Script (`designImccs`): only working (W / w) loads are used for the design. They are taken in list order and fill an
 * IMCC until adding the next load would push its design current (working current x design multiplier) above the ACB
 * limit: the incomer ACB chosen by the user (630, 800, 1000 or 1250 A; default 1000 A) at its maximum loading (90 %, adjustable). Then the next IMCC starts, so the
 * number of IMCCs follows from the loads. Standby (S) and spare (SP) rows are placed with their group's (or the
 * preceding working load's) IMCC and are not counted. Current and breaker sizing reuse `currentFor` / `selectBreaker`.
 */
import { DEFAULT_MCC_ASSUMPTIONS, currentFor, selectBreaker, type BreakerSelection, type MccAssumptions, type MccLoadType } from "./mcc-breaker";

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

/** SS = soft starter. "PWR FDR" in a list is read as Feeder. */
export type StartMethod = "DOL" | "RDOL" | "SS" | "VFD" | "Feeder";

export function normalizeMethod(v: unknown): StartMethod | undefined {
  const t = String(v ?? "").trim().toUpperCase().replace(/\s+/g, " ");
  if (t === "DOL" || t === "RDOL" || t === "VFD" || t === "SS") return t;
  if (t === "SOFT STARTER" || t === "SOFTSTARTER") return "SS";
  if (t === "FEEDER" || t === "PWR FDR" || t === "POWER FEEDER" || t === "PWR FEEDER") return "Feeder";
  return undefined;
}

/** A soft-starter motor is calculated like a DOL motor (pf and efficiency): confirmed, OPEN_ITEMS.md #12. */
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
  /** Voltage level of the row (V), when the list gives one. A row at another voltage than the IMCC's is not placed. */
  voltageV?: number;
  /** Columns whose value could not be read as a number (e.g. "-"). */
  badFields?: string[];
}

let counter = 0;
export const newId = (prefix: string): string => `${prefix}_${Date.now().toString(36)}${(counter++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;

/* ------------------------------ operating state ---------------------------- */

export type SwitchState = "ON" | "OFF";
export interface OperatingState { incomer1: SwitchState; incomer2: SwitchState; coupler: SwitchState }
export interface Validation { ok: boolean; errors: string[]; warnings: string[] }

/** Normal: Incomer 1 ON, Incomer 2 and bus coupler OFF. */
export const NORMAL_STATE: OperatingState = { incomer1: "ON", incomer2: "OFF", coupler: "OFF" };
/** Incomer 1 unhealthy: Incomer 1 OFF, Incomer 2 and bus coupler both ON. */
export const FAILOVER_STATE: OperatingState = { incomer1: "OFF", incomer2: "ON", coupler: "ON" };

/** Operating philosophy as a rule. The default: normal state, and the Incomer 1 failover with interlock. */
export interface OperatingRule {
  /** Incomer 2 and the bus coupler are switched ON when Incomer 1 is unhealthy. */
  failoverOnIncomer1Unhealthy: boolean;
  /** A rule that differs from the default is an engineering change and needs a named authoriser. */
  authorisedBy?: string;
}
export const DEFAULT_RULE: OperatingRule = { failoverOnIncomer1Unhealthy: true };

export function validateOperatingState(s: OperatingState, rule: OperatingRule = DEFAULT_RULE): Validation {
  const errors: string[] = [];
  const warnings: string[] = [];
  if (rule.failoverOnIncomer1Unhealthy !== DEFAULT_RULE.failoverOnIncomer1Unhealthy && !rule.authorisedBy?.trim()) {
    errors.push("The operating philosophy differs from the default (Incomer 2 and the bus coupler ON when Incomer 1 is unhealthy). That is an engineering change and needs a named authoriser.");
  }
  if (s.incomer1 === "ON" && (s.incomer2 === "ON" || s.coupler === "ON")) {
    errors.push(`Interlock: Incomer 1 is ON, so ${s.incomer2 === "ON" && s.coupler === "ON" ? "Incomer 2 and the bus coupler" : s.incomer2 === "ON" ? "Incomer 2" : "the bus coupler"} must stay OFF. Incomer 1 has to be switched OFF first.`);
  }
  if (s.incomer1 === "OFF") {
    if (s.incomer2 === "ON" && s.coupler === "OFF") errors.push("Incomer 2 is ON with the bus coupler OFF: they are switched ON together (Incomer 2 and the bus coupler).");
    else if (s.incomer2 === "OFF" && s.coupler === "ON") errors.push("The bus coupler is ON without Incomer 2: they are switched ON together.");
    else if (s.incomer2 === "OFF" && s.coupler === "OFF") errors.push("No source is ON: the IMCC is not supplied.");
    else if (!rule.failoverOnIncomer1Unhealthy) errors.push("Incomer 2 and the bus coupler ON is not permitted by the operating philosophy in use.");
  }
  if (errors.length === 0) {
    warnings.push(s.incomer1 === "ON"
      ? "Normal state: Incomer 1 ON, Incomer 2 and bus coupler OFF."
      : "Incomer 1 unhealthy: Incomer 1 OFF, Incomer 2 and bus coupler ON. Each carries the whole IMCC.");
  }
  return { ok: errors.length === 0, errors, warnings };
}

/**
 * Incomer 1 unhealthy: the steps to take from `from`. The interlock switches Incomer 1 OFF first if it is ON, then
 * Incomer 2 and the bus coupler are switched ON together.
 */
export function switchToFailover(from: OperatingState): { state: OperatingState; steps: string[] } {
  const steps: string[] = [];
  if (from.incomer1 === "ON") steps.push("Interlock: Incomer 1 switched OFF");
  if (from.incomer2 === "OFF") steps.push("Incomer 2 switched ON");
  if (from.coupler === "OFF") steps.push("Bus coupler switched ON");
  return { state: { ...FAILOVER_STATE }, steps: steps.length ? steps : ["Already in the Incomer 1 unhealthy state"] };
}

/** Incomer 1 healthy again: the reverse order: Incomer 2 and the bus coupler OFF first, then Incomer 1 ON. */
export function switchToNormal(from: OperatingState): { state: OperatingState; steps: string[] } {
  const steps: string[] = [];
  if (from.incomer2 === "ON") steps.push("Incomer 2 switched OFF");
  if (from.coupler === "ON") steps.push("Bus coupler switched OFF");
  if (from.incomer1 === "OFF") steps.push("Incomer 1 switched ON");
  return { state: { ...NORMAL_STATE }, steps: steps.length ? steps : ["Already in the normal state"] };
}

/* --------------------------------- results -------------------------------- */

export interface LoadResult {
  load: ImccLoad;
  /** Current entered in the list (A), if any. */
  enteredA?: number;
  /** Calculated from kW (A): shown beside the entered one to explain the selection. */
  calculatedA?: number;
  /** Entered current if given, else the calculated one. */
  usedA?: number;
  designA?: number;
  kva?: number;
  /** Counted in the working load: status W only. */
  working: boolean;
  workingKva: number;
  workingA: number;
  breaker?: BreakerSelection;
}

export interface BusResult {
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
}

export interface ImccDesign {
  tag: string;
  index: number;
  /** The IMCC bus: all its loads. */
  bus: BusResult;
  /** Incomer 1: the selected incomer ACB. */
  incomer1: BreakerSelection;
  /** Incomer 2 and the bus coupler carry the whole IMCC when Incomer 1 is unhealthy: same rating. */
  incomer2: BreakerSelection;
  coupler: BreakerSelection;
  /** Design current / rating of the incomers (fraction); undefined when there is no rating. */
  loading?: number;
  /** The ACB maximum loading that applies (fraction). */
  loadingLimit: number;
  issues: { severity: "error" | "warning"; message: string }[];
}

export interface ImccDesignSet {
  imccs: ImccDesign[];
  /** Problems that belong to no single IMCC (missing status, a load larger than the limit ...). */
  issues: { severity: "error" | "warning"; message: string }[];
  /** The selected incomer ACB size (A): 630, 800, 1000 or 1250. */
  maxIncomerA: number;
  /** Maximum loading of an ACB (fraction), adjustable. */
  acbMaxLoading: number;
  /** Largest design current one IMCC takes: maxIncomerA x acbMaxLoading. */
  maxDesignA: number;
  voltageV: number;
}

export const LCE_FACTOR = 0.8;
export const DEFAULT_ACB_MAX_LOADING = 0.9;
const SQ3 = Math.sqrt(3);
const calcV = (v: number, a: MccAssumptions) => (Math.round(v) === 433 ? a.calcVoltage433 : v);

/** The incomer ACB sizes the user can choose (A). */
export const INCOMER_ACB_OPTIONS = [630, 800, 1000, 1250] as const;
export const DEFAULT_INCOMER_ACB_A = 1000;
/** Kept for callers that want the default size. */
export const MAX_INCOMER_A = DEFAULT_INCOMER_ACB_A;

/**
 * The incomer / bus-coupler device: the EDO ACB chosen by the user. It carries the whole IMCC, so Incomer 1, Incomer 2 and
 * the bus coupler are all this size. Without a working load there is nothing to rate.
 */
export function incomerAcb(ratingA: number, hasLoad: boolean): BreakerSelection {
  if (!hasLoad) return { kind: "REVIEW", status: "review", message: "No design current available, so no breaker is selected." };
  return { kind: "EDO ACB", ratingA, status: "ok", message: `${ratingA} A EDO ACB (selected incomer size).` };
}

/** Current / kVA of one load with the application's conventions. */
export function loadResult(l: ImccLoad, voltageV: number, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): LoadResult {
  const cur = currentFor({ type: loadTypeOf(l.method), kw: l.kw, ratedA: l.ratedA, voltageV }, a);
  const kva = cur.usedA === undefined ? undefined : (SQ3 * calcV(voltageV, a) * cur.usedA) / 1000;
  const working = l.status === "W";
  return {
    load: l, enteredA: l.ratedA, calculatedA: cur.calculatedA, usedA: cur.usedA, designA: cur.designA, kva, working,
    workingKva: working && kva !== undefined ? kva : 0,
    workingA: working && cur.usedA !== undefined ? cur.usedA : 0,
    breaker: cur.designA === undefined ? undefined : selectBreaker(cur.designA, { mccb630: a.mccb630 }),
  };
}

function busOf(loads: LoadResult[], a: MccAssumptions): BusResult {
  const sum = (f: (x: LoadResult) => number) => loads.reduce((t, x) => t + f(x), 0);
  const st = (s: LoadStatus) => loads.filter((x) => x.load.status === s);
  const workingA = sum((x) => x.workingA);
  const workingKva = sum((x) => x.workingKva);
  return {
    loads, working: st("W").length, standby: st("S").length, spare: st("SP").length,
    workingKva, workingA, designA: workingA * a.designMultiplier,
    standbyKva: st("S").reduce((t, x) => t + (x.kva ?? 0), 0),
    connectedKva: [...st("W"), ...st("S")].reduce((t, x) => t + (x.kva ?? 0), 0),
    lceKva: workingKva * LCE_FACTOR,
  };
}

/**
 * The design script. Working loads in list order fill an IMCC up to `maxIncomerA x acbMaxLoading` of design current;
 * then the next IMCC starts. Returns as many IMCCs as the loads need (none when there is no working load).
 */
export function designImccs(
  loads: readonly ImccLoad[],
  a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS,
  opts: { maxIncomerA?: number; acbMaxLoading?: number; voltageV?: number } = {},
): ImccDesignSet {
  const maxA = opts.maxIncomerA ?? MAX_INCOMER_A;
  const acbMaxLoading = opts.acbMaxLoading ?? DEFAULT_ACB_MAX_LOADING;
  const maxDesignA = maxA * acbMaxLoading;
  const voltageV = opts.voltageV ?? 433;
  const issues: ImccDesignSet["issues"] = [];
  for (const l of loads) {
    if (l.voltageV !== undefined && Math.round(l.voltageV) !== Math.round(voltageV)) issues.push({ severity: "error", message: `"${l.description}": ${l.voltageV} V differs from the IMCC voltage ${voltageV} V. Not counted, not placed.` });
    if (l.badFields?.length) issues.push({ severity: "warning", message: `"${l.description}": could not read a number in ${l.badFields.join(", ")}; ${l.badFields.length >= 2 ? "no current" : "the other value"} is used where available.` });
  }
  const valid = loads.filter((l) => l.status !== undefined && (l.voltageV === undefined || Math.round(l.voltageV) === Math.round(voltageV)));
  for (const l of loads) {
    if (l.status === undefined) issues.push({ severity: "error", message: `"${l.description}": status ${l.statusText ? `"${l.statusText}"` : "missing"}; use W, w, S or Spare. Not counted, not placed.` });
    else if (l.status !== "SP" && l.kw === undefined && l.ratedA === undefined) issues.push({ severity: "warning", message: `"${l.description}": no kW or current, so it carries no load in the calculation.` });
    if (l.status !== undefined && l.status !== "SP" && !l.method) issues.push({ severity: "warning", message: `"${l.description}": method of starting missing or not recognised (DOL, RDOL, SS, VFD, Feeder); treated as a feeder.` });
  }
  const results = valid.map((l) => loadResult(l, voltageV, a));

  // 1. Working loads fill IMCCs in list order.
  const imccOf = new Map<string, number>();
  let k = 0;
  let acc = 0;
  for (const r of results) {
    if (!r.working) continue;
    if ((acc + r.workingA) * a.designMultiplier > maxDesignA + 1e-9 && acc > 0) { k++; acc = 0; }
    if (r.workingA * a.designMultiplier > maxDesignA + 1e-9) {
      issues.push({ severity: "error", message: `"${r.load.description}": design current ${(r.workingA * a.designMultiplier).toFixed(0)} A alone exceeds ${maxDesignA.toFixed(0)} A (${maxA} A ACB at ${(acbMaxLoading * 100).toFixed(0)} % maximum loading). It is placed alone in its own IMCC: engineering review required.` });
    }
    acc += r.workingA;
    imccOf.set(r.load.id, k);
  }
  if (imccOf.size === 0) {
    issues.push({ severity: "error", message: "No working (W / w) load: there is nothing to design an IMCC for." });
    return { imccs: [], issues, maxIncomerA: maxA, acbMaxLoading, maxDesignA, voltageV };
  }
  // 2. Standby and spare rows go with their group's IMCC, else the nearest preceding working load's.
  const groupImcc = new Map<string, number>();
  for (const r of results) if (r.working && r.load.group?.trim()) { const g = r.load.group.trim().toLowerCase(); if (!groupImcc.has(g)) groupImcc.set(g, imccOf.get(r.load.id)!); }
  let last: number | undefined;
  const first = imccOf.get(results.find((r) => r.working)!.load.id)!;
  for (const r of results) {
    if (r.working) { last = imccOf.get(r.load.id)!; continue; }
    const g = r.load.group?.trim().toLowerCase();
    imccOf.set(r.load.id, (g ? groupImcc.get(g) : undefined) ?? last ?? first);
  }

  const n = Math.max(...imccOf.values()) + 1;
  const byImcc: LoadResult[][] = Array.from({ length: n }, () => []);
  for (const r of results) byImcc[imccOf.get(r.load.id)!].push(r);

  const imccs: ImccDesign[] = byImcc.map((ls, i) => {
    const bus = busOf(ls, a);
    const tag = `IMCC-${String(i + 1).padStart(2, "0")}`;
    const incomer1 = incomerAcb(maxA, bus.workingA > 0);
    const imccIssues: ImccDesign["issues"] = [];
    if (bus.workingA > 0 && bus.designA > maxDesignA + 1e-9) imccIssues.push({ severity: "warning", message: `${tag}: design current ${bus.designA.toFixed(0)} A is above ${maxDesignA.toFixed(0)} A (${maxA} A ACB at ${(acbMaxLoading * 100).toFixed(0)} % maximum loading): engineering review required.` });
    const groups = new Map<string, { w: number; s: number }>();
    for (const x of bus.loads) {
      const g = x.load.group?.trim();
      if (!g || x.load.status === "SP") continue;
      const e = groups.get(g) ?? { w: 0, s: 0 };
      if (x.load.status === "W") e.w++; else e.s++;
      groups.set(g, e);
    }
    for (const [g, e] of groups) if (e.w === 0 && e.s > 0) imccIssues.push({ severity: "warning", message: `${tag}: duty/standby group "${g}" has standby but no working load.` });
    return {
      tag, index: i + 1, bus, incomer1, incomer2: { ...incomer1 }, coupler: { ...incomer1 },
      loading: incomer1.ratingA ? bus.designA / incomer1.ratingA : undefined,
      loadingLimit: acbMaxLoading,
      issues: imccIssues,
    };
  });
  return { imccs, issues, maxIncomerA: maxA, acbMaxLoading, maxDesignA, voltageV };
}

/* --------------------- the "Automatic PMCC feeders" switch ------------------ */

/** The PMCC outgoing feeder generated for one IMCC incomer. */
export interface ImccFeederRow {
  description: string;
  imccTag: string;
  incomer: 1 | 2;
  /** Working current of the IMCC (A): each incomer carries the whole IMCC. */
  amps: number;
  kva: number;
  voltageV: number;
  /** W for Incomer 1 (normally ON), S for Incomer 2 (takes over when Incomer 1 is unhealthy). */
  status: "W" | "S";
  /** The IMCC tag: the two incomers go to different transformer sections (existing duty/standby group rule). */
  group: string;
}

/**
 * Two outgoing feeders per IMCC with a working load, one for each incomer. Incomer 1 is working (W), Incomer 2 standby (S):
 * only one of them carries the IMCC at a time, so the working demand is counted once; for transformer sizing the
 * standby feeder is counted too, as every standby load is. Both carry the IMCC tag as group, so the sizing places
 * them on different transformer sections. Sized from the IMCC's working current, never from the incomer breaker rating.
 */
export function imccFeederRows(set: ImccDesignSet): ImccFeederRow[] {
  const out: ImccFeederRow[] = [];
  for (const i of set.imccs) {
    if (i.bus.workingA <= 0) continue;
    for (const n of [1, 2] as const) {
      out.push({ description: `${i.tag} Incomer ${n}`, imccTag: i.tag, incomer: n, amps: i.bus.workingA, kva: i.bus.workingKva, voltageV: set.voltageV, status: n === 1 ? "W" : "S", group: i.tag });
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

