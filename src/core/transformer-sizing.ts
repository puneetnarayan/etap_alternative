/**
 * Transformer count / rating selection for a load list (see docs/transformer-sizing.md).
 *
 * Units: volts (V), amps (A), kW, kVA. All loads are counted at 100 % (standby included).
 */

import { fmtInt, fmtPct } from "./format";

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
  /**
   * Number(s) of transformers to start with: 3 = a set of 3 (3 incomers + 2 couplers), 2 = a set of 2
   * (2 incomers + 1 coupler). The default is to start with 3; a set of 2 is used first only on request,
   * or at the end as the remainder (see `layoutFor`).
   */
  counts: number[];
  /**
   * If the loads cannot be balanced within the limits with `counts`, more transformers are added
   * (more sets of 3 first, a set of 2 at the end for the remainder) until they can or this many
   * transformers is reached (see `layoutFor`).
   */
  maxTransformers: number;
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
  /**
   * Voltage (V) used to turn a current into kVA (sqrt(3) x V x A) for loads given only in amps, by
   * nominal voltage level. A level that is not listed uses its own value. 433 V loads are calculated
   * at 420 V; the 433 V label, and matching loads to the 433 V transformer, are unchanged.
   */
  calculationVoltageV: Record<number, number>;
}

/** Standard transformer ratings offered in the UI (kVA). */
export const STANDARD_RATINGS_KVA = [2000, 2500, 4000];

export const DEFAULT_OPTIONS: SizingOptions = {
  efficiency: 0.95,
  powerFactor: { DOL: 0.7, VFD: 0.9, FEEDER: 0.7 },
  normalLimit: 0.8,
  contingencyLimit: 0.9,
  ratingsKva: STANDARD_RATINGS_KVA,
  counts: [3],
  maxTransformers: 20,
  stepUp: "min-installed",
  totalBasis: "arithmetic",
  calculationVoltageV: { 433: 420 },
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
  const calcV = opts.calculationVoltageV[voltageV] ?? voltageV;
  const kva = (SQRT3 * calcV * amps) / 1000;
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
      basis: `√3 × ${calcV} V × ${amps} A (kW shown at pf ${pf})`,
      group: row.group,
      status: row.status,
      warnings,
    },
  };
}

/**
 * Split n transformers into independent sets of 2 or 3 transformers. Each set is its own switchboard:
 * 2 incomers + 1 bus coupler (set of 2) or 3 incomers + 2 bus couplers (set of 3), with one coupler
 * between neighbouring incomers and no coupler between sets. The split with the fewest sets
 * (most sets of 3) is used because it carries the most load: 2, 3, 2+2, 3+2, 3+3, 3+2+2, 3+3+2, ...
 */
export function layoutFor(count: number): number[] {
  if (!Number.isInteger(count) || count < 2) return [];
  const threes = count % 3 === 0 ? count / 3 : count % 3 === 1 ? (count - 4) / 3 : (count - 2) / 3;
  const twos = (count - 3 * threes) / 2;
  return [...Array<number>(threes).fill(3), ...Array<number>(twos).fill(2)];
}

/** Largest load (kVA) a set of `size` transformers of `ratingKva` can carry within both limits. */
export function setCapacityKva(size: number, ratingKva: number, opts: SizingOptions = DEFAULT_OPTIONS): number {
  return Math.min(opts.normalLimit * size * ratingKva, opts.contingencyLimit * (size - 1) * ratingKva);
}

export interface Candidate {
  count: number;
  /** Sizes of the independent sets, e.g. [3, 2] for 5 transformers. */
  layout: number[];
  ratingKva: number;
  installedKva: number;
  /** Average loading of all transformers in normal running, % of rating. */
  normalPct: number;
  /**
   * Loading of each remaining transformer with one out in each set, % of rating, with the load
   * shared in proportion to what each set can carry (so the figure is the same in every set).
   */
  contingencyPct: number;
  /** The total load is within the normal-running / one-out limits if shared perfectly. */
  normalOk: boolean;
  contingencyOk: boolean;
  /**
   * The loads, placed whole, can be balanced so that every section and every set is within its limit.
   * (Always true for a candidate that has not been tried with real loads yet.)
   */
  balanced: boolean;
  /** Fits = within both limits when shared perfectly AND the loads can really be balanced. */
  fits: boolean;
  /** Largest total load (kVA) this configuration can carry within both limits if shared perfectly. */
  maxLoadKva: number;
}

/** Why a candidate does not fit, in words ("" when it fits). */
export function failureReason(c: Candidate): string {
  const over = [!c.normalOk && "normal", !c.contingencyOk && "one out"].filter(Boolean);
  if (over.length > 0) return `${over.join(" + ")} over limit`;
  if (!c.balanced) return "loads cannot be balanced within the limits";
  return "";
}

export function evaluateCandidate(
  totalKva: number,
  count: number,
  ratingKva: number,
  opts: SizingOptions = DEFAULT_OPTIONS,
): Candidate {
  const layout = layoutFor(count);
  const sets = layout.length;
  const normalPct = (totalKva / (count * ratingKva)) * 100;
  const contingencyPct = (totalKva / ((count - sets) * ratingKva)) * 100;
  const normalOk = normalPct <= opts.normalLimit * 100 + 1e-9;
  const contingencyOk = contingencyPct <= opts.contingencyLimit * 100 + 1e-9;
  const maxLoadKva = layout.reduce((a, k) => a + setCapacityKva(k, ratingKva, opts), 0);
  return {
    count,
    layout,
    ratingKva,
    installedKva: count * ratingKva,
    normalPct,
    contingencyPct,
    normalOk,
    contingencyOk,
    balanced: true,
    fits: totalKva <= maxLoadKva + 1e-9,
    maxLoadKva,
  };
}

/** Order candidates by preference: lowest installed kVA first, or (count-first) by rating then count. */
function orderCandidates(all: Candidate[], opts: SizingOptions): Candidate[] {
  if (opts.stepUp === "min-installed") {
    // Stable sort keeps the rating/count order for equal installed kVA.
    return [...all].sort((a, b) => a.installedKva - b.installedKva);
  }
  return [...all].sort((a, b) => a.ratingKva - b.ratingKva || a.count - b.count);
}

/** Ideal-sharing candidates (no real load placement) for each count × rating, in order of preference. */
export function candidatesInOrder(totalKva: number, opts: SizingOptions = DEFAULT_OPTIONS): Candidate[] {
  const all: Candidate[] = [];
  for (const rating of opts.ratingsKva) {
    for (const count of opts.counts) all.push(evaluateCandidate(totalKva, count, rating, opts));
  }
  return orderCandidates(all, opts);
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
  /** "T1", "T2", ... for a single set; "S1-T1", "S2-T2", ... when there are several sets. */
  name: string;
  loads: SizedLoad[];
  /** % of rated kVA in normal running. */
  loadingPct: number;
  withinNormalLimit: boolean;
}

const groupKey = (g?: string) => (g ?? "").trim().toLowerCase();
const EPS = 1e-9;

/** Small deterministic random generator, so the same loads always give the same placement. */
function mulberry32(seed: number) {
  let a = seed | 0;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Upper limit on candidate placements examined per call, so that hopeless cases stay fast. */
const SEARCH_BUDGET = 1_000_000;

/**
 * Search for a placement where every section is within its normal limit and every set is within what
 * it can carry with one transformer out. Starting from the given placement it repeatedly moves a
 * load to another section, or swaps two loads, whenever that reduces the total excess over the
 * limits; if that gets stuck it perturbs the best placement found and climbs again (seeded, so the
 * result is repeatable). Duty/standby group members never share a section (when the group is small
 * enough for that to be possible). Changes `assign` in place to the best placement found; this is a
 * heuristic, so "cannot be balanced" means none was found, not a proof that none exists.
 */
function improveBalance(
  items: SizedLoad[],
  assign: number[],
  setOf: number[],
  sectionCap: number,
  setCap: number[],
  constrained: (key: string) => boolean,
  seed: number,
): void {
  const n = setOf.length;
  const keys = items.map((l) => {
    const k = groupKey(l.group);
    return k && constrained(k) ? k : "";
  });
  const cs = (kva: number) => Math.max(0, kva - sectionCap);
  const ct = (si: number, kva: number) => Math.max(0, kva - setCap[si]);

  let at: number[] = [];
  let sec: number[] = [];
  let set: number[] = [];
  let members: number[][] = [];
  let groups: Map<string, number>[] = [];
  const load = (from: number[]) => {
    at = from.slice();
    sec = Array<number>(n).fill(0);
    set = Array<number>(setCap.length).fill(0);
    members = Array.from({ length: n }, () => []);
    groups = Array.from({ length: n }, () => new Map());
    items.forEach((l, i) => {
      const t = at[i];
      sec[t] += l.kva;
      set[setOf[t]] += l.kva;
      members[t].push(i);
      if (keys[i]) groups[t].set(keys[i], (groups[t].get(keys[i]) ?? 0) + 1);
    });
  };
  const move = (i: number, to: number) => {
    const from = at[i];
    const x = items[i].kva;
    sec[from] -= x;
    sec[to] += x;
    set[setOf[from]] -= x;
    set[setOf[to]] += x;
    members[from].splice(members[from].indexOf(i), 1);
    members[to].push(i);
    if (keys[i]) {
      groups[from].set(keys[i], (groups[from].get(keys[i]) ?? 1) - 1);
      groups[to].set(keys[i], (groups[to].get(keys[i]) ?? 0) + 1);
    }
    at[i] = to;
  };
  const excess = () => sec.reduce((a, v) => a + cs(v), 0) + set.reduce((a, v, si) => a + ct(si, v), 0);
  const over = (t: number) => sec[t] > sectionCap + EPS || set[setOf[t]] > setCap[setOf[t]] + EPS;

  let budget = SEARCH_BUDGET;
  const climb = () => {
    for (let iter = 0; iter < 1000 && budget > 0; iter++) {
      if (!sec.some((_, t) => over(t))) return;
      let best = { d: -EPS, a: -1, b: -1, t: -1 };
      for (let a = 0; a < items.length; a++) {
        const f = at[a];
        if (!over(f)) continue;
        const x = items[a].kva;
        const ka = keys[a];
        for (let t = 0; t < n; t++) {
          if (t === f) continue;
          const sf = setOf[f];
          const st = setOf[t];
          budget--;
          if (!(ka && (groups[t].get(ka) ?? 0) > 0)) { // move a: f -> t
            let d = cs(sec[f] - x) - cs(sec[f]) + cs(sec[t] + x) - cs(sec[t]);
            if (sf !== st) d += ct(sf, set[sf] - x) - ct(sf, set[sf]) + ct(st, set[st] + x) - ct(st, set[st]);
            if (d < best.d) best = { d, a, b: -1, t };
          }
          for (const b of members[t]) { // swap a (in f) with b (in t)
            budget--;
            const y = items[b].kva;
            const kb = keys[b];
            if (ka && (groups[t].get(ka) ?? 0) - (kb === ka ? 1 : 0) > 0) continue;
            if (kb && (groups[f].get(kb) ?? 0) - (ka === kb ? 1 : 0) > 0) continue;
            let d = cs(sec[f] - x + y) - cs(sec[f]) + cs(sec[t] + x - y) - cs(sec[t]);
            if (sf !== st) d += ct(sf, set[sf] - x + y) - ct(sf, set[sf]) + ct(st, set[st] + x - y) - ct(st, set[st]);
            if (d < best.d) best = { d, a, b, t };
          }
        }
      }
      if (best.a < 0) return;
      const f = at[best.a];
      move(best.a, best.t);
      if (best.b >= 0) move(best.b, f);
    }
  };

  load(assign);
  climb();
  let bestAt = at.slice();
  let bestE = excess();
  const rnd = mulberry32(seed);
  for (let r = 0; r < 300 && budget > 0 && bestE > EPS && n > 1; r++) {
    load(bestAt);
    const kicks = 2 + Math.floor(rnd() * 3);
    for (let k = 0; k < kicks; k++) {
      const i = Math.floor(rnd() * items.length);
      const to = Math.floor(rnd() * n);
      if (to === at[i] || (keys[i] && (groups[to].get(keys[i]) ?? 0) > 0)) continue;
      move(i, to);
    }
    climb();
    const e = excess();
    if (e < bestE - EPS) {
      bestAt = at.slice();
      bestE = e;
    }
  }
  for (let i = 0; i < items.length; i++) assign[i] = bestAt[i];
}

/**
 * Spread whole loads over the bus sections (one per transformer) to balance loading: largest load
 * first, into the set that is emptiest relative to what it can carry, then onto that set's
 * least-loaded section. Loads with the same duty/standby group go on different sections whenever
 * the group is no larger than the number of sections. Loads are never split, so sections will not be
 * exactly equal. If the first placement leaves a section or set over its limit (and `repair` is on),
 * loads are moved and swapped to repair it (see `improveBalance`).
 *
 * With several sets, a set of 2 can carry less than a set of 3 (it is limited by the
 * one-transformer-out case), so sets are filled in proportion to what they can carry.
 */
export function distributeLoads(
  loads: SizedLoad[],
  count: number,
  ratingKva: number,
  opts: SizingOptions = DEFAULT_OPTIONS,
  layout: number[] = layoutFor(count),
  repair = true,
): { sections: Section[]; notes: string[] } {
  const sets = layout.reduce((a, k) => a + k, 0) === count ? layout : [count];
  const names: string[] = [];
  const setOf: number[] = [];
  sets.forEach((k, si) => {
    for (let j = 0; j < k; j++) {
      names.push(sets.length === 1 ? `T${j + 1}` : `S${si + 1}-T${j + 1}`);
      setOf.push(si);
    }
  });
  const setCap = sets.map((k) => setCapacityKva(k, ratingKva, opts));
  const fillCap = sets.map((k, si) => (sets.length === 1 ? 1 : setCap[si]));

  const items = [...loads].sort((a, b) => b.kva - a.kva || a.rowNumber - b.rowNumber);
  const groupSize = new Map<string, number>();
  for (const l of items) if (groupKey(l.group)) groupSize.set(groupKey(l.group), (groupSize.get(groupKey(l.group)) ?? 0) + 1);

  const buckets: number[][] = Array.from({ length: count }, () => []); // item indices
  const kvaOf = (i: number) => buckets[i].reduce((a, x) => a + items[x].kva, 0);
  const setKva = (si: number) => buckets.reduce((a, bk, i) => (setOf[i] === si ? a + bk.reduce((y, x) => y + items[x].kva, 0) : a), 0);
  // Fill the set that is emptiest relative to what it can carry, then its least-loaded section.
  const better = (i: number, j: number) => {
    const fi = setKva(setOf[i]) / fillCap[setOf[i]];
    const fj = setKva(setOf[j]) / fillCap[setOf[j]];
    if (Math.abs(fi - fj) > 1e-12) return fi < fj;
    return kvaOf(i) < kvaOf(j);
  };
  const usedByGroup = new Map<string, Set<number>>();
  const notes: string[] = [];

  items.forEach((load, idx) => {
    const key = groupKey(load.group);
    const used = key ? (usedByGroup.get(key) ?? new Set<number>()) : new Set<number>();
    let candidates = buckets.map((_, i) => i).filter((i) => !used.has(i));
    if (candidates.length === 0) {
      candidates = buckets.map((_, i) => i);
      notes.push(
        `Group "${load.group}" has more loads than the ${count} transformers, so "${load.description || `row ${load.rowNumber}`}" shares a section with another member.`,
      );
    }
    const target = candidates.reduce((best, i) => (better(i, best) ? i : best), candidates[0]);
    buckets[target].push(idx);
    if (key) usedByGroup.set(key, used.add(target));
  });

  // Repair: move / swap loads while a section or set is over its limit.
  const assign = Array<number>(items.length).fill(0);
  buckets.forEach((bk, t) => bk.forEach((i) => { assign[i] = t; }));
  if (repair) {
    const seed = count * 100003 + Math.round(ratingKva) * 31 + items.length;
    improveBalance(items, assign, setOf, opts.normalLimit * ratingKva, setCap, (k) => (groupSize.get(k) ?? 0) <= count, seed);
  }

  const placed: SizedLoad[][] = Array.from({ length: count }, () => []);
  items.forEach((l, i) => placed[assign[i]].push(l));
  const sections: Section[] = placed.map((l, i) => {
    const p = combine(l, opts.totalBasis);
    const loadingPct = (p.kva / ratingKva) * 100;
    return {
      name: names[i], loads: l, ...p, loadingPct,
      withinNormalLimit: loadingPct <= opts.normalLimit * 100 + 1e-9,
    };
  });
  return { sections, notes };
}

export interface OneOutScenario {
  /** Transformer that is out, e.g. "T2" or "S2-T1". */
  outName: string;
  /** kVA of that transformer's own bus section that has to be picked up by the others in its set. */
  movedKva: number;
  /** kVA each remaining transformer of the set carries (the set's load shared equally through its couplers). */
  eachKva: number;
  loadingPct: number;
  ok: boolean;
}

/** One independent set (switchboard) of 2 or 3 transformers. */
export interface SetResult {
  /** 1-based. */
  index: number;
  size: number;
  incomers: number;
  couplers: number;
  sections: Section[];
  /** Load assigned to this set. */
  kva: number;
  kw: number;
  kvar: number;
  /** What this set can carry within both limits. */
  capacityKva: number;
  oneOut: OneOutScenario[];
}

export interface ComboResult {
  /** Total number of transformers. */
  count: number;
  /** Sizes of the independent sets, e.g. [3, 2] for 5 transformers. */
  layout: number[];
  /** Total incomers (one per transformer). */
  incomers: number;
  /** Total bus couplers: one between neighbouring incomers within each set, none between sets. */
  couplers: number;
  /** True when this number of transformers was added because the selected number(s) could not be balanced. */
  added: boolean;
  /** One entry per rating considered, ascending, each tried with the real loads. */
  candidates: Candidate[];
  /** Smallest rating that fits; if none fits, the largest rating considered (see `fits`). */
  chosen: Candidate;
  fits: boolean;
  sets: SetResult[];
  /** All sections / one-out scenarios of all sets, in order. */
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
  /** All counts tried × ratings, in order of preference (each tried with the real loads). */
  candidates: Candidate[];
  recommended?: Candidate;
  /** One result per number of transformers tried (the ones started with, then any added). */
  combos: ComboResult[];
  /**
   * Loads too big to fit any single transformer section within the normal limit, whatever the number
   * of transformers. When there are any, no number of transformers helps and none is added.
   */
  blockers: string[];
  // Convenience for the recommended combination:
  incomers?: number;
  couplers?: number;
  sections?: Section[];
}

function buildCombo(loads: SizedLoad[], totalKva: number, count: number, opts: SizingOptions, added = false): ComboResult {
  const tried = [...opts.ratingsKva]
    .sort((a, b) => a - b)
    .map((r) => {
      const ideal = evaluateCandidate(totalKva, count, r, opts);
      // Only worth searching for a valid placement when the total fits at all if shared perfectly.
      const placement = distributeLoads(loads, count, r, opts, ideal.layout, ideal.fits);
      const balanced = isBalanced(placement.sections, ideal.layout, r, opts);
      return { cand: { ...ideal, balanced, fits: ideal.fits && balanced }, placement };
    });
  const best = tried.find((t) => t.cand.fits) ?? tried[tried.length - 1];
  const chosen = best.cand;
  const R = chosen.ratingKva;
  const { sections, notes } = best.placement;

  let at = 0;
  const sets: SetResult[] = chosen.layout.map((size, i) => {
    const own = sections.slice(at, at + size);
    at += size;
    const p = combine(own.flatMap((s) => s.loads), opts.totalBasis);
    const eachKva = p.kva / (size - 1);
    const loadingPct = (eachKva / R) * 100;
    const capacityKva = setCapacityKva(size, R, opts);
    if (!chosen.fits && p.kva > capacityKva + EPS) {
      notes.push(
        `${chosen.layout.length > 1 ? `Set ${i + 1}` : "The set"} carries ${fmtInt(p.kva)} kVA, above its limit of ${fmtInt(capacityKva)} kVA with one transformer out (loads cannot be split).`,
      );
    }
    return {
      index: i + 1,
      size,
      incomers: size,
      couplers: size - 1,
      sections: own,
      ...p,
      capacityKva,
      oneOut: own.map((s) => ({
        outName: s.name,
        movedKva: s.kva,
        eachKva,
        loadingPct,
        ok: loadingPct <= opts.contingencyLimit * 100 + EPS,
      })),
    };
  });

  if (!chosen.fits) {
    for (const s of sections.filter((x) => !x.withinNormalLimit)) {
      notes.push(
        `${s.name} carries ${fmtInt(s.kva)} kVA, above the ${fmtPct(opts.normalLimit * 100)} % normal limit of ${fmtInt(opts.normalLimit * R)} kVA (loads cannot be split).`,
      );
    }
  }
  return {
    count,
    layout: chosen.layout,
    incomers: count,
    couplers: count - chosen.layout.length,
    added,
    candidates: tried.map((t) => t.cand),
    chosen,
    fits: chosen.fits,
    sets,
    sections,
    oneOut: sets.flatMap((x) => x.oneOut),
    notes,
  };
}

/** Every section within its normal limit and every set within what it can carry with one transformer out. */
function isBalanced(sections: Section[], layout: number[], ratingKva: number, opts: SizingOptions): boolean {
  if (sections.some((s) => !s.withinNormalLimit)) return false;
  let at = 0;
  return layout.every((size) => {
    const own = sections.slice(at, at + size);
    at += size;
    return own.reduce((a, s) => a + s.kva, 0) <= setCapacityKva(size, ratingKva, opts) + EPS;
  });
}

/**
 * The combinations worth showing in full: the ones started with, plus (when numbers were added) only
 * the one that finally fits, or the last one tried if none fits. The others are summarised in a table
 * of numbers tried.
 */
export function displayCombos(result: SizingResult): ComboResult[] {
  const selected = result.combos.filter((c) => !c.added);
  const added = result.combos.filter((c) => c.added);
  if (added.length === 0) return selected;
  return [...selected, added.find((c) => c.fits) ?? added[added.length - 1]];
}

/** What a layout can carry per kVA of rating, used to skip numbers of transformers that add nothing. */
const relativeCapacity = (count: number, opts: SizingOptions) =>
  layoutFor(count).reduce((a, k) => a + setCapacityKva(k, 1, opts), 0);

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

  // A load bigger than one section can carry cannot be placed whatever the number of transformers.
  const biggest = Math.max(0, ...opts.ratingsKva);
  const blockers = loads
    .filter((l) => l.kva > opts.normalLimit * biggest + EPS)
    .map((l) => `${l.description || `Row ${l.rowNumber}`} (row ${l.rowNumber}, ${fmtInt(l.kva)} kVA) is larger than one transformer section can carry: ${fmtPct(opts.normalLimit * 100)} % of ${fmtInt(biggest)} kVA = ${fmtInt(opts.normalLimit * biggest)} kVA.`);

  const start = [...new Set(opts.counts)].filter((n) => n >= 2).sort((a, b) => a - b);
  const combos = loads.length > 0 ? start.map((n) => buildCombo(loads, totalKva, n, opts)) : [];
  // Not balanced within the limits: add transformers (sets of 3 first, a set of 2 at the end for the
  // remainder) until it works or the maximum is reached. Numbers that add no capacity (4, 7, 10 ...) are skipped.
  if (blockers.length === 0 && combos.length > 0) {
    let lastCap = relativeCapacity(Math.max(...start), opts);
    for (let n = Math.max(...start) + 1; !combos.some((c) => c.fits) && n <= opts.maxTransformers; n++) {
      const cap = relativeCapacity(n, opts);
      if (cap <= lastCap + 1e-12) continue;
      combos.push(buildCombo(loads, totalKva, n, opts, true));
      lastCap = cap;
    }
  }
  const candidates = orderCandidates(combos.flatMap((c) => c.candidates), opts);
  const recommended = loads.length > 0 ? candidates.find((c) => c.fits) : undefined;
  const result: SizingResult = {
    selectedVoltageV, loads, issues, totals, totalKva, totalKw: used.kw, candidates, recommended, combos, blockers,
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
