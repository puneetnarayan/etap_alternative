/**
 * Unique BOM items from a load list: one item per type + rating (see docs/bom.md). Each item is matched to a
 * unit-BOM sheet of the template (public/bom/unit_bom_template.xlsx). Where no rule exists yet the item gets an
 * action (shown in red) instead of a guess. Preliminary, like the breaker selection it builds on.
 */
import { ctPrimaryFor, type MccRow } from "./mcc-breaker";

/** What the item is, in BOM terms. */
export type BomType =
  | "Incomer" | "Feeder" | "DOL" | "DOL + soft starter" | "VFD" | "RDOL";

export interface BomItem {
  /** Unique key: type + breaker + motor kW. */
  key: string;
  type: BomType;
  /** Breaker device and rating, e.g. "400 A MCCB"; "" when none. */
  breaker: string;
  /** Breaker / incomer rating (A). */
  ratingA?: number;
  kw?: number;
  /** Current behind the item (from its first load), to explain the breaker selection: used (entered or calculated), calculated from kW, design (x multiplier). */
  usedA?: number;
  calculatedA?: number;
  designA?: number;
  /** Number of loads in the list with this type + rating. */
  count: number;
  /** Rows of the load list (sheet row numbers). */
  rows: number[];
  /** Template sheet(s) this item is built from; empty when no sheet is needed or none exists. */
  sheets: string[];
  /** A sheet is required but none can be chosen yet, or a rule is to be reviewed: shown in red. */
  action?: string;
  /** No BOM sheet is required for this item (e.g. a plain 6 or 10 A MPCB / MCB feeder). */
  noSheet?: boolean;
}

const ACB_OG: Record<number, string> = { 630: "OG 630A ACB Feeder", 1000: "OG 1000A ACB", 1250: "OG 1250A ACB" };
const INCOMER: Record<number, string> = {
  1000: "Incomer 1000A IMCC ", 1250: "Incomer 1250A IMCC", 1600: "Incomer 1600A", 2000: "Incomer 2000A",
  2500: "Incomer 2500A", 3200: "Incomer 3200A", 4000: "Incomer 4000A",
};
export const SOFT_STARTER_ACB_SHEETS = ["OG 315KW Soft starter ", "OG 225KW Soft starter"];
const SOFT_STARTER_MCCB_SHEET = "OG 75KW Soft starter ";
/** Above this motor size (kW) a DOL motor gets a soft starter (BOM index note). */
export const SOFT_STARTER_ABOVE_KW = 200;
/** Plain MPCB / MCB feeders (at or below this rating) need no BOM sheet. */
const PLAIN_SMALL_A = 10;

function classify(r: MccRow): Omit<BomItem, "count" | "rows" | "key"> | null {
  const b = r.isIncomer ? r.incomer ?? r.breaker : r.breaker;
  if (r.current.designA === undefined && !r.isIncomer) return null; // no current: nothing to select
  const rating = b.ratingA;
  const label = b.kind === "REVIEW" || rating === undefined ? "" : `${rating} A ${b.kind}`;
  const none = (action: string) => ({ type: "Feeder" as BomType, breaker: label, ratingA: rating, sheets: [], action });

  if (r.isIncomer) {
    const sheet = rating === undefined ? undefined : INCOMER[rating];
    return {
      type: "Incomer", breaker: label, ratingA: rating, sheets: sheet ? [sheet] : [],
      action: sheet ? undefined : `No incomer sheet for ${label || "this rating"}: engineering review or a new sheet is needed`,
    };
  }
  if (b.kind === "REVIEW" || rating === undefined) {
    return { type: motorType(r), breaker: "", kw: r.kw, sheets: [], action: "Breaker not selected (engineering review): no BOM sheet can be chosen" };
  }

  // RDOL (small reversing motors, e.g. valves): MPCB feeder sheet
  if (/rdol/i.test(r.rawType)) {
    return { type: "RDOL", breaker: label, ratingA: rating, kw: r.kw, sheets: ["RDOL < 5KW"] };
  }
  const motor = r.type === "DOL" || r.type === "VFD";
  const kw = r.kw;
  if (r.type === "DOL" && kw !== undefined && kw > SOFT_STARTER_ABOVE_KW) {
    if (b.kind === "EDO ACB") {
      return {
        type: "DOL + soft starter", breaker: label, ratingA: rating, kw, sheets: [...SOFT_STARTER_ACB_SHEETS],
        action: "Rule needed: 225 kW or 315 kW soft-starter feeder? Both sheets are kept until reviewed",
      };
    }
    return { type: "DOL + soft starter", breaker: label, ratingA: rating, kw, sheets: [SOFT_STARTER_MCCB_SHEET] };
  }
  if (b.kind === "EDO ACB") {
    const sheet = ACB_OG[rating];
    if (motor) return { ...none(`No sheet for a ${r.type} motor feeder on a ${label}`), type: motorType(r), kw };
    return { type: "Feeder", breaker: label, ratingA: rating, sheets: sheet ? [sheet] : [], action: sheet ? undefined : `No outgoing ACB sheet above 1250 A (${label}): new sheet needed` };
  }
  // MCCB / MPCB / MCB
  if (r.type === "VFD") return { type: "VFD", breaker: label, ratingA: rating, kw, sheets: ["55kW VFD"] };
  if (r.type === "DOL") {
    return {
      type: "DOL", breaker: label, ratingA: rating, kw, sheets: ["DOL < 75KW"],
      action: kw !== undefined && kw >= 75 ? "DOL 75 kW and above uses the '< 75 KW' sheet: confirm" : undefined,
    };
  }
  if (rating <= PLAIN_SMALL_A) return { type: "Feeder", breaker: label, ratingA: rating, sheets: [], noSheet: true };
  return {
    type: "Feeder", breaker: label, ratingA: rating,
    sheets: [rating >= 250 ? "OG 250 to 630A drawout Fdr MCCB" : "OG <250A  power FDR - MCCB"],
  };
}

const motorType = (r: MccRow): BomType => (r.type === "VFD" ? "VFD" : r.type === "DOL" ? "DOL" : "Feeder");

/** One item per unique type + rating (breaker rating and, for motors, kW), in a stable order. */
export function bomItems(rows: MccRow[]): BomItem[] {
  const map = new Map<string, BomItem>();
  for (const r of rows) {
    const c = classify(r);
    if (!c) continue;
    const key = [c.type, c.breaker || "review", c.kw ?? ""].join("|");
    const have = map.get(key);
    if (have) {
      have.count++;
      have.rows.push(r.rowNumber);
    } else {
      map.set(key, { ...c, key, count: 1, rows: [r.rowNumber], usedA: r.current.usedA, calculatedA: r.current.calculatedA, designA: r.current.designA });
    }
  }
  const order: BomType[] = ["Incomer", "Feeder", "DOL", "DOL + soft starter", "VFD", "RDOL"];
  return [...map.values()].sort(
    (a, b) => order.indexOf(a.type) - order.indexOf(b.type) || (b.ratingA ?? 0) - (a.ratingA ?? 0) || (b.kw ?? 0) - (a.kw ?? 0),
  );
}

/** Text replacements that set a cloned unit sheet to the item's rating (the sheet layout is never changed). */
export function ratingReplacements(item: BomItem, sheet: string): [RegExp, string][] {
  const a = item.ratingA;
  if (a === undefined) return [];
  const ct = ctPrimaryFor(a);
  const out: [RegExp, string][] = [];
  const ctRule: [RegExp, string][] = ct ? [[/CT CTR:\s*xxx\s*\/\s*1\s*A/i, `CT CTR: ${ct}/1 A`]] : [];
  if (sheet.startsWith("OG 250 to 630A")) {
    out.push([/^250A 4P/, `${a}A 4P`], [/^OG 250A to 630 A drawout Feeder$/, `OG ${a}A drawout Feeder`], [/ELR with CBCT for 250 A/, `ELR with CBCT for ${a} A`], ...ctRule);
  } else if (sheet.startsWith("OG <250A")) {
    out.push([/^< 250A 4P/, `${a}A 4P`], [/^OG <250A\s+power Feeder$/, `OG ${a}A power Feeder`], [/ELR with CBCT for 250 A/, `ELR with CBCT for ${a} A`], ...ctRule);
  } else if (sheet === "DOL < 75KW") {
    out.push([/^MCCB 250A 3P/, `MCCB ${a}A 3P`]);
    if (item.kw !== undefined) out.push([/^DOL < 75KW$/, `DOL ${item.kw} KW`]);
  } else if (sheet === "55kW VFD") {
    out.push([/^< 200A 4P/, `${a}A 4P`], [/^OG <200A\s+power Feeder$/, `OG ${a}A VFD feeder${item.kw !== undefined ? ` (${item.kw} kW)` : ""}`]);
  }
  return out;
}
