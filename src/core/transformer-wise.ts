/**
 * Transformer-wise view of a sizing result: one row per transformer (kW, kVA, current, % loading), the totals, and the
 * tables that show the number of transformers taken is the smallest that fits, not an excess (docs/transformer-sizing.md).
 * Used by the Excel sheet "Transformer-wise" and the page, so both show the same numbers.
 */
import { currentOf, type Candidate, type ComboResult, type SizingOptions, type SizingResult } from "./transformer-sizing";

export interface TransformerRow {
  /** "T1", "S1-T2", ... */
  name: string;
  set: number;
  loads: number;
  /** kW (input) and kvar of the loads on this transformer, kVA of the same. */
  kw: number;
  kvar: number;
  kva: number;
  /** Load current on the secondary (A): kVA x 1000 / (sqrt3 x V). */
  currentA: number;
  ratingKva: number;
  /** Rated secondary current of the transformer (A). */
  ratedCurrentA: number;
  loadingPct: number;
  limitPct: number;
  within: boolean;
  /** kVA still available up to the normal limit. */
  headroomKva: number;
}

export interface TransformerTotals {
  loads: number;
  kw: number;
  kvar: number;
  kva: number;
  currentA: number;
  installedKva: number;
  ratedCurrentA: number;
  /** Total load / installed kVA. */
  loadingPct: number;
  headroomKva: number;
}

/** One number of transformers tried, with the rating chosen for it. */
export interface QuantityRow {
  count: number;
  layout: number[];
  ratingKva: number;
  installedKva: number;
  totalKva: number;
  normalPct: number;
  oneOutPct: number;
  /** What this number can carry within both limits (kVA) and the margin over the load (negative = short). */
  capacityKva: number;
  marginKva: number;
  fits: boolean;
  /** This is the number taken. */
  final: boolean;
}

export interface TransformerWise {
  combo: ComboResult;
  rows: TransformerRow[];
  totals: TransformerTotals;
  quantity: QuantityRow[];
  /** Every count x rating tried, smallest installed kVA first. */
  candidates: (Candidate & { final: boolean })[];
  /** Plain statements that show the number taken is not an excess. */
  checks: string[];
  normalLimitPct: number;
  oneOutLimitPct: number;
}

/** The combination shown: the recommended one, else the last one tried. undefined when there are no loads. */
export function finalCombo(r: SizingResult): ComboResult | undefined {
  if (r.recommended) return r.combos.find((c) => c.count === r.recommended!.count) ?? r.combos[r.combos.length - 1];
  return r.combos[r.combos.length - 1];
}

export function transformerWise(r: SizingResult, opts: SizingOptions): TransformerWise | undefined {
  const combo = finalCombo(r);
  if (!combo) return undefined;
  const v = opts.calculationVoltageV[combo.voltageV] ?? combo.voltageV;
  const rating = combo.chosen.ratingKva;
  const limit = opts.normalLimit;
  const rows: TransformerRow[] = combo.sets.flatMap((set) =>
    set.sections.map((s): TransformerRow => ({
      name: s.name, set: set.index, loads: s.loads.length, kw: s.kw, kvar: s.kvar, kva: s.kva,
      currentA: currentOf(s.kva, v), ratingKva: rating, ratedCurrentA: currentOf(rating, v),
      loadingPct: s.loadingPct, limitPct: limit * 100, within: s.withinNormalLimit, headroomKva: limit * rating - s.kva,
    })),
  );
  const sum = (f: (x: TransformerRow) => number) => rows.reduce((t, x) => t + f(x), 0);
  const installed = rating * rows.length;
  const totals: TransformerTotals = {
    loads: sum((x) => x.loads), kw: sum((x) => x.kw), kvar: sum((x) => x.kvar), kva: sum((x) => x.kva), currentA: sum((x) => x.currentA),
    installedKva: installed, ratedCurrentA: sum((x) => x.ratedCurrentA), loadingPct: installed > 0 ? (sum((x) => x.kva) / installed) * 100 : 0,
    headroomKva: sum((x) => x.headroomKva),
  };

  const quantity: QuantityRow[] = r.combos.map((c) => ({
    count: c.count, layout: c.layout, ratingKva: c.chosen.ratingKva, installedKva: c.chosen.installedKva, totalKva: c.totalKva,
    normalPct: c.chosen.normalPct, oneOutPct: c.chosen.contingencyPct, capacityKva: c.chosen.maxLoadKva, marginKva: c.chosen.maxLoadKva - c.totalKva,
    fits: c.fits, final: c.count === combo.count,
  }));
  const candidates = [...r.candidates].sort((a, b) => a.installedKva - b.installedKva || a.count - b.count).map((c) => ({ ...c, final: c.count === combo.count && c.ratingKva === rating }));

  const checks: string[] = [];
  const fmt = (x: number) => Math.round(x).toLocaleString("en-US");
  const pctOf = (x: number) => x.toFixed(1);
  const feedersKva = combo.mccLoads.reduce((t, l) => t + l.kva, 0);
  if (combo.mccLoads.length > 0) checks.push(`The total includes the fixed PMCC/IMCC feeders: ${fmt(totals.kva)} kVA = load list ${fmt(totals.kva - feedersKva)} kVA + feeders ${fmt(feedersKva)} kVA.`);
  checks.push(`Total load ${fmt(totals.kva)} kVA on ${rows.length} × ${fmt(rating)} kVA = ${fmt(installed)} kVA installed: ${pctOf(totals.loadingPct)} % utilised (normal limit ${pctOf(limit * 100)} %).`);
  if (combo.fits) {
    const maxLoad = combo.chosen.maxLoadKva;
    checks.push(`It can carry ${fmt(maxLoad)} kVA within both limits (${pctOf(limit * 100)} % normal, ${pctOf(opts.contingencyLimit * 100)} % with one out): ${fmt(maxLoad - totals.kva)} kVA of margin, ${pctOf(((maxLoad - totals.kva) / (maxLoad / rows.length)) * 100)} % of what one transformer adds.`);
    const smaller = r.combos.filter((c) => c.count < combo.count);
    const prev = smaller[smaller.length - 1];
    if (prev) {
      checks.push(`The next smaller number tried, ${prev.count} × ${fmt(prev.chosen.ratingKva)} kVA, carries only ${fmt(prev.chosen.maxLoadKva)} kVA within the limits against a load of ${fmt(prev.totalKva)} kVA: short by ${fmt(prev.totalKva - prev.chosen.maxLoadKva)} kVA${prev.chosen.balanced ? "" : " (and the loads cannot be balanced)"}. ${opts.pairsOnly ? "" : "Numbers that add no capacity (4, 7, 10 ...) are not tried."}`);
    } else {
      checks.push(`No smaller number is tried: ${combo.count} is the number started with and it fits.`);
    }
    const smallerRatings = combo.candidates.filter((c) => c.ratingKva < rating);
    if (smallerRatings.length) checks.push(`A smaller rating at this number does not fit: ${smallerRatings.map((c) => `${fmt(c.ratingKva)} kVA (one out ${pctOf(c.contingencyPct)} %, normal ${pctOf(c.normalPct)} %)`).join("; ")}.`);
    const fitting = candidates.filter((c) => c.fits);
    const cheapest = fitting[0];
    if (cheapest) {
      checks.push(cheapest.final
        ? `Of all the combinations tried that fit, this one has the smallest installed capacity (${fmt(cheapest.installedKva)} kVA).`
        : `Another combination that fits has less installed capacity: ${cheapest.count} × ${fmt(cheapest.ratingKva)} kVA = ${fmt(cheapest.installedKva)} kVA (the preference setting orders by count first).`);
    }
    if (opts.pairsOnly) {
      for (const set of combo.sets) {
        for (const o of set.oneOut) {
          checks.push(`Pair ${set.index}: if ${o.outName} is out, the other transformer takes the entire bus load of the pair, ${fmt(o.eachKva)} kVA = ${pctOf(o.loadingPct)} % of ${fmt(rating)} kVA (limit ${pctOf(opts.contingencyLimit * 100)} %).`);
        }
      }
    }
  } else {
    checks.push("No configuration fits within both limits: see the tables above.");
  }
  return { combo, rows, totals, quantity, candidates, checks, normalLimitPct: limit * 100, oneOutLimitPct: opts.contingencyLimit * 100 };
}
