import { parseLoadSheet } from "./transformer-sizing";

/**
 * Preliminary breaker selection for the MCC load list (see docs/mcc-load-list.md).
 * Pure functions: the Excel workbook uses the same rules as formulas, and the tests compare the two.
 * Everything here is a preliminary aid and needs engineering approval.
 */

/** The only permitted MCCB ratings (A). A 630 A MCCB is never selected. */
export const MCCB_RATINGS_A = [32, 63, 100, 160, 250, 400] as const;
/** Standard ACB ratings (A). */
export const ACB_RATINGS_A = [630, 800, 1000, 1250] as const;
/** From this requirement (A) upward a 630 A EDO ACB is the starting standard instead of an MCCB. */
export const ACB_THRESHOLD_A = 500;
/** Standard CT primary currents (A), for indicative incomer CT ratios (secondary 5 A). */
export const CT_PRIMARIES_A = [100, 150, 200, 250, 300, 400, 500, 600, 750, 800, 1000, 1250, 1500, 2000] as const;

export type BreakerKind = "MCCB" | "EDO ACB" | "REVIEW";

export interface BreakerSelection {
  kind: BreakerKind;
  /** Undefined when engineering review is required (no rating is chosen). */
  ratingA?: number;
  /** "ok" = preliminary selection; "verify" = selected but must be verified; "review" = no selection made. */
  status: "ok" | "verify" | "review";
  message: string;
}

export interface SelectOptions {
  /** The source explicitly identifies an 800 A ACB as the incomer: standardise to 1000 A EDO ACB. */
  existing800AcbIncomer?: boolean;
  acbThresholdA?: number;
}

/** Text shown for a selection, e.g. "63 A MCCB", "630 A EDO ACB" or "Engineering review required". */
export function breakerLabel(s: BreakerSelection): string {
  return s.kind === "REVIEW" || s.ratingA === undefined ? "Engineering review required" : `${s.ratingA} A ${s.kind}`;
}

/** Preliminary breaker for a design current (A). No design current -> review (never a made-up zero). */
export function selectBreaker(designA: number | undefined | null, opts: SelectOptions = {}): BreakerSelection {
  const threshold = opts.acbThresholdA ?? ACB_THRESHOLD_A;
  if (opts.existing800AcbIncomer) {
    return {
      kind: "EDO ACB", ratingA: 1000, status: "verify",
      message: "Existing 800 A ACB incomer standardised to 1000 A EDO ACB; re-check CT ratio, busbar rating and fault level.",
    };
  }
  if (designA === undefined || designA === null || !Number.isFinite(designA) || designA <= 0) {
    return { kind: "REVIEW", status: "review", message: "No design current available, so no breaker is selected." };
  }
  if (designA < threshold) {
    const mccb = MCCB_RATINGS_A.find((r) => designA <= r);
    if (mccb !== undefined) return { kind: "MCCB", ratingA: mccb, status: "ok", message: "Next permitted MCCB rating (preliminary)." };
    return {
      kind: "REVIEW", status: "review",
      message: `Design current ${designA.toFixed(0)} A is above the largest permitted MCCB (400 A) but below the ${threshold} A threshold for a 630 A EDO ACB: engineering review required.`,
    };
  }
  if (designA <= 630) return { kind: "EDO ACB", ratingA: 630, status: "ok", message: "630 A EDO ACB (preliminary; a 630 A MCCB is never used)." };
  const acb = ACB_RATINGS_A.find((r) => designA <= r);
  if (acb !== undefined) {
    return { kind: "EDO ACB", ratingA: acb, status: "verify", message: `Above 630 A: ${acb} A EDO ACB, engineering verification required.` };
  }
  return { kind: "REVIEW", status: "review", message: "Design current is above the largest standard ACB (1250 A): engineering review required." };
}

/** Indicative CT primary (A) for a breaker rating: the smallest standard primary not below it. */
export function ctPrimaryFor(ratingA: number | undefined): number | undefined {
  if (ratingA === undefined) return undefined;
  return CT_PRIMARIES_A.find((p) => p >= ratingA);
}

export type MccLoadType = "DOL" | "VFD" | "FEEDER";

export interface MccAssumptions {
  pfGeneral: number;
  pfDol: number;
  pfVfd: number;
  /** Provisional motor efficiency, applied to DOL motors given in kW only. */
  efficiency: number;
  designMultiplier: number;
  /** Voltage used in the current formula for the 433 V level (V). */
  calcVoltage433: number;
}

export const DEFAULT_MCC_ASSUMPTIONS: MccAssumptions = {
  pfGeneral: 0.7, pfDol: 0.7, pfVfd: 0.9, efficiency: 0.93, designMultiplier: 1.25, calcVoltage433: 433,
};

export interface CurrentInput {
  type: MccLoadType;
  kw?: number;
  voltageV?: number;
  /** Rated / entered current (A) from the load list. */
  ratedA?: number;
}

export interface CurrentResult {
  /** Calculated from kW (blank when kW or voltage is missing). */
  calculatedA?: number;
  /** Rated current if given, else the calculated one. */
  usedA?: number;
  source: "Entered current" | "Calculated" | "Missing input";
  designA?: number;
  pf: number;
}

const pos = (n: number | undefined): n is number => n !== undefined && Number.isFinite(n) && n > 0;

/** Same rules as the workbook formulas: I = kW x 1000 / (sqrt(3) x V x PF [x efficiency for DOL]). */
export function currentFor(i: CurrentInput, a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS): CurrentResult {
  const pf = i.type === "VFD" ? a.pfVfd : i.type === "DOL" ? a.pfDol : a.pfGeneral;
  const v = pos(i.voltageV) ? (Math.round(i.voltageV) === 433 ? a.calcVoltage433 : i.voltageV) : undefined;
  let calculatedA: number | undefined;
  if (pos(i.kw) && v !== undefined) {
    calculatedA = (i.kw * 1000) / (Math.sqrt(3) * v * pf * (i.type === "DOL" ? a.efficiency : 1));
  }
  const usedA = pos(i.ratedA) ? i.ratedA : calculatedA;
  return {
    calculatedA, usedA, pf,
    source: pos(i.ratedA) ? "Entered current" : calculatedA !== undefined ? "Calculated" : "Missing input",
    designA: usedA === undefined ? undefined : usedA * a.designMultiplier,
  };
}

/* ------------------------- rows for the page (same rules as the workbook) ------------------------- */


export interface MccRow {
  rowNumber: number;
  description: string;
  status?: "W" | "S";
  type: MccLoadType;
  voltageV?: number;
  kw?: number;
  ratedA?: number;
  current: CurrentResult;
  /** Feeder breaker (preliminary). */
  breaker: BreakerSelection;
  /** Overload relay setting (A): DOL motors only. */
  relayA?: number;
  isIncomer: boolean;
  incomer?: BreakerSelection;
  ctPrimaryA?: number;
  /** Short flags for the engineering-checks column. */
  flags: string[];
}

export interface MccSummary {
  loads: number;
  mccb: number;
  acb: number;
  review: number;
  missing: number;
  verify: number;
}

const typeOf = (raw: string): MccLoadType => (/vfd/i.test(raw) ? "VFD" : /dol/i.test(raw) ? "DOL" : "FEEDER");

/** Read a load sheet (first row = headers) and work out current and preliminary breaker for each load. */
export function mccRows(
  sheet: readonly (readonly unknown[])[],
  a: MccAssumptions = DEFAULT_MCC_ASSUMPTIONS,
): { rows: MccRow[]; summary: MccSummary; errors: string[] } {
  const parsed = parseLoadSheet(sheet.map((r) => [...r]));
  if (parsed.errors.length) return { rows: [], summary: { loads: 0, mccb: 0, acb: 0, review: 0, missing: 0, verify: 0 }, errors: parsed.errors };
  const rows = parsed.rows.map((r): MccRow => {
    const type = typeOf(r.type);
    const current = currentFor({ type, kw: r.kw, voltageV: r.voltageV, ratedA: r.amps }, a);
    const breaker = selectBreaker(current.designA);
    const isIncomer = /incomer/i.test(r.description);
    const incomer = isIncomer ? (r.amps === 800 ? selectBreaker(current.designA, { existing800AcbIncomer: true }) : breaker) : undefined;
    const flags: string[] = [];
    if (current.usedA === undefined) flags.push("Missing input: give a rated current, or kW and voltage");
    if (current.source === "Calculated" && type === "DOL") flags.push("Provisional motor efficiency");
    if (type === "VFD") flags.push("Verify VFD input current with the manufacturer");
    if (breaker.status === "review" && current.designA !== undefined) flags.push(breaker.message);
    if (breaker.status === "verify") flags.push("Above 630 A: verify ACB rating");
    if (incomer && r.amps === 800) flags.push("800 A ACB incomer standardised to 1000 A EDO ACB");
    return {
      rowNumber: r.rowNumber, description: r.description, status: r.status, type, voltageV: r.voltageV, kw: r.kw, ratedA: r.amps,
      current, breaker,
      relayA: type === "DOL" && current.usedA !== undefined ? Math.round(current.usedA * 10) / 10 : undefined,
      isIncomer, incomer, ctPrimaryA: incomer ? ctPrimaryFor(incomer.ratingA) : undefined, flags,
    };
  });
  const summary: MccSummary = {
    loads: rows.length,
    mccb: rows.filter((r) => r.breaker.kind === "MCCB").length,
    acb: rows.filter((r) => r.breaker.kind === "EDO ACB").length,
    review: rows.filter((r) => r.current.designA !== undefined && r.breaker.kind === "REVIEW").length,
    missing: rows.filter((r) => r.current.usedA === undefined).length,
    verify: rows.filter((r) => r.breaker.status === "verify").length,
  };
  return { rows, summary, errors: [] };
}
