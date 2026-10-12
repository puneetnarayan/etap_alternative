/**
 * Number display rules shared by the page and the Excel report:
 *  - only percentages have a decimal place (one);
 *  - every other number is whole, with thousands separators;
 *  - power factor is a ratio below 1 and keeps two decimals.
 * These only affect display; calculations use the full values.
 */
const whole = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const oneDecimal = new Intl.NumberFormat("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });

/** 3006.2 -> "3,006". */
export const fmtInt = (v: number): string => {
  const t = whole.format(v);
  return t === "-0" ? "0" : t; // a tiny negative number must not show as "-0"
};
/** 45.234 -> "45.2" (no % sign; add " %" where wanted). */
export const fmtPct = (v: number): string => oneDecimal.format(v);
/** 0.7 -> "0.70". */
export const fmtPf = (v: number): string => v.toFixed(2);
