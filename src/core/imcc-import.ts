/**
 * Reads an IMCC / MCC load-list sheet in the layout of the reference workbook (sheets "ECR-9 IMCC", "ECR-10 MCC"):
 * a header row with S. NO., KW, W/S/Spare, Method of starting, Group ...; the description in the column after S. NO.;
 * and rows "INCOMER - 1" / "INCOMER - 2" that start each bus section. Nothing is hardcoded to those two sheets.
 *
 * The reference file's own "Current in Amps" column is a formula of kW (415 V, 1.73): it is NOT imported. The
 * application recalculates current with its own conventions (docs/imcc.md).
 */
import { newId, normalizeMethod, normalizeStatus, type ImccLoad, type BusSection } from "./project";

type Cell = unknown;
const norm = (h: Cell) => String(h ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const blank = (c: Cell) => c === null || c === undefined || (typeof c === "string" && c.trim() === "");
const num = (c: Cell): number | undefined => {
  if (typeof c === "number" && Number.isFinite(c)) return c;
  if (typeof c === "string" && c.trim() !== "" && Number.isFinite(Number(c))) return Number(c);
  return undefined;
};

export interface ImccImport { loads: ImccLoad[]; errors: string[]; warnings: string[]; incomerRows: number }

export function parseImccSheet(data: readonly (readonly Cell[])[]): ImccImport {
  const errors: string[] = [];
  const warnings: string[] = [];
  const headerIdx = data.findIndex((r) => r.some((c) => /^s\.?no\.?$/.test(norm(c)) || norm(c) === "sno"));
  if (headerIdx < 0) return { loads: [], errors: ["No header row with \"S. NO.\" was found."], warnings, incomerRows: 0 };
  const header = data[headerIdx];
  const find = (test: (n: string) => boolean) => header.findIndex((h) => test(norm(h)));
  const sNo = find((n) => n === "sno");
  const kw = find((n) => n === "kw");
  const status = find((n) => n.startsWith("wsspare") || n === "ws" || n.startsWith("workingstandby"));
  const method = find((n) => n.startsWith("methodofstarting"));
  const group = find((n) => n === "group");
  const feederDesc = find((n) => n.startsWith("feederdisc"));
  const desc = sNo + 1;
  const missing = [["KW", kw], ["W/S/Spare", status], ["Method of starting", method]].filter(([, i]) => (i as number) < 0).map(([n]) => n);
  if (missing.length) errors.push(`Missing column(s): ${missing.join(", ")}.`);
  if (errors.length) return { loads: [], errors, warnings, incomerRows: 0 };
  const lastLabelled = Math.max(...header.map((h, i) => (blank(h) ? -1 : i)));

  const loads: ImccLoad[] = [];
  let section: BusSection | undefined;
  let incomerRows = 0;
  for (let r = headerIdx + 1; r < data.length; r++) {
    const row = data[r];
    const marker = row.map((c) => /^\s*incomer\s*-?\s*([12])\s*$/i.exec(String(c ?? ""))).find(Boolean);
    if (marker) { section = Number(marker[1]) as BusSection; incomerRows++; continue; }
    const description = String(row[desc] ?? "").trim();
    if (!description) continue; // blank line or a totals row
    const noteParts: string[] = [];
    for (const i of [desc + 1, feederDesc]) {
      if (i > 0 && i !== kw && i !== status && i !== method && i !== group && typeof row[i] === "string" && String(row[i]).trim()) noteParts.push(String(row[i]).trim());
    }
    for (let i = lastLabelled + 1; i < row.length; i++) if (typeof row[i] === "string" && String(row[i]).trim()) noteParts.push(String(row[i]).trim());
    const rawStatus = row[status];
    loads.push({
      id: newId("load"), section, sNo: blank(row[sNo]) ? undefined : String(row[sNo]), description,
      kw: num(row[kw]), status: normalizeStatus(rawStatus), statusText: blank(rawStatus) ? undefined : String(rawStatus).trim(),
      method: normalizeMethod(row[method]), group: group >= 0 && !blank(row[group]) ? String(row[group]).trim() : undefined,
      note: noteParts.length ? noteParts.join("; ") : undefined,
    });
  }
  if (loads.length === 0) errors.push("No load rows were found under the header.");
  if (incomerRows === 0) warnings.push("No \"INCOMER - 1 / INCOMER - 2\" section rows found: loads have no bus section and are not counted until one is set.");
  return { loads, errors, warnings, incomerRows };
}
