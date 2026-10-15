/**
 * Reads an IMCC load list. The input is the same layout as the PMCC load list (docs/imcc.md): header row with
 * Description, W/S, Type (RDOL/DOL/VFD/SS/Feeder), Amps rating (A), kW rating (kW), Voltage level (V) and an optional
 * Duty/standby group; first sheet only. Column reading and number parsing are the PMCC reader's (`parseLoadSheet`).
 * W and w are working, S standby, Spare (or SP) a spare feeder; a blank or other W/S value is kept and flagged.
 */
import { newId, normalizeMethod, normalizeStatus, type ImccLoad } from "./imcc";
import { parseLoadSheet } from "./transformer-sizing";

export interface ImccImport { loads: ImccLoad[]; errors: string[]; warnings: string[] }

export function parseImccLoadList(sheet: readonly (readonly unknown[])[]): ImccImport {
  const parsed = parseLoadSheet(sheet as never);
  if (parsed.errors.length) return { loads: [], errors: parsed.errors, warnings: [] };
  const warnings: string[] = [];
  const loads: ImccLoad[] = parsed.rows.map((r) => {
    // parseLoadSheet reads W and S; "Spare" arrives as unrecognised text and is classified here.
    const status = r.status ?? normalizeStatus(r.statusText);
    return {
      id: newId("load"), srcRow: r.rowNumber, description: r.description,
      kw: r.kw, ratedA: r.amps, voltageV: r.voltageV,
      status, statusText: r.status ? r.status : r.statusText,
      method: normalizeMethod(r.type), group: r.group, badFields: r.badFields,
      note: r.type && !normalizeMethod(r.type) ? `Type "${r.type}" not recognised` : undefined,
    };
  });
  if (loads.length === 0) return { loads, errors: ["No load rows were found under the header."], warnings };
  return { loads, errors: [], warnings };
}
