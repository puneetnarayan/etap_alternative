"use client";

import { useEffect, useState } from "react";
import readExcelFile from "read-excel-file/universal";
import { REPORT_COLORS, kindColor } from "@/core";

const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = { padding: "2px 6px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const, textAlign: "center" as const };
const td = { padding: "2px 6px", border: BORDER, textAlign: "left" as const, verticalAlign: "top" as const };

type Kind = "PMCC" | "IMCC";

const COLUMNS: Record<Kind, { name: string; need: string; text: string }[]> = {
  PMCC: [
    { name: "Description", need: "Required", text: "Name of the load or feeder." },
    { name: "W/S", need: "Optional", text: "W or w = working, S = standby. Blank: the load is counted without a W/S tag. All loads count at 100 % for the transformer sizing, standby included. (\"Spare\" is not recognised in the PMCC list.)" },
    { name: "Type (RDOL/DOL/VFD/SS/Feeder)", need: "Required", text: "DOL, RDOL, SS (soft starter, calculated like DOL), VFD or Feeder. kW motors use pf 0.7 (DOL / RDOL / SS), 0.9 (VFD) and the motor efficiency; amps-only rows are feeders at pf 0.7." },
    { name: "Amps rating (A)", need: "Amps or kW", text: "Entered current. When given it is used in preference to a current calculated from kW." },
    { name: "kW rating (kW)", need: "Amps or kW", text: "Motor or load power. The current is calculated from it when no amps are given (the MCC and BOM tabs show it in a \"Calculated current\" column)." },
    { name: "Voltage level (V)", need: "Required", text: "433 or 690. Rows that do not match the selected transformer voltage are listed as excluded." },
    { name: "Duty/standby group", need: "Optional", text: "Loads with the same group name are placed on different transformer sections." },
  ],
  IMCC: [
    { name: "Description", need: "Required", text: "Name of the load or feeder." },
    { name: "W/S", need: "Required", text: "W or w = working (the only loads used for the design), S = standby, Spare = full spare feeder (shown as SP). A blank or other value is flagged and the row is not counted or placed." },
    { name: "Type (RDOL/DOL/VFD/SS/Feeder)", need: "Required", text: "DOL, RDOL, SS (soft starter, calculated like DOL), VFD or Feeder (\"PWR FDR\" is read as Feeder)." },
    { name: "Amps rating (A)", need: "Amps or kW", text: "Entered current. When given it is used in preference to the current calculated from kW; the IMCC panel shows entered, calculated and used current side by side." },
    { name: "kW rating (kW)", need: "Amps or kW", text: "Motor or load power. A number that cannot be read (e.g. \"-\") is flagged." },
    { name: "Voltage level (V)", need: "Required", text: "433 (the IMCC voltage). A row at another voltage is not placed." },
    { name: "Duty/standby group", need: "Optional", text: "Standby and spare rows stay with the IMCC of their group; a group with standby but no working load is flagged." },
  ],
};

const NOTES: Record<Kind, string[]> = {
  PMCC: [
    "File name must start with PMCC (any case), e.g. PMCC_sample_load_list.xlsx. It goes to the PMCC design: transformer sizing, MCC Load List and BOM.",
    "First sheet only; keep the header row. Columns are found by their header words (description, type, amps, kW, voltage, group, working / standby), so the order does not matter and extra columns are ignored.",
    "Excel only: .xlsx.",
  ],
  IMCC: [
    "File name must start with IMCC (any case), e.g. IMCC_sample_load_list.xlsx. It goes to the IMCC design. If a PMCC and an IMCC file are dropped together, a pop-up asks for confirmation and each is processed by its own logic.",
    "Same layout as the PMCC list. First sheet only; keep the header row; extra columns are ignored.",
    "Only working loads drive the IMCC design; the number of IMCCs follows from them (incomer ACB and its maximum loading are set on the IMCC page).",
  ],
};

export default function FormatPage({ kind }: { kind: Kind }) {
  const file = `/${kind}_sample_load_list.xlsx`;
  const [rows, setRows] = useState<unknown[][] | null>(null);
  const [error, setError] = useState("");
  useEffect(() => {
    let live = true;
    (async () => {
      try {
        const res = await fetch(`${window.location.origin}${file}`);
        if (!res.ok) throw new Error(`could not load ${file} (${res.status})`);
        const sheets = await readExcelFile(await res.arrayBuffer());
        if (live) setRows(sheets[0].data as unknown[][]);
      } catch (e) { if (live) setError((e as Error).message); }
    })();
    return () => { live = false; };
  }, [file]);

  const header = rows?.[0] ?? [];
  const shown = rows?.slice(1, 16) ?? [];
  const cellStyle = (h: string, v: unknown) => {
    const t = String(h ?? "").toLowerCase();
    const s = String(v ?? "");
    const bg = t.includes("w/s") ? (s.toLowerCase() === "spare" ? REPORT_COLORS.spare : kindColor(s)) : t.startsWith("type") ? kindColor(/vfd/i.test(s) ? "VFD" : /dol|^ss$/i.test(s) ? "DOL" : s) : null;
    return { ...td, textAlign: "center" as const, ...(bg ? { background: bg, color: "#000" } : {}) };
  };

  return (
    <div>
      <p style={{ background: kind === "IMCC" ? "#E3F4F1" : "#E4EEF9", padding: "4px 10px", borderRadius: 6 }}>
        The {kind} load list is an Excel file whose name starts with <b>{kind}</b>.{" "}
        <a href={file} download>Download the sample {kind} load list (.xlsx)</a>
      </p>
      <h3>Columns</h3>
      <table style={{ borderCollapse: "collapse" }}>
        <thead><tr><th style={th}>Column header</th><th style={th}>Needed</th><th style={th}>What goes in it</th></tr></thead>
        <tbody>
          {COLUMNS[kind].map((c) => (
            <tr key={c.name}><td style={{ ...td, fontWeight: "bold" }}>{c.name}</td><td style={td}>{c.need}</td><td style={{ ...td, maxWidth: 620 }}>{c.text}</td></tr>
          ))}
        </tbody>
      </table>
      <h3>Rules</h3>
      <ul style={{ marginTop: 2 }}>{NOTES[kind].map((n) => <li key={n}>{n}</li>)}</ul>
      <h3>The sample file ({rows ? `${rows.length - 1} rows, first ${shown.length} shown` : "loading…"})</h3>
      {error && <p role="alert" style={{ color: "crimson" }}>{error}</p>}
      {rows && (
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse" }}>
            <thead><tr>{header.map((h, i) => <th key={i} style={th}>{String(h ?? "")}</th>)}</tr></thead>
            <tbody>
              {shown.map((r, n) => (
                <tr key={n}>{header.map((h, i) => <td key={i} style={{ ...cellStyle(String(h ?? ""), r[i]), ...(i === 0 ? { textAlign: "left" as const } : {}) }}>{r[i] === null || r[i] === undefined ? "" : typeof r[i] === "number" ? (Number.isInteger(r[i]) ? String(r[i]) : (r[i] as number).toFixed(1)) : String(r[i])}</td>)}</tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
