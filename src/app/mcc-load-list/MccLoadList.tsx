"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { readSheet } from "read-excel-file/universal";
import { DEFAULT_MCC_ASSUMPTIONS, REPORT_COLORS, breakerLabel, fmtInt, fmtPf, kindColor, mccRows, type MccRow } from "@/core";
import { loadLoadList, type StoredLoadList } from "../loadListStore";
import { useShell } from "../AppShell";

export default function MccLoadList() {
  const { status, generate } = useShell();
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const input = useRef<HTMLInputElement>(null);
  // Read after mount (and after each generate): the server render has no browser storage.
  const [stored, setStored] = useState<StoredLoadList | null>(null);
  useEffect(() => { setStored(loadLoadList()); }, [status]);

  const result = useMemo(() => (stored ? mccRows(stored.sheet) : null), [stored]);

  async function handleFile(file: File | undefined) {
    if (!file) return;
    setError("");
    if (!/\.xlsx$/i.test(file.name)) {
      setError("Please use an Excel .xlsx file.");
      return;
    }
    try {
      const sheet = await readSheet(await file.arrayBuffer());
      await generate({ fileName: file.name, sheet });
    } catch (e) {
      setError(`Could not read that file: ${(e as Error).message}`);
    }
  }

  return (
    <div>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); void handleFile(e.dataTransfer.files[0]); }}
        onClick={() => input.current?.click()}
        style={{ border: "2px dashed #888", borderRadius: 8, padding: "28px 16px", textAlign: "center", cursor: "pointer", background: dragging ? "#C6EFCE" : "transparent" }}
      >
        Drag and drop a load list (.xlsx) here, or click to choose a file. The workbook is created and downloaded straight away.
        <input ref={input} type="file" accept=".xlsx" hidden onChange={(e) => { void handleFile(e.target.files?.[0]); e.target.value = ""; }} />
      </div>
      {error && <p role="alert" style={{ color: "#c0392b" }}>{error}</p>}
      <p>
        {stored ? <>Current load list: <b>{stored.fileName}</b> (the last one used in this tab). </> : <>No load list has been used in this tab yet. </>}
        <button type="button" disabled={!stored || status.kind === "busy"} onClick={() => void generate()}>
          Generate the workbook again
        </button>
      </p>
      <div aria-live="polite">
        {status.kind === "busy" && <p>Generating the workbook…</p>}
        {status.kind === "done" && (
          <p>
            ✔ Downloaded <b>{status.name}</b> ({status.rows} loads{status.missing > 0 ? `, ${status.missing} without enough data for a current` : ""}).
            A browser cannot launch Excel by itself: open the downloaded file from your downloads list.
          </p>
        )}
        {status.kind === "no-list" && <p>No load list yet: drop one above, or drop one on the Transformer sizing page first.</p>}
        {status.kind === "error" && <p role="alert" style={{ color: "#c0392b" }}>Could not generate the workbook: {status.message}</p>}
      </div>
      {result && <Results result={result} />}
    </div>
  );
}

// Same look as the Transformer sizing page: light orange headers, thin borders, centred numbers, whole numbers.
const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = { padding: "4px 10px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const, textAlign: "center" as const };
const td = { padding: "4px 10px", border: BORDER, textAlign: "center" as const };
const left = { textAlign: "left" as const };
const YELLOW = "#FFF2CC"; // preliminary / verify (same yellow as the workbook)

const kind = (v: string) => {
  const bg = kindColor(v);
  return { ...td, ...(bg ? { background: bg, color: "#000" } : {}) };
};
const selectionCell = (r: MccRow) => ({
  ...td, color: "#000", textAlign: "left" as const,
  background: r.current.designA === undefined ? "transparent" : r.breaker.status === "review" ? REPORT_COLORS.over : r.breaker.status === "verify" ? YELLOW : REPORT_COLORS.within,
});

function Results({ result }: { result: ReturnType<typeof mccRows> }) {
  if (result.errors.length) return <p role="alert" style={{ color: "crimson" }}>{result.errors.join(" ")}</p>;
  const { rows, summary: s } = result;
  const a = DEFAULT_MCC_ASSUMPTIONS;
  return (
    <section>
      <h3>Results (preliminary)</h3>
      <p>
        <b>{s.loads}</b> loads · {s.mccb} MCCB · {s.acb} EDO ACB ({s.verify} to verify) ·{" "}
        <span style={{ background: s.review ? REPORT_COLORS.over : "transparent", padding: "0 4px" }}>{s.review} need engineering review</span>
        {s.missing > 0 && <> · {s.missing} without enough data for a current</>}
      </p>
      <p style={{ fontSize: 14 }}>
        Assumptions: power factor {fmtPf(a.pfGeneral)} general feeder / {fmtPf(a.pfDol)} DOL / {fmtPf(a.pfVfd)} VFD; DOL motor efficiency {fmtInt(a.efficiency * 100)} % (provisional);
        433 V loads calculated at {fmtInt(a.calcVoltage433)} V; design current = current × {a.designMultiplier}. A rated current in the load list is used in preference to kW.
        MCCB 32 / 63 / 100 / 160 / 250 / 400 A; 500 A and above: EDO ACB; 400 to 500 A: engineering review.
      </p>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse" }}>
          <thead><tr>
            <th style={th}>Row</th><th style={th}>Description</th><th style={th}>W/S</th><th style={th}>Treated as</th>
            <th style={th}>V</th><th style={th}>kW</th><th style={th}>Entered current (A)</th><th style={th}>Calculated current (A)</th>
            <th style={th}>Current used (A)</th><th style={th}>Design current (A)</th><th style={th}>Selected breaker</th>
            <th style={th}>Relay (A)</th><th style={th}>Incomer</th><th style={th}>CT</th><th style={th}>Checks</th>
          </tr></thead>
          <tbody>
            {rows.map((r) => {
              const treated = r.type === "FEEDER" ? "Feeder" : `${r.type} motor`;
              return (
                <tr key={r.rowNumber}>
                  <td style={td}>{r.rowNumber}</td><td style={{ ...td, ...left, minWidth: 190 }}>{r.description}</td>
                  <td style={kind(r.status ?? "")}>{r.status ?? ""}</td><td style={kind(treated)}>{treated}</td>
                  <td style={td}>{r.voltageV === undefined ? "" : fmtInt(r.voltageV)}</td>
                  <td style={td}>{r.kw === undefined ? "" : fmtInt(r.kw)}</td>
                  <td style={td}>{r.ratedA === undefined ? "" : fmtInt(r.ratedA)}</td>
                  <td style={td}>{r.current.calculatedA === undefined ? "" : fmtInt(r.current.calculatedA)}</td>
                  <td style={td}>{r.current.usedA === undefined ? "" : fmtInt(r.current.usedA)}</td>
                  <td style={td}>{r.current.designA === undefined ? "" : fmtInt(r.current.designA)}</td>
                  <td style={{ ...selectionCell(r), minWidth: 150 }}>{r.current.designA === undefined ? "" : breakerLabel(r.breaker)}</td>
                  <td style={td}>{r.relayA === undefined ? "" : r.relayA.toFixed(1)}</td>
                  <td style={td}>{r.incomer ? breakerLabel(r.incomer) : ""}</td>
                  <td style={td}>{r.ctPrimaryA ? `${r.ctPrimaryA}/5 A` : ""}</td>
                  <td style={{ ...td, ...left, fontSize: 13, minWidth: 340 }}>{r.flags.join("; ")}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <p style={{ fontSize: 14 }}>
        Selections are preliminary and need engineering approval: check maximum demand, cable capacity and voltage drop, fault level, motor starting,
        discrimination and busbar rating. The downloaded workbook has the same figures as formulas, plus the Notes sheet.
      </p>
    </section>
  );
}
