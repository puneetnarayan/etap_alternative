"use client";

import { useMemo, useState } from "react";
import { REPORT_COLORS, breakerLabel, fmtInt, kindColor, mccRows, type MccRow } from "@/core";
import { useShell } from "../AppShell";

export default function MccLoadList() {
  const { status, generate, assumptions: a, setAssumptions, loadList: stored } = useShell();
  // Edited as text, applied when it is a valid number (otherwise the previous value stays).
  const [text, setText] = useState({
    efficiency: String(Math.round(a.efficiency * 100)), pfGeneral: String(a.pfGeneral), pfDol: String(a.pfDol), pfVfd: String(a.pfVfd),
    designMultiplier: String(a.designMultiplier), calcVoltage433: String(a.calcVoltage433),
  });
  const result = useMemo(() => (stored ? mccRows(stored.sheet, a) : null), [stored, a]);
  const edit = (key: keyof typeof text, min: number, max: number, scale = 1) => (v: string) => {
    setText((t) => ({ ...t, [key]: v }));
    const n = Number(v);
    if (v.trim() !== "" && Number.isFinite(n) && n > min && n <= max) setAssumptions({ ...a, [key]: n / scale });
  };

  return (
    <div>
      <p>
        {stored ? <>Load list: <b>{stored.fileName}</b>. </> : <>No load list yet: add one at the top of the sidebar. </>}
        <button type="button" disabled={!stored || status.kind === "busy"} onClick={() => void generate()}>
          Download the MCC workbook
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
        {status.kind === "no-list" && <p>No load list yet: add one at the top of the sidebar.</p>}
        {status.kind === "error" && <p role="alert" style={{ color: "#c0392b" }}>Could not generate the workbook: {status.message}</p>}
      </div>
      <fieldset style={{ border: "1px solid #888", borderRadius: 6, margin: "12px 0" }}>
        <legend>Assumptions (editable)</legend>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: 8, padding: "8px 0" }}>
          <label>Power factor, general feeder <input style={{ width: 60 }} value={text.pfGeneral} onChange={(e) => edit("pfGeneral", 0, 1)(e.target.value)} /></label>
          <label>Power factor, DOL <input style={{ width: 60 }} value={text.pfDol} onChange={(e) => edit("pfDol", 0, 1)(e.target.value)} /></label>
          <label>Power factor, VFD <input style={{ width: 60 }} value={text.pfVfd} onChange={(e) => edit("pfVfd", 0, 1)(e.target.value)} /></label>
          <label>DOL motor efficiency (%, provisional) <input style={{ width: 60 }} value={text.efficiency} onChange={(e) => edit("efficiency", 0, 100, 100)(e.target.value)} /></label>
          <label>Design current multiplier <input style={{ width: 60 }} value={text.designMultiplier} onChange={(e) => edit("designMultiplier", 0, 5)(e.target.value)} /></label>
          <label>
            <input type="checkbox" checked={a.mccb630} onChange={(e) => setAssumptions({ ...a, mccb630: e.target.checked })} /> 630 A as MCCB (instead of ACB)
          </label>
          <label>Voltage used for 433 V loads (V) <input style={{ width: 60 }} value={text.calcVoltage433} onChange={(e) => edit("calcVoltage433", 0, 1000)(e.target.value)} /></label>
        </div>
        <p style={{ margin: "4px 0", fontSize: 14 }}>
          The workbook uses the same values (on its Notes sheet). Rules: MCCB 6 / 10 / 16 / 25 / 32 / 50 / 63 / 100 / 125 / 160 / 200 / 250 / 400 A; 630 A as MCCB or ACB; EDO ACB from 500 A up to 4000 A; 400 to 500 A engineering review unless 630 A is taken as MCCB. A rated current in the load list is used in preference to kW.
        </p>
      </fieldset>
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
  return (
    <section>
      <h3>Results (preliminary)</h3>
      <p>
        <b>{s.loads}</b> loads · {s.mccb} MCCB · {s.acb} EDO ACB ({s.verify} to verify) ·{" "}
        <span style={{ background: s.review ? REPORT_COLORS.over : "transparent", padding: "0 4px" }}>{s.review} need engineering review</span>
        {s.missing > 0 && <> · {s.missing} without enough data for a current</>}
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
                  <td style={td}>{r.ctPrimaryA ? `${r.ctPrimaryA}/1 A` : ""}</td>
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
