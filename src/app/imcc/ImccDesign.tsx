"use client";

import { useRef, useState } from "react";
import {
  DEFAULT_MCC_ASSUMPTIONS, NORMAL_STATE, REPORT_COLORS, breakerLabel, fmtInt, fmtPct, kindColor, loadTypeOf, validateOperatingState, type ImccDesign, type OperatingState, type SectionResult, type SwitchState,
} from "@/core";
import { useShell } from "../AppShell";

const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = { padding: "2px 6px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const, textAlign: "center" as const };
const td = { padding: "2px 6px", border: BORDER, textAlign: "center" as const };
const left = { textAlign: "left" as const };
const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:]/g, "").replace("T", "-");

export default function ImccDesign() {
  const { autoFeeders, imccList, imccDesign: design, imccFeeders: feeders, imccError, intake, intakeNotice, loadError } = useShell();
  const [dragging, setDragging] = useState(false);
  const [states, setStates] = useState<Record<string, OperatingState>>({});
  const input = useRef<HTMLInputElement>(null);

  async function download() {
    if (!design) return;
    const { imccWorkbookBuffer } = await import("@/core/imcc-workbook");
    const buf = await imccWorkbookBuffer(design, DEFAULT_MCC_ASSUMPTIONS, { sourceName: imccList?.fileName, linked: autoFeeders });
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url; a.download = `IMCC_design_${stamp()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  }

  const toggle = (s: SwitchState): SwitchState => (s === "ON" ? "OFF" : "ON");
  const all = design?.issues ?? [];

  return (
    <div>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); void intake(Array.from(e.dataTransfer.files)); }}
        onClick={() => input.current?.click()} role="button" tabIndex={0}
        onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") input.current?.click(); }}
        style={{ border: "2px dashed #009688", borderRadius: 8, padding: "12px 16px", textAlign: "center", cursor: "pointer", background: dragging ? "#B2DFDB" : "transparent", color: dragging ? "#000" : undefined }}
      >
        {imccList ? <>IMCC load list: <b>{imccList.fileName}</b>. Drop another IMCC file to replace it.</> : <>Drag and drop the IMCC load list (.xlsx) here, or click to choose a file.</>}
        {" "}The file name must start with <b>IMCC</b> (a file starting with <b>PMCC</b> goes to the PMCC design; both can be dropped together).
        <input ref={input} type="file" accept=".xlsx" multiple hidden onChange={(e) => { void intake(Array.from(e.target.files ?? [])); e.target.value = ""; }} />
      </div>
      {(loadError || imccError) && <p role="alert" style={{ color: "crimson" }}>{loadError || imccError}</p>}
      {intakeNotice && <p aria-live="polite">Loaded: {intakeNotice}</p>}
      <p style={{ fontSize: 13, margin: "2px 0" }}><a href="/IMCC_sample_load_list.xlsx" download>Download a sample IMCC load list (.xlsx)</a> · <a href="/PMCC_sample_load_list.xlsx" download>sample PMCC load list (.xlsx)</a></p>

      <p style={{ margin: "6px 0", padding: "2px 8px", background: autoFeeders ? REPORT_COLORS.within : REPORT_COLORS.over, color: "#000" }}>
        Automatic PMCC feeders (sidebar switch): <b>{autoFeeders ? "ON" : "OFF"}</b>.{" "}
        {autoFeeders
          ? "Each incomer below is also an outgoing feeder of the PMCC load list: the transformer sizing is recalculated with them (the number of transformers may change)."
          : "The IMCC design is independent: nothing changes in the PMCC (not advisable: the PMCC loading then does not include the IMCCs)."}
      </p>

      {design && (
        <section>
          <h3>Design: {design.imccs.length} IMCC{design.imccs.length === 1 ? "" : "s"}</h3>
          <p style={{ fontSize: 13, margin: "2px 0" }}>
            Only working (W / w) loads are used. They fill the incomers in list order, up to {design.maxIncomerA} A of design current (working current x {DEFAULT_MCC_ASSUMPTIONS.designMultiplier}) per incomer; a new
            section starts when the next load would exceed it. Two sections (Incomer 1, Incomer 2) make one IMCC. Standby and spare rows stay with their group; they are not counted. Incomer 2 and the bus coupler follow Incomer 1.
            SS = soft starter (calculated like DOL).
          </p>
          <p><button type="button" disabled={design.imccs.length === 0} onClick={() => void download()}>Download Excel (all IMCCs)</button></p>
          {all.length > 0 && (
            <ul style={{ margin: "4px 0", fontSize: 13 }}>
              {all.slice(0, 10).map((x, n) => <li key={n} style={{ color: x.severity === "error" ? "crimson" : undefined }}>{x.severity === "error" ? "Error" : "Check"}: {x.message}</li>)}
              {all.length > 10 && <li>… and {all.length - 10} more.</li>}
            </ul>
          )}
          {design.imccs.map((i) => (
            <Board key={i.tag} imcc={i} maxA={design.maxIncomerA} state={states[i.tag] ?? NORMAL_STATE} onToggle={(k) => setStates((x) => ({ ...x, [i.tag]: { ...(x[i.tag] ?? NORMAL_STATE), [k]: toggle((x[i.tag] ?? NORMAL_STATE)[k]) } }))} />
          ))}

          <h3>PMCC outgoing feeders generated for the incomers</h3>
          <p style={{ fontSize: 13, margin: "2px 0" }}>
            Incomer 1 is working (W), Incomer 2 is standby (S): they are alternative sources, never added together as demand. Both carry the IMCC tag as group, so the transformer sizing places them
            on different transformer sections. Sized from the section&apos;s working current, not from the incomer breaker rating. {autoFeeders ? "The switch is ON: they are in the PMCC load list now." : "The switch is OFF: they are not in the PMCC load list."}
          </p>
          <table style={{ borderCollapse: "collapse", opacity: autoFeeders ? 1 : 0.5 }}>
            <thead><tr><th style={th}>Outgoing feeder</th><th style={th}>Status</th><th style={th}>Current (A)</th><th style={th}>kVA</th><th style={th}>Group</th></tr></thead>
            <tbody>
              {feeders.map((f) => (
                <tr key={f.description}><td style={{ ...td, ...left }}>{f.description}</td><td style={td}>{f.status}</td><td style={td}>{fmtInt(f.amps)}</td><td style={td}>{fmtInt(f.kva)}</td><td style={td}>{f.group}</td></tr>
              ))}
            </tbody>
          </table>
        </section>
      )}
    </div>
  );
}

/* ------------------------------ the IMCC board ----------------------------- */

const pct = (a: number, rating: number | undefined) => (rating ? a / rating : undefined);
const loadCell = (v: number | undefined) => ({ ...td, ...(v === undefined ? {} : { background: v > 1 + 1e-9 ? REPORT_COLORS.over : REPORT_COLORS.within, color: "#000" }) });
const bar = (bg: string) => ({ ...td, background: bg, color: "#000", fontWeight: "bold" as const });

/**
 * One IMCC in bus order, as in the transformer report: Incomer 1 (cyan) | its loads | Bus coupler (light yellow) | the
 * loads of section 2 | Incomer 2 (cyan). W lavender, S pink, spare gold, DOL / VFD / Feeder pastels, loading green / red.
 */
function Board({ imcc, maxA, state, onToggle }: { imcc: ImccDesign; maxA: number; state: OperatingState; onToggle: (k: keyof OperatingState) => void }) {
  const [s1, s2] = imcc.sections;
  const v = validateOperatingState(state);
  const rating = imcc.incomer1.ratingA;
  const l1 = pct(s1.designA, rating), l2 = pct(s2.designA, rating);
  const toggleBtn = (k: keyof OperatingState, label: string) => (
    <button key={k} type="button" style={{ marginRight: 3 }} onClick={() => onToggle(k)}>{label}: {state[k]}</button>
  );
  const inc = (n: 1 | 2, sec: SectionResult) => (
    <tr key={`inc${n}`}>
      <td style={bar(REPORT_COLORS.incomer)}>Incomer {n}</td>
      <td style={{ ...bar(REPORT_COLORS.incomer), ...left }} colSpan={5}>
        {imcc.tag} Incomer {n}: {n === 1 ? "calculated from its section's working load" : "design rating = Incomer 1"} · normally {n === 1 ? "ON" : "OFF"}
      </td>
      <td style={bar(REPORT_COLORS.incomer)}>{fmtInt(sec.workingKva)}</td><td style={bar(REPORT_COLORS.incomer)}>{fmtInt(sec.workingA)}</td>
      <td style={bar(REPORT_COLORS.incomer)}>{fmtInt(sec.designA)}</td>
      <td style={bar(REPORT_COLORS.incomer)}>{(n === 1 ? sec.workingA > 0 : true) && imcc.incomer1.ratingA ? breakerLabel(n === 1 ? imcc.incomer1 : imcc.incomer2) : ""}</td>
    </tr>
  );
  const rowsOf = (sec: SectionResult) => sec.loads.map((x) => {
    const st = x.load.status;
    const typeLabel = loadTypeOf(x.load.method) === "FEEDER" ? "Feeder" : loadTypeOf(x.load.method);
    const typeBg = x.load.method ? kindColor(typeLabel) : null;
    return (
      <tr key={x.load.id}>
        <td style={td}>{x.load.sNo ?? ""}</td><td style={{ ...td, ...left }}>{x.load.description}</td>
        <td style={{ ...td, ...(typeBg ? { background: typeBg, color: "#000" } : {}) }}>{x.load.method ?? ""}</td>
        <td style={{ ...td, color: "#000", background: st === "SP" ? REPORT_COLORS.spare : st ? kindColor(st) ?? undefined : REPORT_COLORS.over }}>{st ?? "?"}</td>
        <td style={td}>{x.load.kw ?? ""}</td><td style={td}>{x.usedA === undefined ? "" : fmtInt(x.usedA)}</td>
        <td style={td}>{x.kva === undefined ? "" : fmtInt(x.kva)}</td><td style={td}>{x.workingKva ? fmtInt(x.workingKva) : ""}</td>
        <td style={td}>{x.workingA ? fmtInt(x.workingA) : ""}</td><td style={td}>{x.designA === undefined ? "" : fmtInt(x.designA)}</td>
        <td style={td}>{x.breaker ? breakerLabel(x.breaker) : ""}</td>
      </tr>
    );
  });
  return (
    <section style={{ marginBottom: 14 }}>
      <h3 style={{ marginBottom: 4 }}>{imcc.tag}</h3>
      <table style={{ borderCollapse: "collapse" }}>
        <thead><tr><th style={th}>Element</th><th style={th}>Normal state</th><th style={th}>Working kVA</th><th style={th}>Working A</th><th style={th}>Design A</th><th style={th}>Calculated rating</th><th style={th}>Loading</th></tr></thead>
        <tbody>
          <tr><td style={bar(REPORT_COLORS.incomer)}>Incomer 1</td><td style={td}>ON</td><td style={td}>{fmtInt(s1.workingKva)}</td><td style={td}>{fmtInt(s1.workingA)}</td><td style={td}>{fmtInt(s1.designA)}</td><td style={td}>{s1.workingA > 0 ? breakerLabel(imcc.incomer1) : ""}</td><td style={loadCell(l1)}>{l1 === undefined ? "" : fmtPct(l1 * 100)} %</td></tr>
          <tr><td style={bar(REPORT_COLORS.incomer)}>Incomer 2</td><td style={td}>OFF</td><td style={td}>{fmtInt(s2.workingKva)}</td><td style={td}>{fmtInt(s2.workingA)}</td><td style={td}>{fmtInt(s2.designA)}</td><td style={td}>= Incomer 1 ({imcc.incomer2.ratingA ? breakerLabel(imcc.incomer2) : "-"})</td><td style={loadCell(s2.workingA > 0 ? l2 : undefined)}>{s2.workingA > 0 && l2 !== undefined ? `${fmtPct(l2 * 100)} %` : ""}</td></tr>
          <tr><td style={bar(REPORT_COLORS.coupler)}>Bus coupler</td><td style={td}>OFF</td><td style={td} colSpan={3} /><td style={td}>= Incomer 1 ({imcc.coupler.ratingA ? breakerLabel(imcc.coupler) : "-"})</td><td style={td} /></tr>
        </tbody>
      </table>
      <p style={{ margin: "4px 0", fontSize: 13 }}>
        Operating state: {toggleBtn("incomer1", "Incomer 1")}{toggleBtn("incomer2", "Incomer 2")}{toggleBtn("coupler", "Bus coupler")}
        <span style={{ background: v.ok ? REPORT_COLORS.within : REPORT_COLORS.over, color: "#000", padding: "1px 6px" }}>
          {v.ok ? "Valid: one incomer ON, the other incomer and the bus coupler OFF." : v.errors.join(" ")}
        </span>
        {" "}Limit {maxA} A design current per incomer.
      </p>
      {imcc.issues.length > 0 && <ul style={{ margin: "2px 0", fontSize: 13 }}>{imcc.issues.map((x, n) => <li key={n}>{x.message}</li>)}</ul>}
      <details open>
        <summary style={{ cursor: "pointer", fontSize: 13 }}>Bus layout: Incomer 1 · loads · bus coupler · loads · Incomer 2 ({s1.loads.length + s2.loads.length} rows)</summary>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table style={{ borderCollapse: "collapse" }}>
            <thead><tr>
              {["S. NO.", "Load", "Method", "W/S/SP", "kW", "Current A", "kVA", "Working kVA", "Working A", "Design A", "Breaker"].map((h) => <th key={h} style={{ ...th, position: "sticky", top: 0 }}>{h}</th>)}
            </tr></thead>
            <tbody>
              {inc(1, s1)}
              {rowsOf(s1)}
              <tr><td style={bar(REPORT_COLORS.coupler)}>Bus coupler</td><td style={{ ...bar(REPORT_COLORS.coupler), ...left }} colSpan={9}>Between the Incomer 1 and Incomer 2 sections · normally OFF</td><td style={bar(REPORT_COLORS.coupler)}>{imcc.coupler.ratingA ? breakerLabel(imcc.coupler) : ""}</td></tr>
              {rowsOf(s2)}
              {inc(2, s2)}
            </tbody>
          </table>
        </div>
      </details>
    </section>
  );
}
