"use client";

import { useEffect, useRef, useState } from "react";
import {
  DEFAULT_INCOMER_ACB_A, INCOMER_ACB_OPTIONS, NORMAL_STATE, REPORT_COLORS, breakerLabel, fmtInt, fmtPct, kindColor, loadTypeOf, switchToFailover, switchToNormal, validateOperatingState,
  type BusResult, type ImccDesign, type OperatingState, type SwitchState,
} from "@/core";
import { useShell } from "../AppShell";

const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = { padding: "2px 6px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const, textAlign: "center" as const };
const td = { padding: "2px 6px", border: BORDER, textAlign: "center" as const };
const left = { textAlign: "left" as const };
const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:]/g, "").replace("T", "-");

export default function ImccDesign() {
  const { pendingDownload, clearDownload, assumptions, autoFeeders, imccList, imccDesign: design, imccFeeders: feeders, imccError, intake, intakeNotice, loadError, acbMaxLoading, setAcbMaxLoading, incomerAcbA, setIncomerAcbA } = useShell();
  const [acbText, setAcbText] = useState(String(Math.round(acbMaxLoading * 100)));
  const [dragging, setDragging] = useState(false);
  const [states, setStates] = useState<Record<string, OperatingState>>({});
  const input = useRef<HTMLInputElement>(null);

  // Clicking the IMCC Design tab in the sidebar downloads the IMCC workbook of the (already calculated) design.
  useEffect(() => {
    if (pendingDownload !== "imcc") return;
    clearDownload();
    if (design && design.imccs.length) void download();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingDownload]);

  async function download() {
    if (!design) return;
    const { imccWorkbookBuffer } = await import("@/core/imcc-workbook");
    const buf = await imccWorkbookBuffer(design, assumptions, { sourceName: imccList?.fileName, linked: autoFeeders });
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url; a.download = `IMCC_design_${stamp()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  }

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
            Only working (W / w) loads are used. They fill an IMCC in list order up to {fmtInt(design.maxDesignA)} A of design current (working current x {assumptions.designMultiplier}) = the {design.maxIncomerA} A ACB at
            {" "}{fmtPct(design.acbMaxLoading * 100)} % maximum loading; then the next IMCC starts. Incomer 1 is at the beginning of the panel; Incomer 2 and the bus coupler are side by side at the other end, and all three have the same
            rating because each carries the whole IMCC. Standby and spare rows stay with their group; they are not counted. SS = soft starter (calculated like DOL).
          </p>
          <fieldset style={{ margin: "6px 0", padding: "4px 10px" }}>
            <legend>Incomer ACB (Incomer 1, Incomer 2 and bus coupler)</legend>
            {INCOMER_ACB_OPTIONS.map((r) => (
              <label key={r} style={{ marginRight: 16 }}>
                <input type="radio" name="incomer-acb" checked={incomerAcbA === r} onChange={() => setIncomerAcbA(r)} /> {r} A ACB{r === DEFAULT_INCOMER_ACB_A ? " (default)" : ""}
              </label>
            ))}
            <span style={{ marginLeft: 8 }}>
              <label>ACB maximum loading (%){" "}
                <input style={{ width: 60 }} value={acbText} onChange={(e) => { setAcbText(e.target.value); const n = Number(e.target.value); if (e.target.value.trim() !== "" && Number.isFinite(n) && n >= 50 && n <= 100) setAcbMaxLoading(n / 100); }} />
              </label>
            </span>
            <div style={{ fontSize: 12 }}>The largest design current per IMCC is the ACB size x its maximum loading = {fmtInt(design.maxDesignA)} A (adjustable: 50 to 100 %, default 90 %). The number of IMCCs follows.</div>
          </fieldset>
          <p>
            <button type="button" disabled={design.imccs.length === 0} onClick={() => void download()}>Download Excel (all IMCCs)</button>{" "}
            <b>{design.imccs.length} IMCC{design.imccs.length === 1 ? "" : "s"}</b> will be calculated ({design.imccs.length * 2} PMCC incomer feeders{autoFeeders ? "" : ", not added: switch OFF"}).
          </p>
          {all.length > 0 && (
            <ul style={{ margin: "4px 0", fontSize: 13 }}>
              {all.slice(0, 10).map((x, n) => <li key={n} style={{ color: x.severity === "error" ? "crimson" : undefined }}>{x.severity === "error" ? "Error" : "Check"}: {x.message}</li>)}
              {all.length > 10 && <li>… and {all.length - 10} more.</li>}
            </ul>
          )}
          {design.imccs.map((i) => (
            <Board key={i.tag} imcc={i} state={states[i.tag] ?? NORMAL_STATE} setState={(st) => setStates((x) => ({ ...x, [i.tag]: st }))} />
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

const onOff = (v: SwitchState) => ({ ...td, background: v === "ON" ? REPORT_COLORS.within : REPORT_COLORS.over, color: "#000", fontWeight: "bold" as const });
const bar = (bg: string) => ({ ...td, background: bg, color: "#000", fontWeight: "bold" as const });
const flip = (v: SwitchState): SwitchState => (v === "ON" ? "OFF" : "ON");

/**
 * One IMCC in panel order, as in the Excel: Incomer 1 (cyan) at the beginning | the loads | Incomer 2 (cyan) and the bus
 * coupler (light yellow) side by side at the other end. ON is light green, OFF light red; W lavender, S pink, spare gold,
 * DOL / VFD / Feeder pastels; loading green within / red over the limit.
 */
function Board({ imcc, state, setState }: { imcc: ImccDesign; state: OperatingState; setState: (s: OperatingState) => void }) {
  const bus: BusResult = imcc.bus;
  const v = validateOperatingState(state);
  const [steps, setSteps] = useState<string[]>([]);
  const loadBg = (x: number | undefined) => (x === undefined ? undefined : x > imcc.loadingLimit + 1e-9 ? REPORT_COLORS.over : REPORT_COLORS.within);
  const rate = (b: ImccDesign["incomer1"]) => (b.ratingA ? breakerLabel(b) : bus.workingA > 0 ? "(review)" : "");
  const set = (k: keyof OperatingState) => { setSteps([]); setState({ ...state, [k]: flip(state[k]) }); };
  const element = (label: string, color: string, k: keyof OperatingState, normal: SwitchState, fail: SwitchState) => (
    <tr key={label}>
      <td style={bar(color)}>{label}</td>
      <td style={onOff(normal)}>{normal}</td><td style={onOff(fail)}>{fail}</td>
      <td style={td}>{fmtInt(bus.workingKva)}</td><td style={td}>{fmtInt(bus.workingA)}</td><td style={td}>{fmtInt(bus.designA)}</td>
      <td style={td}>{rate(imcc.incomer1)}</td>
      <td style={{ ...td, background: loadBg(imcc.loading), color: "#000" }}>{imcc.loading === undefined ? "" : `${fmtPct(imcc.loading * 100)} %`}</td>
      <td style={onOff(state[k])}>
        <button type="button" style={{ background: "transparent", border: "none", font: "inherit", fontWeight: "bold", cursor: "pointer", padding: 0 }} title="Click to switch (checked against the interlock)" onClick={() => set(k)}>{state[k]}</button>
      </td>
    </tr>
  );
  const row = (x: BusResult["loads"][number]) => {
    const st = x.load.status;
    const typeLabel = loadTypeOf(x.load.method) === "FEEDER" ? "Feeder" : loadTypeOf(x.load.method);
    const typeBg = x.load.method ? kindColor(typeLabel) : null;
    return (
      <tr key={x.load.id}>
        <td style={td}>{x.load.srcRow ?? ""}</td><td style={{ ...td, ...left }}>{x.load.description}</td>
        <td style={{ ...td, ...(typeBg ? { background: typeBg, color: "#000" } : {}) }}>{x.load.method ?? ""}</td>
        <td style={{ ...td, color: "#000", background: st === "SP" ? REPORT_COLORS.spare : st ? kindColor(st) ?? undefined : REPORT_COLORS.over }}>{st ?? "?"}</td>
        <td style={td}>{x.load.kw ?? ""}</td>
        <td style={td}>{x.enteredA === undefined ? "" : fmtInt(x.enteredA)}</td><td style={td}>{x.calculatedA === undefined ? "" : fmtInt(x.calculatedA)}</td>
        <td style={{ ...td, fontWeight: "bold" }} title={x.enteredA !== undefined ? "Entered current is used in preference to the calculated one" : "Calculated from kW"}>{x.usedA === undefined ? "" : fmtInt(x.usedA)}</td>
        <td style={td}>{x.kva === undefined ? "" : fmtInt(x.kva)}</td><td style={td}>{x.workingKva ? fmtInt(x.workingKva) : ""}</td>
        <td style={td}>{x.workingA ? fmtInt(x.workingA) : ""}</td><td style={td}>{x.designA === undefined ? "" : fmtInt(x.designA)}</td>
        <td style={td}>{x.breaker ? breakerLabel(x.breaker) : ""}</td>
      </tr>
    );
  };
  const incRow = (n: 1 | 2) => (
    <tr key={`inc${n}`}>
      <td style={bar(REPORT_COLORS.incomer)}>Incomer {n}</td>
      <td style={{ ...bar(REPORT_COLORS.incomer), ...left }} colSpan={7}>{imcc.tag} Incomer {n} · {n === 1 ? "beginning of the panel" : "other end, next to the bus coupler"}</td>
      <td style={bar(REPORT_COLORS.incomer)}>{fmtInt(bus.workingKva)}</td><td style={bar(REPORT_COLORS.incomer)}>{fmtInt(bus.workingA)}</td>
      <td style={bar(REPORT_COLORS.incomer)}>{fmtInt(bus.designA)}</td><td style={bar(REPORT_COLORS.incomer)}>{rate(n === 1 ? imcc.incomer1 : imcc.incomer2)}</td>
    </tr>
  );
  return (
    <section style={{ marginBottom: 14 }}>
      <h3 style={{ marginBottom: 4 }}>{imcc.tag}</h3>
      <table style={{ borderCollapse: "collapse" }}>
        <thead><tr><th style={th}>Element</th><th style={th}>Normal</th><th style={th}>Incomer 1 unhealthy</th><th style={th}>Working kVA</th><th style={th}>Working A</th><th style={th}>Design A</th><th style={th}>Calculated rating</th><th style={th}>Loading (limit {fmtPct(imcc.loadingLimit * 100)} %)</th><th style={th}>Now</th></tr></thead>
        <tbody>
          {element("Incomer 1", REPORT_COLORS.incomer, "incomer1", "ON", "OFF")}
          {element("Bus coupler", REPORT_COLORS.coupler, "coupler", "OFF", "ON")}
          {element("Incomer 2", REPORT_COLORS.incomer, "incomer2", "OFF", "ON")}
        </tbody>
      </table>
      <p style={{ margin: "4px 0", fontSize: 13 }}>
        <button type="button" onClick={() => { const r = switchToFailover(state); setState(r.state); setSteps(r.steps); }}>Incomer 1 unhealthy</button>{" "}
        <button type="button" onClick={() => { const r = switchToNormal(state); setState(r.state); setSteps(r.steps); }}>Incomer 1 healthy again</button>{" "}
        <span style={{ background: v.ok ? REPORT_COLORS.within : REPORT_COLORS.over, color: "#000", padding: "1px 6px" }}>{v.ok ? v.warnings.join(" ") : v.errors.join(" ")}</span>
        {steps.length > 0 && <span> Sequence: {steps.join(" → ")}.</span>}
      </p>
      {imcc.issues.length > 0 && <ul style={{ margin: "2px 0", fontSize: 13 }}>{imcc.issues.map((x, n) => <li key={n}>{x.message}</li>)}</ul>}
      <details open>
        <summary style={{ cursor: "pointer", fontSize: 13 }}>Panel: Incomer 1 · loads · bus coupler · Incomer 2 ({bus.loads.length} load rows)</summary>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table style={{ borderCollapse: "collapse" }}>
            <thead><tr>
              {["Row", "Load", "Method", "W/S/SP", "kW", "Entered A", "Calculated from kW (A)", "Used A", "kVA", "Working kVA", "Working A", "Design A", "Breaker"].map((h) => <th key={h} style={{ ...th, position: "sticky", top: 0 }}>{h}</th>)}
            </tr></thead>
            <tbody>
              {incRow(1)}
              {bus.loads.map(row)}
              <tr>
                <td style={bar(REPORT_COLORS.coupler)}>Bus coupler</td>
                <td style={{ ...bar(REPORT_COLORS.coupler), ...left }} colSpan={7}>Next to Incomer 2 (below it) · switched ON with it if Incomer 1 is unhealthy</td>
                <td style={bar(REPORT_COLORS.coupler)}>{fmtInt(bus.workingKva)}</td><td style={bar(REPORT_COLORS.coupler)}>{fmtInt(bus.workingA)}</td>
                <td style={bar(REPORT_COLORS.coupler)}>{fmtInt(bus.designA)}</td><td style={bar(REPORT_COLORS.coupler)}>{rate(imcc.coupler)}</td>
              </tr>
              {incRow(2)}
            </tbody>
          </table>
        </div>
      </details>
    </section>
  );
}
