"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import readExcelFile from "read-excel-file/universal";
import {
  DEFAULT_MCC_ASSUMPTIONS, MAX_INCOMER_A, NORMAL_STATE, REPORT_COLORS, breakerLabel, designImccs, fmtInt, imccFeederRows, parseImccSheet,
  validateOperatingState, type ImccLoad, type OperatingState, type SwitchState,
} from "@/core";
import { useShell } from "../AppShell";

const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = { padding: "2px 6px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const, textAlign: "center" as const };
const td = { padding: "2px 6px", border: BORDER, textAlign: "center" as const };
const left = { textAlign: "left" as const };
const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:]/g, "").replace("T", "-");
const KEY = "opd:imcc-loads:v1";

interface Stored { fileName: string; sheetName: string; loads: ImccLoad[] }

export default function ImccDesign() {
  const { setImccFeeders, autoFeeders } = useShell();
  const [stored, setStored] = useState<Stored | null>(null);
  const [sheets, setSheets] = useState<{ sheet: string; data: unknown[][] }[]>([]);
  const [fileName, setFileName] = useState("");
  const [sheetIdx, setSheetIdx] = useState(0);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [dragging, setDragging] = useState(false);
  const [states, setStates] = useState<Record<string, OperatingState>>({});
  const input = useRef<HTMLInputElement>(null);

  // This tab keeps its own load list (separate from the sidebar list of the other tabs).
  useEffect(() => { try { const t = sessionStorage.getItem(KEY); if (t) setStored(JSON.parse(t) as Stored); } catch { /* none */ } }, []);

  const design = useMemo(() => (stored ? designImccs(stored.loads, DEFAULT_MCC_ASSUMPTIONS, { maxIncomerA: MAX_INCOMER_A }) : null), [stored]);
  // The only link to the other tabs: the incomer feeders (used when the sidebar switch is ON).
  useEffect(() => { setImccFeeders(design ? imccFeederRows(design) : []); }, [design, setImccFeeders]);
  const feeders = useMemo(() => (design ? imccFeederRows(design) : []), [design]);

  async function handleFile(file: File | undefined) {
    if (!file) return;
    setError(""); setNotice("");
    if (!/\.xlsx$/i.test(file.name)) { setError("Please use an Excel .xlsx file."); return; }
    try {
      const all = (await readExcelFile(await file.arrayBuffer())).map((s) => ({ sheet: s.sheet, data: s.data as unknown[][] }));
      setSheets(all); setFileName(file.name); setSheetIdx(0);
      if (all.length === 1) importSheet(all, 0, file.name);
    } catch (e) { setError(`Could not read that file: ${(e as Error).message}`); }
  }

  function importSheet(all = sheets, idx = sheetIdx, name = fileName) {
    const r = parseImccSheet(all[idx].data);
    if (r.errors.length) { setError(r.errors.join(" ")); return; }
    const s: Stored = { fileName: name, sheetName: all[idx].sheet.trim(), loads: r.loads };
    setStored(s); setSheets([]); setError(""); setStates({});
    setNotice(`Imported ${r.loads.length} rows from "${s.sheetName}".`);
    try { sessionStorage.setItem(KEY, JSON.stringify(s)); } catch { /* ignore */ }
  }

  async function download() {
    if (!design) return;
    const { imccWorkbookBuffer } = await import("@/core/imcc-workbook");
    const buf = await imccWorkbookBuffer(design, DEFAULT_MCC_ASSUMPTIONS, { sourceName: stored?.fileName, linked: autoFeeders });
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
        onDrop={(e) => { e.preventDefault(); setDragging(false); void handleFile(e.dataTransfer.files[0]); }}
        onClick={() => input.current?.click()} role="button" tabIndex={0}
        onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") input.current?.click(); }}
        style={{ border: "2px dashed #009688", borderRadius: 8, padding: "12px 16px", textAlign: "center", cursor: "pointer", background: dragging ? "#B2DFDB" : "transparent", color: dragging ? "#000" : undefined }}
      >
        {stored ? <>IMCC load list: <b>{stored.fileName}</b> ({stored.sheetName}). Drop another file to replace it.</> : <>Drag and drop the IMCC load list (.xlsx, ECR layout) here, or click to choose a file. This list is used only by this tab.</>}
        <input ref={input} type="file" accept=".xlsx" hidden onChange={(e) => { void handleFile(e.target.files?.[0]); e.target.value = ""; }} />
      </div>
      {sheets.length > 1 && (
        <p>
          Sheet: <select value={sheetIdx} onChange={(e) => setSheetIdx(Number(e.target.value))}>{sheets.map((s, n) => <option key={n} value={n}>{s.sheet.trim()}</option>)}</select>{" "}
          <button type="button" onClick={() => importSheet()}>Use this sheet</button>
        </p>
      )}
      {error && <p role="alert" style={{ color: "crimson" }}>{error}</p>}
      {notice && <p aria-live="polite">{notice}</p>}

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
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={th} rowSpan={2}>IMCC</th><th style={th} rowSpan={2}>Section</th><th style={th} colSpan={3}>Loads</th>
                  <th style={th} rowSpan={2}>Working kVA</th><th style={th} rowSpan={2}>Working A</th><th style={th} rowSpan={2}>Design A</th><th style={th} rowSpan={2}>Rating</th><th style={th} rowSpan={2}>Operating state</th>
                </tr>
                <tr><th style={th}>W</th><th style={th}>S</th><th style={th}>SP</th></tr>
              </thead>
              <tbody>
                {design.imccs.flatMap((i) => i.sections.map((s) => {
                  const st = states[i.tag] ?? NORMAL_STATE;
                  const v = validateOperatingState(st);
                  const over = s.designA > design.maxIncomerA + 1e-9;
                  return (
                    <tr key={`${i.tag}-${s.section}`}>
                      {s.section === 1 && <td style={{ ...td, fontWeight: "bold" }} rowSpan={2}>{i.tag}</td>}
                      <td style={td}>Incomer {s.section}</td><td style={td}>{s.working}</td><td style={td}>{s.standby}</td><td style={td}>{s.spare}</td>
                      <td style={td}>{fmtInt(s.workingKva)}</td><td style={td}>{fmtInt(s.workingA)}</td>
                      <td style={{ ...td, background: over ? REPORT_COLORS.over : REPORT_COLORS.within, color: "#000" }}>{fmtInt(s.designA)}</td>
                      <td style={td}>{s.section === 1 ? (s.workingA > 0 ? breakerLabel(s.breaker) : "") : `= Incomer 1 (${i.incomer2.ratingA ? breakerLabel(i.incomer2) : "-"})`}</td>
                      {s.section === 1 && (
                        <td style={{ ...td, ...left, fontSize: 12 }} rowSpan={2}>
                          {(["incomer1", "incomer2", "coupler"] as const).map((k) => (
                            <button key={k} type="button" style={{ marginRight: 3 }} onClick={() => setStates((x) => ({ ...x, [i.tag]: { ...st, [k]: toggle(st[k]) } }))}>
                              {k === "coupler" ? "Coupler" : k === "incomer1" ? "Inc 1" : "Inc 2"} {st[k]}
                            </button>
                          ))}
                          <div style={{ background: v.ok ? REPORT_COLORS.within : REPORT_COLORS.over, color: "#000", padding: "1px 4px", marginTop: 2, maxWidth: 360 }}>
                            {v.ok ? "Valid: one incomer ON, other incomer and coupler OFF." : v.errors.join(" ")}
                          </div>
                          <div style={{ marginTop: 2 }}>Bus coupler: {breakerLabel(i.coupler)} (= Incomer 1)</div>
                        </td>
                      )}
                    </tr>
                  );
                }))}
              </tbody>
            </table>
          </div>
          {design.imccs.some((i) => i.issues.length > 0) && (
            <ul style={{ margin: "4px 0", fontSize: 13 }}>
              {design.imccs.flatMap((i) => i.issues.map((x) => `${i.tag}: ${x.message}`)).slice(0, 12).map((m, n) => <li key={n}>{m}</li>)}
            </ul>
          )}

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
