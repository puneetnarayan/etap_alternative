"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import readExcelFile from "read-excel-file/universal";
import {
  DEFAULT_RULE, REPORT_COLORS, addPmcc, breakerLabel, computeImcc, deleteImcc, deletePmcc, duplicateImcc, emptyDesignProject, fmtInt, migrateProject, newId,
  normalizeMethod, normalizeStatus, parseImccSheet, pmccFeeders, pmccSummary, saveImcc, validateOperatingState,
  type BusSection, type DesignProject, type Imcc, type ImccLoad, type OperatingState, type SwitchState,
} from "@/core";
import { useShell } from "../AppShell";
import { loadDesignProject, saveDesignProject } from "../projectStore";

const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = { padding: "2px 6px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const, textAlign: "center" as const };
const td = { padding: "2px 6px", border: BORDER, textAlign: "center" as const };
const left = { textAlign: "left" as const };
const inp = { width: "100%", boxSizing: "border-box" as const, font: "inherit", padding: "1px 3px" };
const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:]/g, "").replace("T", "-");

function save(buffer: ArrayBuffer, name: string) {
  const url = URL.createObjectURL(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a");
  a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

const blankImcc = (p: DesignProject): Imcc => ({
  id: newId("imcc"), tag: "", description: "", voltageV: 433,
  incomer1: { pmccId: p.pmccs[0]?.id ?? "", feederRef: "" }, incomer2: { pmccId: p.pmccs[1]?.id ?? p.pmccs[0]?.id ?? "", feederRef: "" },
  normallyOn: 1, rule: { ...DEFAULT_RULE }, loads: [],
});

export default function ImccDesign() {
  const { assumptions, loadList } = useShell();
  const [project, setProject] = useState<DesignProject>(() => emptyDesignProject());
  const [ready, setReady] = useState(false);
  const [draft, setDraft] = useState<Imcc | null>(null);
  const [errors, setErrors] = useState<string[]>([]);
  const [notice, setNotice] = useState("");
  const [pmccTag, setPmccTag] = useState("");
  const [useList, setUseList] = useState(true);
  const [state, setState] = useState<OperatingState>({ incomer1: "ON", incomer2: "OFF", coupler: "OFF" });
  const [sheets, setSheets] = useState<{ sheet: string; data: unknown[][] }[]>([]);
  const [sheetIdx, setSheetIdx] = useState(0);
  const fileRef = useRef<HTMLInputElement>(null);
  const projRef = useRef<HTMLInputElement>(null);

  useEffect(() => { const p = loadDesignProject(); if (p) setProject(p); setReady(true); }, []);
  // One committed project: every change goes through a pure function that returns the whole new project, then is stored.
  const commit = (p: DesignProject) => { setProject(p); saveDesignProject(p); };

  const pmccName = (id: string) => project.pmccs.find((p) => p.id === id)?.tag ?? "(missing)";
  const result = useMemo(() => (draft ? computeImcc(draft, assumptions) : null), [draft, assumptions]);
  const opState = useMemo(() => (draft ? validateOperatingState(state, draft.rule) : null), [draft, state]);
  // What saving would produce, without committing it.
  const preview = useMemo(() => { if (!draft) return null; const r = saveImcc(project, draft, assumptions); return r.ok ? r.value : null; }, [draft, project, assumptions]);
  const exists = draft ? project.imccs.some((i) => i.id === draft.id) : false;
  const edit = (patch: Partial<Imcc>) => { setDraft((d) => (d ? { ...d, ...patch } : d)); setErrors([]); setNotice(""); };
  const editLoad = (id: string, patch: Partial<ImccLoad>) => edit({ loads: draft!.loads.map((l) => (l.id === id ? { ...l, ...patch } : l)) });

  function addPmccClick() {
    const r = addPmcc(project, pmccTag, useList && loadList ? { fileName: loadList.fileName, sheet: loadList.sheet } : undefined);
    if (!r.ok) { setErrors(r.errors); return; }
    commit(r.value); setPmccTag(""); setErrors([]);
  }

  function onSave() {
    if (!draft) return;
    const r = saveImcc(project, draft, assumptions);
    if (!r.ok) { setErrors(r.errors); setNotice(""); return; }
    commit(r.value); setErrors([]);
    setNotice(`Saved ${draft.tag}. Linked feeders updated in ${[...new Set(r.value.generated.filter((g) => g.imccId === draft.id).map((g) => pmccName(g.pmccId)))].join(" and ")}.`);
  }

  function onDelete(i: Imcc) {
    const ask = deleteImcc(project, i.id, false, assumptions);
    if (ask.errors.length) { setErrors(ask.errors); return; }
    if (ask.needsConfirmation) {
      const list = ask.dependencies.map((d) => `${pmccName(d.pmccId)}: ${d.description}`).join("\n");
      if (!window.confirm(`Delete ${i.tag}? These PMCC outgoing feeders (generated for it) will be removed:\n${list}\n\nManually entered PMCC feeders are not touched.`)) return;
    }
    const done = deleteImcc(project, i.id, true, assumptions);
    if (done.project) { commit(done.project); if (draft?.id === i.id) setDraft(null); setNotice(`Deleted ${i.tag}.`); }
  }

  async function readFile(file: File | undefined, target: "imcc" | "project") {
    if (!file) return;
    try {
      if (target === "project") {
        commit(migrateProject(JSON.parse(await file.text())));
        setDraft(null); setNotice(`Opened project from ${file.name}.`);
      } else {
        const all = await readExcelFile(await file.arrayBuffer());
        setSheets(all.map((s) => ({ sheet: s.sheet, data: s.data as unknown[][] }))); setSheetIdx(0);
        setNotice(`${file.name}: choose the sheet and click Import.`);
      }
    } catch (e) { setErrors([`Could not read ${file.name}: ${(e as Error).message}`]); }
  }

  function importSheet() {
    if (!draft || !sheets[sheetIdx]) return;
    const r = parseImccSheet(sheets[sheetIdx].data);
    if (r.errors.length) { setErrors(r.errors); return; }
    edit({ loads: r.loads });
    setNotice(`Imported ${r.loads.length} rows from "${sheets[sheetIdx].sheet.trim()}". ${r.warnings.join(" ")}`);
    setSheets([]);
  }

  async function exportImcc(i: Imcc) {
    const { imccWorkbookBuffer } = await import("@/core/imcc-workbook");
    save(await imccWorkbookBuffer(project, i.id, assumptions), `${i.tag}_IMCC_load_list_${stamp()}.xlsx`);
  }
  async function exportPmcc(id: string) {
    try {
      const { pmccWorkbookBuffer } = await import("@/core/imcc-workbook");
      save(await pmccWorkbookBuffer(project, id, assumptions), `${pmccName(id)}_load_list_${stamp()}.xlsx`);
    } catch (e) { setErrors([(e as Error).message]); }
  }
  function downloadProject() {
    const url = URL.createObjectURL(new Blob([JSON.stringify(project, null, 1)], { type: "application/json" }));
    const a = document.createElement("a"); a.href = url; a.download = `design_project_${stamp()}.json`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  }

  const toggle = (s: SwitchState): SwitchState => (s === "ON" ? "OFF" : "ON");

  return (
    <div>
      <p>
        <button type="button" onClick={downloadProject}>Download project (.json)</button>{" "}
        <button type="button" onClick={() => projRef.current?.click()}>Open project (.json)</button>
        <input ref={projRef} type="file" accept=".json" hidden onChange={(e) => { void readFile(e.target.files?.[0], "project"); e.target.value = ""; }} />
        <span style={{ fontSize: 13 }}> The project is also kept in this browser. {ready ? "" : "Loading…"}</span>
      </p>
      {errors.length > 0 && <div role="alert" style={{ color: "crimson" }}>{errors.map((e, n) => <p key={n} style={{ margin: "2px 0" }}>{e}</p>)}</div>}
      {notice && <p aria-live="polite">{notice}</p>}

      <h3>PMCC boards</h3>
      <p style={{ fontSize: 13 }}>
        Each PMCC carries its own load list (taken from the one in the sidebar when it is created). IMCC incomers appear in its outgoing feeders automatically.
      </p>
      <p>
        <label>PMCC tag <input style={{ width: 120 }} value={pmccTag} onChange={(e) => setPmccTag(e.target.value)} placeholder="PMCC-01" /></label>{" "}
        <label><input type="checkbox" checked={useList} onChange={(e) => setUseList(e.target.checked)} /> use the sidebar load list{loadList ? ` (${loadList.fileName})` : " (none yet)"}</label>{" "}
        <button type="button" onClick={addPmccClick}>Add PMCC</button>
      </p>
      {project.pmccs.length > 0 && (
        <table style={{ borderCollapse: "collapse" }}>
          <thead><tr><th style={th}>PMCC</th><th style={th}>Load list</th><th style={th}>Manual feeders</th><th style={th}>Auto from IMCC</th><th style={th}>Working kVA</th><th style={th}>Standby kVA</th><th style={th} /></tr></thead>
          <tbody>
            {project.pmccs.map((p) => {
              const s = pmccSummary(pmccFeeders(project, p.id, assumptions));
              return (
                <tr key={p.id}>
                  <td style={{ ...td, ...left }}>{p.tag}</td><td style={{ ...td, ...left }}>{p.fileName ?? "(none)"}</td>
                  <td style={td}>{s.manual}</td><td style={td}>{s.generated}</td><td style={td}>{fmtInt(s.workingKva)}</td><td style={td}>{fmtInt(s.standbyKva)}</td>
                  <td style={td}>
                    <button type="button" disabled={!p.sheet} onClick={() => void exportPmcc(p.id)}>Excel</button>{" "}
                    <button type="button" onClick={() => { const r = deletePmcc(project, p.id); if (r.ok) commit(r.value); else setErrors(r.errors); }}>Delete</button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}

      <h3>IMCC designs</h3>
      <p>
        <button type="button" disabled={project.pmccs.length === 0} onClick={() => { setDraft(blankImcc(project)); setErrors([]); setNotice(""); setState({ incomer1: "ON", incomer2: "OFF", coupler: "OFF" }); }}>New IMCC</button>
        {project.pmccs.length === 0 && <span style={{ fontSize: 13 }}> Add at least one PMCC first.</span>}
      </p>
      {project.imccs.length > 0 && (
        <table style={{ borderCollapse: "collapse" }}>
          <thead><tr><th style={th}>IMCC</th><th style={th}>Incomer 1 from</th><th style={th}>Incomer 2 from</th><th style={th}>Incomer rating</th><th style={th}>Loads</th><th style={th} /></tr></thead>
          <tbody>
            {project.imccs.map((i) => {
              const r = computeImcc(i, assumptions);
              return (
                <tr key={i.id}>
                  <td style={{ ...td, ...left }}>{i.tag}</td>
                  <td style={td}>{pmccName(i.incomer1.pmccId)} / {i.incomer1.feederRef}</td><td style={td}>{pmccName(i.incomer2.pmccId)} / {i.incomer2.feederRef}</td>
                  <td style={td}>{r.incomer1.ratingA ? breakerLabel(r.incomer1) : "(review)"}</td><td style={td}>{i.loads.length}</td>
                  <td style={td}>
                    <button type="button" onClick={() => { setDraft(structuredClone(i)); setErrors([]); setNotice(""); setState({ incomer1: i.normallyOn === 1 ? "ON" : "OFF", incomer2: i.normallyOn === 2 ? "ON" : "OFF", coupler: "OFF" }); }}>Edit</button>{" "}
                    <button type="button" onClick={() => { const r2 = duplicateImcc(project, i.id, assumptions); if (r2.ok) commit(r2.value); else setErrors(r2.errors); }}>Duplicate</button>{" "}
                    <button type="button" onClick={() => void exportImcc(i)}>Excel</button>{" "}
                    <button type="button" onClick={() => onDelete(i)}>Delete</button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}

      {draft && result && opState && (
        <section style={{ border: BORDER, padding: "6px 10px", marginTop: 10 }}>
          <h3 style={{ marginTop: 0 }}>{exists ? `Edit ${draft.tag || "IMCC"}` : "New IMCC"}</h3>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 6 }}>
            <label>Tag / name <input style={inp} value={draft.tag} onChange={(e) => edit({ tag: e.target.value })} placeholder="IMCC-01" /></label>
            <label>Description <input style={inp} value={draft.description} onChange={(e) => edit({ description: e.target.value })} /></label>
            <label>Voltage (V) <input style={inp} value={draft.voltageV} onChange={(e) => edit({ voltageV: Number(e.target.value) || 0 })} /></label>
            {([1, 2] as const).map((n) => {
              const src = n === 1 ? draft.incomer1 : draft.incomer2;
              const set = (patch: Partial<typeof src>) => edit(n === 1 ? { incomer1: { ...src, ...patch } } : { incomer2: { ...src, ...patch } });
              return (
                <div key={n} style={{ display: "flex", gap: 6 }}>
                  <label style={{ flex: 1 }}>Incomer {n}: upstream PMCC
                    <select style={inp} value={src.pmccId} onChange={(e) => set({ pmccId: e.target.value })}>
                      <option value="">(select)</option>
                      {project.pmccs.map((p) => <option key={p.id} value={p.id}>{p.tag}</option>)}
                    </select>
                  </label>
                  <label style={{ flex: 1 }}>PMCC feeder reference <input style={inp} value={src.feederRef} onChange={(e) => set({ feederRef: e.target.value })} /></label>
                </div>
              );
            })}
          </div>

          <h4>Operating state</h4>
          <p style={{ margin: "2px 0" }}>
            Normally ON: <select value={draft.normallyOn} onChange={(e) => { const v = Number(e.target.value) as BusSection; edit({ normallyOn: v }); setState({ incomer1: v === 1 ? "ON" : "OFF", incomer2: v === 2 ? "ON" : "OFF", coupler: "OFF" }); }}>
              <option value={1}>Incomer 1</option><option value={2}>Incomer 2</option></select>{" "}
            Check a state:{" "}
            {(["incomer1", "incomer2", "coupler"] as const).map((k) => (
              <button key={k} type="button" onClick={() => setState((s) => ({ ...s, [k]: toggle(s[k]) }))} style={{ marginRight: 4 }}>
                {k === "coupler" ? "Bus coupler" : k === "incomer1" ? "Incomer 1" : "Incomer 2"}: {state[k]}
              </button>
            ))}
          </p>
          <p role="status" style={{ margin: "2px 0", padding: "2px 6px", background: opState.ok ? REPORT_COLORS.within : REPORT_COLORS.over }}>
            {opState.ok ? opState.warnings.join(" ") || "State is valid." : opState.errors.join(" ")}
          </p>
          <p style={{ fontSize: 13, margin: "2px 0" }}>
            <label><input type="checkbox" checked={draft.rule.exactlyOneIncomerOn === false || draft.rule.couplerMustBeOff === false}
              onChange={(e) => edit({ rule: e.target.checked ? { exactlyOneIncomerOn: false, couplerMustBeOff: false, authorisedBy: draft.rule.authorisedBy } : { ...DEFAULT_RULE } })} /> change the operating philosophy (an engineering change; needs a named authoriser)</label>{" "}
            {(draft.rule.exactlyOneIncomerOn === false || draft.rule.couplerMustBeOff === false) &&
              <label>Authorised by <input style={{ width: 160 }} value={draft.rule.authorisedBy ?? ""} onChange={(e) => edit({ rule: { ...draft.rule, authorisedBy: e.target.value } })} /></label>}
          </p>

          <h4>Calculation (preliminary)</h4>
          <table style={{ borderCollapse: "collapse" }}>
            <thead><tr><th style={th}>Section</th><th style={th}>W</th><th style={th}>S</th><th style={th}>SP</th><th style={th}>Connected kVA</th><th style={th}>Working kVA</th><th style={th}>Working A</th><th style={th}>LCE kVA*</th><th style={th}>Design A</th><th style={th}>Calculated rating</th></tr></thead>
            <tbody>
              {result.sections.map((s) => (
                <tr key={s.section}>
                  <td style={td}>Incomer {s.section}</td><td style={td}>{s.working}</td><td style={td}>{s.standby}</td><td style={td}>{s.spare}</td>
                  <td style={td}>{fmtInt(s.connectedKva)}</td><td style={td}>{fmtInt(s.workingKva)}</td><td style={td}>{fmtInt(s.workingA)}</td><td style={td}>{fmtInt(s.lceKva)}</td>
                  <td style={td}>{fmtInt(s.designA)}</td><td style={td}>{s.workingA > 0 ? breakerLabel(s.breaker) : ""}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <p style={{ margin: "4px 0" }}>
            <b>Incomer 1</b> {result.incomer1.ratingA ? breakerLabel(result.incomer1) : "not calculated"} · <b>Incomer 2</b> = {result.incomer2.ratingA ? breakerLabel(result.incomer2) : "-"} · <b>Bus coupler</b> = {result.coupler.ratingA ? breakerLabel(result.coupler) : "-"} (calculated design ratings; the last two follow Incomer 1).
          </p>
          <p style={{ fontSize: 12, margin: "2px 0" }}>*LCE kVA = working kVA x 0.8 as in the reference workbook; its meaning is an open item and it is not used for sizing.</p>
          {result.issues.length > 0 && (
            <ul style={{ margin: "4px 0", fontSize: 13 }}>
              {result.issues.slice(0, 12).map((x, n) => <li key={n} style={{ color: x.severity === "error" ? "crimson" : undefined }}>{x.severity === "error" ? "Error" : "Check"}: {x.message}</li>)}
              {result.issues.length > 12 && <li>… and {result.issues.length - 12} more (all are in the Excel notes).</li>}
            </ul>
          )}

          <h4>Linked PMCC outgoing feeders (preview of what Save writes)</h4>
          {preview ? (
            <>
              <table style={{ borderCollapse: "collapse" }}>
                <thead><tr><th style={th}>PMCC</th><th style={th}>Feeder ref</th><th style={th}>Outgoing feeder</th><th style={th}>Destination</th><th style={th}>Status</th><th style={th}>Working A</th><th style={th}>kVA</th><th style={th}>PMCC feeder rating</th></tr></thead>
                <tbody>
                  {preview.generated.filter((g) => g.imccId === draft.id).map((g) => (
                    <tr key={g.id}>
                      <td style={td}>{pmccName(g.pmccId)}</td><td style={td}>{g.feederRef}</td><td style={{ ...td, ...left }}>{g.description}</td><td style={td}>{g.destination}</td>
                      <td style={td}>{g.status}</td><td style={td}>{fmtInt(g.ratedA)}</td><td style={td}>{fmtInt(g.kva)}</td><td style={td}>{g.rating?.ratingA ? breakerLabel(g.rating) : "(review)"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <p style={{ fontSize: 13, margin: "4px 0" }}>
                Dependency: {pmccName(draft.incomer1.pmccId)} feeder {draft.incomer1.feederRef} → {draft.tag || "IMCC"} Incomer 1 → bus section 1 → {result.sections[0].loads.length} loads
                {" · "}{pmccName(draft.incomer2.pmccId)} feeder {draft.incomer2.feederRef} → Incomer 2 → bus section 2 → {result.sections[1].loads.length} loads. Bus coupler OFF.
                The two incomers are alternative sources: only Incomer {draft.normallyOn}&apos;s feeder counts as working demand (W); the other is standby (S).
              </p>
              {exists && <p style={{ background: "#FFF2CC", color: "#000", padding: "2px 6px", margin: "2px 0" }}>Saving changes feeder records in {[...new Set([draft.incomer1.pmccId, draft.incomer2.pmccId].map(pmccName))].join(" and ")}.</p>}
            </>
          ) : <p style={{ fontSize: 13 }}>Fill in the tag, both upstream PMCCs and feeder references to see the linked feeders.</p>}

          <h4>Load list</h4>
          <p style={{ margin: "2px 0" }}>
            <button type="button" onClick={() => fileRef.current?.click()}>Import from Excel…</button>
            <input ref={fileRef} type="file" accept=".xlsx" hidden onChange={(e) => { void readFile(e.target.files?.[0], "imcc"); e.target.value = ""; }} />
            {sheets.length > 0 && <>{" "}<select value={sheetIdx} onChange={(e) => setSheetIdx(Number(e.target.value))}>{sheets.map((s, n) => <option key={n} value={n}>{s.sheet.trim()}</option>)}</select>{" "}<button type="button" onClick={importSheet}>Import this sheet (replaces the rows)</button></>}{" "}
            <button type="button" onClick={() => edit({ loads: [...draft.loads, { id: newId("load"), section: 1, description: "", status: "W", statusText: "W", method: "DOL" }] })}>Add row</button>
          </p>
          <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead><tr><th style={th}>Section</th><th style={th}>S. NO.</th><th style={th}>Description</th><th style={th}>kW</th><th style={th}>Current (A)</th><th style={th}>W/S/Spare</th><th style={th}>Method</th><th style={th}>Group</th><th style={th}>Note</th><th style={th} /></tr></thead>
              <tbody>
                {draft.loads.map((l) => (
                  <tr key={l.id} style={l.section === undefined || l.status === undefined ? { background: REPORT_COLORS.over } : undefined}>
                    <td style={td}><select value={l.section ?? ""} onChange={(e) => editLoad(l.id, { section: e.target.value ? (Number(e.target.value) as BusSection) : undefined })}><option value="">(none)</option><option value={1}>1</option><option value={2}>2</option></select></td>
                    <td style={td}><input style={{ ...inp, width: 48 }} value={l.sNo ?? ""} onChange={(e) => editLoad(l.id, { sNo: e.target.value })} /></td>
                    <td style={td}><input style={{ ...inp, minWidth: 220 }} value={l.description} onChange={(e) => editLoad(l.id, { description: e.target.value })} /></td>
                    <td style={td}><input style={{ ...inp, width: 56 }} value={l.kw ?? ""} onChange={(e) => editLoad(l.id, { kw: e.target.value.trim() === "" ? undefined : Number(e.target.value) })} /></td>
                    <td style={td}><input style={{ ...inp, width: 56 }} value={l.ratedA ?? ""} onChange={(e) => editLoad(l.id, { ratedA: e.target.value.trim() === "" ? undefined : Number(e.target.value) })} /></td>
                    <td style={td}><input style={{ ...inp, width: 60 }} value={l.statusText ?? l.status ?? ""} onChange={(e) => editLoad(l.id, { statusText: e.target.value, status: normalizeStatus(e.target.value) })} /></td>
                    <td style={td}>
                      <select value={l.method ?? ""} onChange={(e) => editLoad(l.id, { method: normalizeMethod(e.target.value) })}>
                        <option value="">(none)</option>{["DOL", "RDOL", "VFD", "PWR FDR"].map((m) => <option key={m}>{m}</option>)}
                      </select>
                    </td>
                    <td style={td}><input style={{ ...inp, width: 110 }} value={l.group ?? ""} onChange={(e) => editLoad(l.id, { group: e.target.value || undefined })} /></td>
                    <td style={{ ...td, ...left, fontSize: 12, maxWidth: 220 }}>{l.note ?? ""}</td>
                    <td style={td}><button type="button" onClick={() => edit({ loads: draft.loads.filter((x) => x.id !== l.id) })}>✕</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p style={{ margin: "8px 0 2px" }}>
            <button type="button" onClick={onSave}><b>Save IMCC</b></button>{" "}
            <button type="button" onClick={() => { setDraft(null); setErrors([]); }}>Close</button>
          </p>
          <p style={{ fontSize: 12, margin: "2px 0" }}>Red rows: bus section or status missing (not counted, not guessed). Current from kW uses this application&apos;s power factors, efficiency and voltage (MCC Load List assumptions).</p>
        </section>
      )}
    </div>
  );
}
