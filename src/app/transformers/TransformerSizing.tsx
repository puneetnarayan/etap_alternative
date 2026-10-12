"use client";

import { useMemo, useRef, useState } from "react";
import { readSheet } from "read-excel-file/universal";
import {
  DEFAULT_OPTIONS, TRANSFORMER_OPTIONS, parseLoadSheet, sizeTransformers,
  type RawLoadRow, type SizingOptions, type SizingResult,
} from "@/core";

const th = { padding: "4px 10px", borderBottom: "1px solid #888", textAlign: "right" as const, whiteSpace: "nowrap" as const };
const td = { padding: "4px 10px", textAlign: "right" as const };
const left = { textAlign: "left" as const };
const fmt = (n: number, d = 1) => n.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const GOOD = "rgba(40,160,70,0.18)";

function num(s: string, fallback: number, min: number, max: number) {
  const v = Number(s);
  return Number.isFinite(v) && v > min && v <= max ? v : fallback;
}

export default function TransformerSizing() {
  const [rows, setRows] = useState<RawLoadRow[] | null>(null);
  const [fileName, setFileName] = useState("");
  const [loadError, setLoadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [txId, setTxId] = useState<string>(TRANSFORMER_OPTIONS[0].id);
  const [eff, setEff] = useState("95");
  const [pfDol, setPfDol] = useState("0.7");
  const [pfVfd, setPfVfd] = useState("0.9");
  const [pfFeeder, setPfFeeder] = useState("0.7");
  const [normalLimit, setNormalLimit] = useState("80");
  const [contLimit, setContLimit] = useState("90");
  const [stepUp, setStepUp] = useState<SizingOptions["stepUp"]>("min-installed");
  const input = useRef<HTMLInputElement>(null);

  const tx = TRANSFORMER_OPTIONS.find((t) => t.id === txId)!;
  const opts: SizingOptions = useMemo(() => ({
    ...DEFAULT_OPTIONS,
    efficiency: num(eff, 95, 0, 100) / 100,
    powerFactor: {
      DOL: num(pfDol, 0.7, 0, 1), VFD: num(pfVfd, 0.9, 0, 1), FEEDER: num(pfFeeder, 0.7, 0, 1),
    },
    normalLimit: num(normalLimit, 80, 0, 100) / 100,
    contingencyLimit: num(contLimit, 90, 0, 100) / 100,
    stepUp,
  }), [eff, pfDol, pfVfd, pfFeeder, normalLimit, contLimit, stepUp]);

  const result = useMemo(() => (rows ? sizeTransformers(rows, tx.secondaryV, opts) : null), [rows, tx, opts]);

  async function handleFile(file: File | undefined) {
    if (!file) return;
    setLoadError("");
    if (!/\.xlsx$/i.test(file.name)) {
      setLoadError("Please use an Excel .xlsx file.");
      return;
    }
    try {
      const parsed = parseLoadSheet(await readSheet(await file.arrayBuffer()));
      if (parsed.errors.length) {
        setRows(null);
        setLoadError(parsed.errors.join(" "));
        return;
      }
      setRows(parsed.rows);
      setFileName(file.name);
    } catch (e) {
      setRows(null);
      setLoadError(`Could not read that file: ${(e as Error).message}`);
    }
  }

  return (
    <div>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); void handleFile(e.dataTransfer.files[0]); }}
        onClick={() => input.current?.click()}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") input.current?.click(); }}
        style={{
          border: "2px dashed #888", borderRadius: 8, padding: "28px 16px", textAlign: "center", cursor: "pointer",
          background: dragging ? GOOD : "transparent",
        }}
      >
        {fileName ? <>Loaded <b>{fileName}</b>. Drop another file to replace it.</> : <>Drag and drop your load list (.xlsx) here, or click to choose a file</>}
        <input ref={input} type="file" accept=".xlsx" hidden onChange={(e) => { void handleFile(e.target.files?.[0]); e.target.value = ""; }} />
      </div>
      {loadError && <p role="alert" style={{ color: "crimson" }}>{loadError}</p>}

      <h3>Transformer</h3>
      <label>
        Voltage:{" "}
        <select value={txId} onChange={(e) => setTxId(e.target.value)}>
          {TRANSFORMER_OPTIONS.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
        </select>
      </label>
      <p style={{ margin: "4px 0" }}>Only loads at {tx.secondaryV} V are included; other voltage levels are listed as excluded.</p>

      <details style={{ margin: "12px 0" }}>
        <summary>Assumptions (editable)</summary>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: 8, padding: "8px 0" }}>
          <label>Motor efficiency (%) <input style={{ width: 60 }} value={eff} onChange={(e) => setEff(e.target.value)} /></label>
          <label>Power factor DOL <input style={{ width: 60 }} value={pfDol} onChange={(e) => setPfDol(e.target.value)} /></label>
          <label>Power factor VFD <input style={{ width: 60 }} value={pfVfd} onChange={(e) => setPfVfd(e.target.value)} /></label>
          <label>Power factor feeder <input style={{ width: 60 }} value={pfFeeder} onChange={(e) => setPfFeeder(e.target.value)} /></label>
          <label>Max loading, normal (%) <input style={{ width: 60 }} value={normalLimit} onChange={(e) => setNormalLimit(e.target.value)} /></label>
          <label>Max loading, one out (%) <input style={{ width: 60 }} value={contLimit} onChange={(e) => setContLimit(e.target.value)} /></label>
        </div>
        <label>
          Step-up order when 2 × 2000 kVA does not fit:{" "}
          <select value={stepUp} onChange={(e) => setStepUp(e.target.value as SizingOptions["stepUp"])}>
            <option value="min-installed">Lowest total installed kVA first</option>
            <option value="count-first">Add a 3rd transformer before increasing rating</option>
          </select>
        </label>
        <p style={{ margin: "4px 0" }}>All loads are counted at 100 % (standby included). Ratings considered: {DEFAULT_OPTIONS.ratingsKva.join(", ")} kVA.</p>
      </details>

      {result && <Results result={result} tx={tx.label} opts={opts} />}
    </div>
  );
}

function Results({ result, tx, opts }: { result: SizingResult; tx: string; opts: SizingOptions }) {
  const rec = result.recommended;
  const errors = result.issues.filter((i) => i.severity === "error");
  const excluded = result.issues.filter((i) => i.severity === "excluded");
  const warnings = result.issues.filter((i) => i.severity === "warning");

  return (
    <div>
      <h3>Result</h3>
      {result.loads.length === 0 ? (
        <p role="alert" style={{ color: "crimson" }}>No usable loads at {result.selectedVoltageV} V in this file.</p>
      ) : rec ? (
        <div style={{ background: GOOD, padding: "12px 16px", borderRadius: 8 }}>
          <div style={{ fontSize: 20 }}>
            <b>{rec.count} × {fmt(rec.ratingKva, 0)} kVA</b> transformers, {tx}
          </div>
          <div>
            {result.incomers} incomers, {result.couplers} bus coupler{result.couplers === 1 ? "" : "s"}. Total load{" "}
            <b>{fmt(result.totalKva)} kVA</b> ({fmt(result.totalKw)} kW).
          </div>
          <div>
            Shared equally: {fmt(rec.normalPct)} % each in normal running (limit {fmt(opts.normalLimit * 100, 0)} %); with one out,{" "}
            {fmt(rec.contingencyPct)} % each on the remaining {rec.count - 1} (limit {fmt(opts.contingencyLimit * 100, 0)} %).
          </div>
        </div>
      ) : (
        <p role="alert" style={{ color: "crimson" }}>
          Total load {fmt(result.totalKva)} kVA does not fit any configuration considered (up to 3 × 4000 kVA).
        </p>
      )}

      {result.loads.length > 0 && (
        <>
          <h4>Configurations checked (in order of preference)</h4>
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead><tr>
                <th style={{ ...th, ...left }}>Configuration</th><th style={th}>Installed kVA</th><th style={th}>Normal %</th>
                <th style={th}>One out %</th><th style={th}>Max load kVA</th><th style={{ ...th, ...left }}>Result</th>
              </tr></thead>
              <tbody>
                {result.candidates.map((c) => {
                  const isRec = rec && c.count === rec.count && c.ratingKva === rec.ratingKva;
                  return (
                    <tr key={`${c.count}x${c.ratingKva}`} style={{ background: isRec ? GOOD : undefined }}>
                      <td style={{ ...td, ...left }}>{c.count} × {fmt(c.ratingKva, 0)} kVA</td>
                      <td style={td}>{fmt(c.installedKva, 0)}</td>
                      <td style={td}>{fmt(c.normalPct)}</td>
                      <td style={td}>{fmt(c.contingencyPct)}</td>
                      <td style={td}>{fmt(c.maxLoadKva, 0)}</td>
                      <td style={{ ...td, ...left }}>
                        {c.fits ? (isRec ? "✔ recommended" : "✔ fits") : `✘ ${[!c.normalOk && "normal", !c.contingencyOk && "one out"].filter(Boolean).join(" + ")} over limit`}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}

      {result.sections && rec && (
        <>
          <h4>Load distribution (draft)</h4>
          <p style={{ margin: "4px 0" }}>
            {result.sections.map((s) => s.name).join(" — coupler — ")}. Each load goes on one transformer&apos;s bus section,
            balanced by kVA (largest first). Loads are not split, so sections are not exactly equal.
          </p>
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead><tr>
                <th style={{ ...th, ...left }}>Transformer / bus</th><th style={th}>Loads</th><th style={th}>kVA</th>
                <th style={th}>Loading %</th><th style={{ ...th, ...left }}>Normal limit</th>
              </tr></thead>
              <tbody>
                {result.sections.map((s) => (
                  <tr key={s.name}>
                    <td style={{ ...td, ...left }}>
                      <details><summary>{s.name} ({fmt(rec.ratingKva, 0)} kVA)</summary>
                        <ul style={{ margin: "4px 0" }}>
                          {s.loads.map((l) => <li key={l.rowNumber}>{l.description || `Row ${l.rowNumber}`} – {fmt(l.kva)} kVA</li>)}
                        </ul>
                      </details>
                    </td>
                    <td style={td}>{s.loads.length}</td><td style={td}>{fmt(s.kva)}</td><td style={td}>{fmt(s.loadingPct)}</td>
                    <td style={{ ...td, ...left }}>{s.withinNormalLimit ? "✔ within" : "✘ over"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p style={{ margin: "4px 0" }}>
            With one transformer out, the one-out figure above assumes its load is shared equally by the remaining
            transformers.
          </p>
        </>
      )}

      {result.loads.length > 0 && (
        <>
          <h4>Loads included ({result.loads.length})</h4>
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead><tr>
                <th style={th}>Row</th><th style={{ ...th, ...left }}>Description</th><th style={{ ...th, ...left }}>Treated as</th>
                <th style={th}>kW</th><th style={th}>kVA</th><th style={{ ...th, ...left }}>Basis</th>
              </tr></thead>
              <tbody>
                {result.loads.map((l) => (
                  <tr key={l.rowNumber}>
                    <td style={td}>{l.rowNumber}</td><td style={{ ...td, ...left }}>{l.description}</td>
                    <td style={{ ...td, ...left }}>{l.kind === "feeder" ? "Feeder" : `${l.type} motor`}</td>
                    <td style={td}>{fmt(l.kw)}</td><td style={td}>{fmt(l.kva)}</td><td style={{ ...td, ...left }}>{l.basis}</td>
                  </tr>
                ))}
                <tr><td style={td} /><td style={{ ...td, ...left }}><b>Total</b></td><td /><td style={td}><b>{fmt(result.totalKw)}</b></td><td style={td}><b>{fmt(result.totalKva)}</b></td><td /></tr>
              </tbody>
            </table>
          </div>
        </>
      )}

      <IssueList title="Rows with errors (not included)" items={errors} color="crimson" />
      <IssueList title={`Rows excluded (not at ${result.selectedVoltageV} V)`} items={excluded} />
      <IssueList title="Warnings" items={warnings} />
    </div>
  );
}

function IssueList({ title, items, color }: { title: string; items: SizingResult["issues"]; color?: string }) {
  if (items.length === 0) return null;
  return (
    <div>
      <h4 style={{ color }}>{title}</h4>
      <ul style={{ margin: "4px 0" }}>
        {items.map((i, n) => <li key={n}>Row {i.rowNumber}{i.description ? ` (${i.description})` : ""}: {i.message}</li>)}
      </ul>
    </div>
  );
}
