"use client";

import { useMemo, useRef, useState } from "react";
import { readSheet } from "read-excel-file/universal";
import {
  DEFAULT_OPTIONS, STANDARD_RATINGS_KVA, TRANSFORMER_OPTIONS, buildReportBlob, parseLoadSheet, sizeTransformers,
  type Candidate, type ComboResult, type RawLoadRow, type SizingOptions, type SizingResult,
} from "@/core";

const th = { padding: "4px 10px", borderBottom: "1px solid #888", textAlign: "right" as const, whiteSpace: "nowrap" as const };
const td = { padding: "4px 10px", textAlign: "right" as const };
const left = { textAlign: "left" as const };
const fmt = (n: number, d = 1) => n.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const GOOD = "rgba(40,160,70,0.18)";
const BAD = "crimson";

type Combination = "auto" | "2" | "3";

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
  const [combination, setCombination] = useState<Combination>("auto");
  const [rating, setRating] = useState<number>(STANDARD_RATINGS_KVA[0]);
  const [eff, setEff] = useState("95");
  const [pfDol, setPfDol] = useState("0.7");
  const [pfVfd, setPfVfd] = useState("0.9");
  const [pfFeeder, setPfFeeder] = useState("0.7");
  const input = useRef<HTMLInputElement>(null);

  const tx = TRANSFORMER_OPTIONS.find((t) => t.id === txId)!;

  const opts: SizingOptions = useMemo(() => ({
    ...DEFAULT_OPTIONS,
    efficiency: num(eff, 95, 0, 100) / 100,
    powerFactor: { DOL: num(pfDol, 0.7, 0, 1), VFD: num(pfVfd, 0.9, 0, 1), FEEDER: num(pfFeeder, 0.7, 0, 1) },
    ratingsKva: [rating],
    counts: combination === "auto" ? [2, 3] : [Number(combination)],
  }), [eff, pfDol, pfVfd, pfFeeder, rating, combination]);

  const result = useMemo(() => (rows ? sizeTransformers(rows, tx.secondaryV, opts) : null), [rows, tx, opts]);
  // When the chosen rating cannot carry the load, find what would (any standard rating, same number of transformers).
  const alternative = useMemo(
    () => (rows && result && result.loads.length > 0 && !result.recommended
      ? sizeTransformers(rows, tx.secondaryV, { ...opts, ratingsKva: STANDARD_RATINGS_KVA }).recommended
      : undefined),
    [rows, result, tx, opts],
  );

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
      {loadError && <p role="alert" style={{ color: BAD }}>{loadError}</p>}

      <h3>Transformers</h3>
      <div style={{ display: "grid", gap: 10 }}>
        <label>
          Voltage:{" "}
          <select value={txId} onChange={(e) => setTxId(e.target.value)}>
            {TRANSFORMER_OPTIONS.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
          </select>
          <span> Only loads at {tx.secondaryV} V are included; other voltage levels are listed as excluded.</span>
        </label>
        <fieldset style={{ border: "1px solid #888", borderRadius: 6 }}>
          <legend>Transformer rating</legend>
          {STANDARD_RATINGS_KVA.map((r) => (
            <label key={r} style={{ marginRight: 16, whiteSpace: "nowrap" }}>
              <input type="radio" name="rating" checked={rating === r} onChange={() => setRating(r)} /> {fmt(r, 0)} kVA
            </label>
          ))}
        </fieldset>
        <fieldset style={{ border: "1px solid #888", borderRadius: 6 }}>
          <legend>Number of transformers</legend>
          {([["auto", "Automatic (2 or 3, show both)"], ["2", "2 transformers (2 incomers, 1 bus coupler)"], ["3", "3 transformers (3 incomers, 2 bus couplers)"]] as const).map(([v, label]) => (
            <label key={v} style={{ marginRight: 16, whiteSpace: "nowrap" }}>
              <input type="radio" name="combination" checked={combination === v} onChange={() => setCombination(v)} /> {label}
            </label>
          ))}
          <div style={{ margin: "4px 0 6px" }}>If the selected number of transformers does not fit, more are added automatically (up to {DEFAULT_OPTIONS.maxTransformers}).</div>
        </fieldset>
      </div>

      <details style={{ margin: "12px 0" }}>
        <summary>Assumptions (editable)</summary>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: 8, padding: "8px 0" }}>
          <label>Motor efficiency (%) <input style={{ width: 60 }} value={eff} onChange={(e) => setEff(e.target.value)} /></label>
          <label>Power factor DOL <input style={{ width: 60 }} value={pfDol} onChange={(e) => setPfDol(e.target.value)} /></label>
          <label>Power factor VFD <input style={{ width: 60 }} value={pfVfd} onChange={(e) => setPfVfd(e.target.value)} /></label>
          <label>Power factor feeder <input style={{ width: 60 }} value={pfFeeder} onChange={(e) => setPfFeeder(e.target.value)} /></label>
        </div>
        <p style={{ margin: "4px 0" }}>
          Fixed: maximum loading {fmt(DEFAULT_OPTIONS.normalLimit * 100, 0)} % in normal running and {fmt(DEFAULT_OPTIONS.contingencyLimit * 100, 0)} % on each remaining
          transformer with one out. Total kVA is the sum of each load&apos;s kVA. All loads count at 100 %, working (W) and standby (S) alike.
        </p>
      </details>

      {result && <Results result={result} tx={tx.label} opts={opts} fileName={fileName} alternative={alternative} />}
    </div>
  );
}

function Results({ result, tx, opts, fileName, alternative }: {
  result: SizingResult; tx: string; opts: SizingOptions; fileName: string; alternative?: Candidate;
}) {
  const rec = result.recommended;
  const errors = result.issues.filter((i) => i.severity === "error");
  const excluded = result.issues.filter((i) => i.severity === "excluded");
  const warnings = result.issues.filter((i) => i.severity === "warning");
  const t = result.totals;

  return (
    <div>
      <h3>Result</h3>
      {result.loads.length === 0 ? (
        <p role="alert" style={{ color: BAD }}>No usable loads at {result.selectedVoltageV} V in this file.</p>
      ) : (
        <>
          <div style={{ background: rec ? GOOD : undefined, padding: "12px 16px", borderRadius: 8, border: rec ? undefined : `1px solid ${BAD}` }}>
            {rec ? (
              <div style={{ fontSize: 20 }}><b>{rec.count} × {fmt(rec.ratingKva, 0)} kVA</b> transformers, {tx}</div>
            ) : (
              <div role="alert" style={{ color: BAD, fontSize: 18 }}>
                No configuration fits: total load {fmt(result.totalKva)} kVA is above what up to {opts.maxTransformers} transformers of {fmt(opts.ratingsKva[0], 0)} kVA can carry within the limits.
                {alternative
                  ? <> It would fit with <b>{alternative.count} × {fmt(alternative.ratingKva, 0)} kVA</b>.</>
                  : <> No standard rating fits within {opts.maxTransformers} transformers.</>}
              </div>
            )}
            <div>
              Total load <b>{fmt(result.totalKva)} kVA</b> · {fmt(t.kw)} kW · {fmt(t.kvar)} kvar · power factor {fmt(t.powerFactor, 2)}

            </div>
            {rec && (
              <div>
                {rec.count} incomers, {rec.count - 1} bus coupler{rec.count === 2 ? "" : "s"}. Shared equally: {fmt(rec.normalPct)} % each in normal running
                (limit {fmt(opts.normalLimit * 100, 0)} %); with one out, {fmt(rec.contingencyPct)} % each on the remaining {rec.count - 1} (limit {fmt(opts.contingencyLimit * 100, 0)} %).
              </div>
            )}
          </div>
          <DownloadPanel result={result} tx={tx} opts={opts} fileName={fileName} />
          {result.combos.map((c) => <Combo key={c.count} combo={c} recommended={rec?.count === c.count} opts={opts} tx={tx} />)}
        </>
      )}

      {result.loads.length > 0 && (
        <>
          <h4>Loads included ({result.loads.length}: {result.loads.filter((l) => l.status === "W").length} W, {result.loads.filter((l) => l.status === "S").length} S)</h4>
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead><tr>
                <th style={th}>Row</th><th style={{ ...th, ...left }}>Description</th><th style={th}>W/S</th><th style={{ ...th, ...left }}>Group</th>
                <th style={{ ...th, ...left }}>Treated as</th><th style={th}>Rated kW</th><th style={th}>kW (input)</th><th style={th}>kvar</th>
                <th style={th}>pf</th><th style={th}>kVA</th><th style={{ ...th, ...left }}>Basis</th>
              </tr></thead>
              <tbody>
                {result.loads.map((l) => (
                  <tr key={l.rowNumber}>
                    <td style={td}>{l.rowNumber}</td><td style={{ ...td, ...left }}>{l.description}</td><td style={td}>{l.status ?? ""}</td><td style={{ ...td, ...left }}>{l.group ?? ""}</td>
                    <td style={{ ...td, ...left }}>{l.kind === "feeder" ? "Feeder" : `${l.type} motor`}</td>
                    <td style={td}>{l.ratedKw === undefined ? "" : fmt(l.ratedKw)}</td>
                    <td style={td}>{fmt(l.kw)}</td><td style={td}>{fmt(l.kvar)}</td><td style={td}>{fmt(l.powerFactor, 2)}</td>
                    <td style={td}>{fmt(l.kva)}</td><td style={{ ...td, ...left }}>{l.basis}</td>
                  </tr>
                ))}
                <tr>
                  <td style={td} /><td style={{ ...td, ...left }}><b>Total</b></td><td /><td /><td /><td />
                  <td style={td}><b>{fmt(t.kw)}</b></td><td style={td}><b>{fmt(t.kvar)}</b></td><td style={td}><b>{fmt(t.powerFactor, 2)}</b></td>
                  <td style={td}><b>{fmt(result.totalKva)}</b></td><td />
                </tr>
              </tbody>
            </table>
          </div>
          <p style={{ margin: "4px 0" }}>
            Total kVA is the arithmetic sum of the loads&apos; kVA, so it is slightly above √(kW² + kvar²); this is the conservative figure used for sizing.
          </p>
        </>
      )}

      <IssueList title="Rows with errors (not included)" items={errors} color={BAD} />
      <IssueList title={`Rows excluded (not at ${result.selectedVoltageV} V)`} items={excluded} />
      <IssueList title="Warnings" items={warnings} />
    </div>
  );
}

function Combo({ combo: c, recommended, opts, tx }: { combo: ComboResult; recommended: boolean; opts: SizingOptions; tx: string }) {
  return (
    <section style={{ border: `1px solid ${recommended ? "rgb(40,160,70)" : "#888"}`, borderRadius: 8, padding: "4px 16px 12px", margin: "16px 0" }}>
      <h4 style={{ marginBottom: 4 }}>
        {c.count} transformers: {c.fits ? <>{fmt(c.chosen.ratingKva, 0)} kVA each, {tx}</> : <span style={{ color: BAD }}>does not fit ({fmt(c.chosen.ratingKva, 0)} kVA shown)</span>}
        {recommended && <span style={{ background: GOOD, borderRadius: 4, padding: "2px 8px", marginLeft: 8 }}>recommended</span>}
        {c.added && <span style={{ border: "1px solid #888", borderRadius: 4, padding: "2px 8px", marginLeft: 8 }}>added: fewer transformers did not fit</span>}
      </h4>
      <p style={{ margin: "4px 0" }}>
        {c.incomers} incomers, {c.couplers} bus coupler{c.couplers === 1 ? "" : "s"}: {c.sections.map((s) => s.name).join(" — coupler — ")}.
      </p>

      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse" }}>
          <caption style={{ textAlign: "left" }}><b>Ratings checked</b></caption>
          <thead><tr>
            <th style={{ ...th, ...left }}>Rating</th><th style={th}>Installed kVA</th><th style={th}>Normal %</th><th style={th}>One out %</th><th style={{ ...th, ...left }}>Result</th>
          </tr></thead>
          <tbody>
            {c.candidates.map((x) => (
              <tr key={x.ratingKva} style={{ background: x.ratingKva === c.chosen.ratingKva && c.fits ? GOOD : undefined }}>
                <td style={{ ...td, ...left }}>{c.count} × {fmt(x.ratingKva, 0)} kVA</td><td style={td}>{fmt(x.installedKva, 0)}</td>
                <td style={td}>{fmt(x.normalPct)}</td><td style={td}>{fmt(x.contingencyPct)}</td>
                <td style={{ ...td, ...left }}>
                  {x.fits ? "✔ fits" : `✘ ${[!x.normalOk && "normal", !x.contingencyOk && "one out"].filter(Boolean).join(" + ")} over limit`}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ overflowX: "auto", marginTop: 12 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <caption style={{ textAlign: "left" }}><b>Load distribution, normal running</b> (limit {fmt(opts.normalLimit * 100, 0)} %)</caption>
          <thead><tr>
            <th style={{ ...th, ...left }}>Transformer / bus section</th><th style={th}>Loads</th><th style={th}>kW</th><th style={th}>kvar</th>
            <th style={th}>kVA</th><th style={th}>Loading %</th><th style={{ ...th, ...left }}>Normal limit</th>
          </tr></thead>
          <tbody>
            {c.sections.map((s) => (
              <tr key={s.name}>
                <td style={{ ...td, ...left }}>
                  <details><summary>{s.name} ({fmt(c.chosen.ratingKva, 0)} kVA)</summary>
                    <ul style={{ margin: "4px 0" }}>
                      {s.loads.map((l) => <li key={l.rowNumber}>{l.description || `Row ${l.rowNumber}`}{l.status ? ` (${l.status})` : ""}{l.group ? ` [${l.group}]` : ""} – {fmt(l.kva)} kVA</li>)}
                    </ul>
                  </details>
                </td>
                <td style={td}>{s.loads.length} ({s.loads.filter((l) => l.status === "W").length} W, {s.loads.filter((l) => l.status === "S").length} S)</td><td style={td}>{fmt(s.kw)}</td><td style={td}>{fmt(s.kvar)}</td>
                <td style={td}>{fmt(s.kva)}</td><td style={td}>{fmt(s.loadingPct)}</td>
                <td style={{ ...td, ...left }}>{s.withinNormalLimit ? "✔ within" : "✘ over"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ overflowX: "auto", marginTop: 12 }}>
        <table style={{ borderCollapse: "collapse" }}>
          <caption style={{ textAlign: "left" }}><b>One transformer out</b> (remaining units share the whole load; limit {fmt(opts.contingencyLimit * 100, 0)} %)</caption>
          <thead><tr>
            <th style={{ ...th, ...left }}>Out</th><th style={th}>Its bus load kVA</th><th style={th}>Each remaining kVA</th><th style={th}>Loading %</th><th style={{ ...th, ...left }}>Result</th>
          </tr></thead>
          <tbody>
            {c.oneOut.map((s) => (
              <tr key={s.outName}>
                <td style={{ ...td, ...left }}>{s.outName}</td><td style={td}>{fmt(s.movedKva)}</td><td style={td}>{fmt(s.eachKva)}</td>
                <td style={td}>{fmt(s.loadingPct)}</td><td style={{ ...td, ...left }}>{s.ok ? "✔ OK" : "✘ over limit"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {c.notes.map((n, i) => <p key={i} style={{ margin: "8px 0 0", color: BAD }}>Note: {n}</p>)}
    </section>
  );
}

function DownloadPanel({ result, tx, opts, fileName }: { result: SizingResult; tx: string; opts: SizingOptions; fileName: string }) {
  const [open, setOpen] = useState(false);
  const [includePower, setIncludePower] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function download() {
    setBusy(true);
    setError("");
    try {
      const blob = await buildReportBlob(result, { transformerLabel: tx, fileName, options: opts, includePower });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `transformer-sizing-${result.selectedVoltageV}V.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setOpen(false);
    } catch (e) {
      setError(`Could not create the report: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ margin: "12px 0" }}>
      {!open ? (
        <button onClick={() => setOpen(true)} style={{ padding: "6px 14px", fontSize: 15 }}>Download Excel report…</button>
      ) : (
        <div role="dialog" aria-label="Download Excel report" style={{ border: "1px solid #888", borderRadius: 8, padding: 12, display: "inline-block" }}>
          <div><b>Excel report</b>: summary, loads, and the distribution for each transformer combination (2 and 3).</div>
          <label style={{ display: "block", margin: "10px 0", cursor: "pointer" }}>
            <input type="checkbox" role="switch" checked={includePower} onChange={(e) => setIncludePower(e.target.checked)} />{" "}
            Include kW, kvar and power factor
          </label>
          <button onClick={() => void download()} disabled={busy} style={{ padding: "6px 14px", marginRight: 8 }}>{busy ? "Preparing…" : "Download"}</button>
          <button onClick={() => setOpen(false)} disabled={busy} style={{ padding: "6px 14px" }}>Cancel</button>
          {error && <p role="alert" style={{ color: BAD }}>{error}</p>}
        </div>
      )}
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
