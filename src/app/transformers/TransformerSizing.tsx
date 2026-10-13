"use client";

import { useMemo, useRef, useState } from "react";
import { readSheet } from "read-excel-file/universal";
import {
  DEFAULT_OPTIONS, STANDARD_MCC_FEEDERS, STANDARD_RATINGS_KVA, TRANSFORMER_OPTIONS, buildReportBlob, displayCombos, parseLoadSheet,
  sizeTransformers,
  REPORT_COLORS, failureReason, fmtInt, fmtPct, fmtPf, kindColor,
  type Candidate, type ComboResult, type RawLoadRow, type SizingOptions, type SizingResult,
} from "@/core";

// Tables follow the Excel report: light orange headers, thin borders, numbers and percentages centred,
// green / red for within / over a limit, pastels for W, S, DOL, VFD and Feeder.
// Numbers: only percentages have a decimal place; everything else is whole (power factor keeps two decimals).
const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = {
  padding: "4px 10px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const,
  textAlign: "center" as const, whiteSpace: "nowrap" as const,
};
const td = { padding: "4px 10px", border: BORDER, textAlign: "center" as const };
const left = { textAlign: "left" as const };
const shade = (background: string) => ({ background, color: "#000" });
const withinColor = (ok: boolean) => (ok ? REPORT_COLORS.within : REPORT_COLORS.over);
/** A centred cell shaded green when the value is within its limit (inclusive), light red when over. */
const limitCell = (value: number, limitPct: number) => ({ ...td, ...shade(withinColor(value <= limitPct + 1e-9)) });
/** A centred verdict cell (fits / OK / within) shaded green or light red. */
const verdictCell = (ok: boolean) => ({ ...td, ...shade(withinColor(ok)) });
/** A cell shaded with the pastel for a W / S / DOL / VFD / Feeder value, if it has one. */
const kindCell = (value: string, align: "center" | "left" = "left") => {
  const bg = kindColor(value);
  return { ...td, textAlign: align, ...(bg ? shade(bg) : {}) };
};
const GOOD = REPORT_COLORS.within;
const BAD = "crimson";

/** W / S tag shown inside lists. */
function Tag({ value }: { value: string }) {
  const bg = kindColor(value);
  return <span style={{ ...(bg ? shade(bg) : {}), padding: "0 5px", borderRadius: 3, border: BORDER }}>{value}</span>;
}

type Combination = "3" | "2";
type ReportStatus = { kind: "idle" } | { kind: "busy" } | { kind: "done"; name: string } | { kind: "error"; message: string };

function num(s: string, fallback: number, min: number, max: number) {
  const v = Number(s);
  return Number.isFinite(v) && v > min && v <= max ? v : fallback;
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/** "My loads (v2).xlsx", 433 -> "My_loads_v2_transformer_sizing_433V.xlsx" */
function reportFileName(inputName: string, volts: number) {
  const base = inputName.replace(/\.xlsx$/i, "").replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "load_list";
  return `${base}_transformer_sizing_${volts}V.xlsx`;
}

export default function TransformerSizing() {
  const [rows, setRows] = useState<RawLoadRow[] | null>(null);
  // The uploaded sheet exactly as read, included in the Excel report as "Original load list".
  const [rawSheet, setRawSheet] = useState<readonly (readonly unknown[])[]>([]);
  const [fileName, setFileName] = useState("");
  const [loadError, setLoadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [txId, setTxId] = useState<string>(TRANSFORMER_OPTIONS[0].id);
  const [combination, setCombination] = useState<Combination>("3");
  const [rating, setRating] = useState<number>(STANDARD_RATINGS_KVA[0]);
  const [eff, setEff] = useState("95");
  const [pfDol, setPfDol] = useState("0.7");
  const [pfVfd, setPfVfd] = useState("0.9");
  const [pfFeeder, setPfFeeder] = useState("0.7");
  const [includePower, setIncludePower] = useState(true);
  const [withMcc, setWithMcc] = useState(true);
  const [report, setReport] = useState<ReportStatus>({ kind: "idle" });
  const input = useRef<HTMLInputElement>(null);

  const tx = TRANSFORMER_OPTIONS.find((t) => t.id === txId)!;
  const calcV = DEFAULT_OPTIONS.calculationVoltageV[tx.secondaryV] ?? tx.secondaryV;

  const opts: SizingOptions = useMemo(() => ({
    ...DEFAULT_OPTIONS,
    efficiency: num(eff, 95, 0, 100) / 100,
    powerFactor: { DOL: num(pfDol, 0.7, 0, 1), VFD: num(pfVfd, 0.9, 0, 1), FEEDER: num(pfFeeder, 0.7, 0, 1) },
    ratingsKva: [rating],
    counts: [Number(combination)],
    mccFeeders: withMcc ? STANDARD_MCC_FEEDERS : [],
  }), [eff, pfDol, pfVfd, pfFeeder, rating, combination, withMcc]);

  const result = useMemo(() => (rows ? sizeTransformers(rows, tx.secondaryV, opts) : null), [rows, tx, opts]);
  // When the chosen rating cannot carry the load even with the maximum number of transformers,
  // find what would (any standard rating).
  const alternative = useMemo(
    () => (rows && result && result.loads.length > 0 && !result.recommended
      ? sizeTransformers(rows, tx.secondaryV, { ...opts, ratingsKva: STANDARD_RATINGS_KVA }).recommended
      : undefined),
    [rows, result, tx, opts],
  );

  async function makeReport(r: SizingResult, sourceName: string, sheet: readonly (readonly unknown[])[]) {
    setReport({ kind: "busy" });
    try {
      const blob = await buildReportBlob(r, {
        transformerLabel: tx.label, fileName: sourceName, options: opts, includePower, originalSheet: sheet,
      });
      const name = reportFileName(sourceName, r.selectedVoltageV);
      downloadBlob(blob, name);
      setReport({ kind: "done", name });
    } catch (e) {
      setReport({ kind: "error", message: `Could not create the Excel report: ${(e as Error).message}` });
    }
  }

  async function handleFile(file: File | undefined) {
    if (!file) return;
    setLoadError("");
    setReport({ kind: "idle" });
    if (!/\.xlsx$/i.test(file.name)) {
      setLoadError("Please use an Excel .xlsx file.");
      return;
    }
    try {
      const sheet = await readSheet(await file.arrayBuffer());
      const parsed = parseLoadSheet(sheet);
      if (parsed.errors.length) {
        setRows(null);
        setLoadError(parsed.errors.join(" "));
        return;
      }
      setRows(parsed.rows);
      setRawSheet(sheet);
      setFileName(file.name);
      // The Excel report is created and downloaded as soon as the file is in.
      const first = sizeTransformers(parsed.rows, tx.secondaryV, opts);
      if (first.loads.length > 0) await makeReport(first, file.name, sheet);
      else setReport({ kind: "error", message: `No usable loads at ${tx.secondaryV} V, so no Excel report was created.` });
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
      <p style={{ margin: "8px 0" }}>
        <label style={{ cursor: "pointer" }}>
          <input type="checkbox" role="switch" checked={includePower} onChange={(e) => setIncludePower(e.target.checked)} />{" "}
          Include kW, kvar and power factor in the Excel report
        </label>
        <span> — the report is created and downloaded automatically as soon as you drop the file.</span>
      </p>
      <div aria-live="polite">
        {report.kind === "busy" && <p>Preparing the Excel report…</p>}
        {report.kind === "done" && <p>✔ Excel report downloaded: <b>{report.name}</b></p>}
        {report.kind === "error" && <p role="alert" style={{ color: BAD }}>{report.message}</p>}
      </div>

      <h3>Transformers</h3>
      <div style={{ display: "grid", gap: 10 }}>
        <label>
          Voltage:{" "}
          <select value={txId} onChange={(e) => setTxId(e.target.value)}>
            {TRANSFORMER_OPTIONS.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
          </select>
          <span> Only loads at {tx.secondaryV} V are included; other voltage levels are listed as excluded. Loads given in amps are converted to kVA at {fmtInt(calcV)} V.</span>
        </label>
        <fieldset style={{ border: "1px solid #888", borderRadius: 6 }}>
          <legend>Transformer rating</legend>
          {STANDARD_RATINGS_KVA.map((r) => (
            <label key={r} style={{ marginRight: 16, whiteSpace: "nowrap" }}>
              <input type="radio" name="rating" checked={rating === r} onChange={() => setRating(r)} /> {fmtInt(r)} kVA
            </label>
          ))}
        </fieldset>
        <label>
          <input type="checkbox" checked={withMcc} onChange={(e) => setWithMcc(e.target.checked)} />{" "}
          Add PMCC/IMCC feeders to every set: {STANDARD_MCC_FEEDERS.map((f) => `${f.count} × ${f.amps} A`).join(", ")} (counted in the load; half on each side of the middle transformer of a set of 3, or of the bus coupler of a set of 2)
        </label>
        <fieldset style={{ border: "1px solid #888", borderRadius: 6 }}>
          <legend>Start with</legend>
          {([["3", "3 transformers (3 incomers, 2 bus couplers)"], ["2", "2 transformers (2 incomers, 1 bus coupler)"]] as const).map(([v, label]) => (
            <label key={v} style={{ marginRight: 16, whiteSpace: "nowrap" }}>
              <input type="radio" name="combination" checked={combination === v} onChange={() => setCombination(v)} /> {label}
            </label>
          ))}
          <div style={{ margin: "4px 0 6px" }}>
            The number of transformers is not fixed. If the loads cannot be balanced within the limits, more are added (up to {DEFAULT_OPTIONS.maxTransformers}):
            more sets of 3 first, and a set of 2 at the end for the remainder. A bus coupler sits between neighbouring incomers in a set and none between sets.
          </div>
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
          Fixed: maximum loading {fmtPct(DEFAULT_OPTIONS.normalLimit * 100)} % in normal running and {fmtPct(DEFAULT_OPTIONS.contingencyLimit * 100)} % on each remaining
          transformer with one out. Total kVA is the sum of each load&apos;s kVA. All loads count at 100 %, working (W) and standby (S) alike.
        </p>
      </details>

      {result && (
        <Results
          result={result} tx={tx.label} opts={opts} alternative={alternative}
          busy={report.kind === "busy"} onDownload={() => void makeReport(result, fileName, rawSheet)}
        />
      )}
    </div>
  );
}

const layoutText = (c: { layout: number[] }) => (c.layout.length > 1 ? ` (sets of ${c.layout.join(" + ")})` : "");

function Results({ result, tx, opts, alternative, busy, onDownload }: {
  result: SizingResult; tx: string; opts: SizingOptions; alternative?: Candidate; busy: boolean; onDownload: () => void;
}) {
  const rec = result.recommended;
  const recCombo = rec ? result.combos.find((c) => c.count === rec.count) : undefined;
  const errors = result.issues.filter((i) => i.severity === "error");
  const excluded = result.issues.filter((i) => i.severity === "excluded");
  const warnings = result.issues.filter((i) => i.severity === "warning");
  const t = result.totals;
  const sets = recCombo?.layout.length ?? 1;
  const normalPct = fmtPct(opts.normalLimit * 100);
  const contPct = fmtPct(opts.contingencyLimit * 100);

  return (
    <div>
      <h3>Result</h3>
      {result.loads.length === 0 ? (
        <p role="alert" style={{ color: BAD }}>No usable loads at {result.selectedVoltageV} V in this file.</p>
      ) : (
        <>
          <div style={{ ...(rec ? shade(GOOD) : {}), padding: "12px 16px", borderRadius: 8, border: rec ? undefined : `1px solid ${BAD}` }}>
            {rec ? (
              <div style={{ fontSize: 20 }}><b>{rec.count} × {fmtInt(rec.ratingKva)} kVA</b> transformers{layoutText(rec)}, {tx}</div>
            ) : (
              <div role="alert" style={{ color: BAD, fontSize: 18 }}>
                {result.blockers.length > 0 ? (
                  <>
                    These loads are too big for any single transformer section, so adding transformers cannot help:
                    <ul style={{ margin: "4px 0", fontSize: 15 }}>{result.blockers.map((m, i) => <li key={i}>{m}</li>)}</ul>
                  </>
                ) : (
                  <>No configuration fits: with up to {opts.maxTransformers} transformers of {fmtInt(opts.ratingsKva[0])} kVA the loads (total {fmtInt(result.totalKva)} kVA) cannot be balanced within the limits.</>
                )}
                {alternative
                  ? <> It would fit with <b>{alternative.count} × {fmtInt(alternative.ratingKva)} kVA</b>{layoutText(alternative)}.</>
                  : <> No standard rating fits within {opts.maxTransformers} transformers.</>}
              </div>
            )}
            <div>
              {opts.mccFeeders.length > 0 ? "Load list" : "Total load"} <b>{fmtInt(result.totalKva)} kVA</b> · {fmtInt(t.kw)} kW · {fmtInt(t.kvar)} kvar · power factor {fmtPf(t.powerFactor)}
            </div>
            {rec && recCombo && (
              <div>
                {recCombo.incomers} incomers, {recCombo.couplers} bus coupler{recCombo.couplers === 1 ? "" : "s"}
                {sets > 1 ? ` in ${sets} independent sets` : ""}. Shared equally: {fmtPct(rec.normalPct)} % {sets > 1 ? "on average " : "each "}in normal running
                (limit {normalPct} %); with one out, {fmtPct(rec.contingencyPct)} % each on the remaining {sets > 1 ? "transformers of that set" : `${rec.count - 1}`} (limit {contPct} %).
              </div>
            )}
          </div>
          <div style={{ margin: "12px 0" }}>
            <button onClick={onDownload} disabled={busy} style={{ padding: "6px 14px", fontSize: 15 }}>
              {busy ? "Preparing…" : "Download Excel report again"}
            </button>
            <span> for the current settings (voltage, rating, number of transformers).</span>
          </div>
          {result.combos.some((c) => c.added) && <Tried combos={result.combos} recommendedCount={rec?.count} opts={opts} />}
          {displayCombos(result).map((c) => <Combo key={c.count} combo={c} recommended={rec?.count === c.count} opts={opts} tx={tx} />)}
        </>
      )}

      {result.loads.length > 0 && (
        <>
          <h4>Loads included ({result.loads.length}: {result.loads.filter((l) => l.status === "W").length} W, {result.loads.filter((l) => l.status === "S").length} S)</h4>
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead><tr>
                <th style={th}>Row</th><th style={th}>Description</th><th style={th}>W/S</th><th style={th}>Group</th>
                <th style={th}>Treated as</th><th style={th}>Rated kW</th><th style={th}>kW (input)</th><th style={th}>kvar</th>
                <th style={th}>Power factor</th><th style={th}>kVA</th><th style={th}>Basis</th>
              </tr></thead>
              <tbody>
                {result.loads.map((l) => {
                  const treated = l.kind === "feeder" ? "Feeder" : `${l.type} motor`;
                  return (
                    <tr key={l.rowNumber}>
                      <td style={td}>{l.rowNumber}</td><td style={{ ...td, ...left }}>{l.description}</td>
                      <td style={kindCell(l.status ?? "", "center")}>{l.status ?? ""}</td><td style={{ ...td, ...left }}>{l.group ?? ""}</td>
                      <td style={kindCell(treated)}>{treated}</td>
                      <td style={td}>{l.ratedKw === undefined ? "" : fmtInt(l.ratedKw)}</td>
                      <td style={td}>{fmtInt(l.kw)}</td><td style={td}>{fmtInt(l.kvar)}</td><td style={td}>{fmtPf(l.powerFactor)}</td>
                      <td style={td}>{fmtInt(l.kva)}</td><td style={{ ...td, ...left }}>{l.basis}</td>
                    </tr>
                  );
                })}
                <tr>
                  <td style={td} /><td style={{ ...td, ...left }}><b>Total</b></td><td style={td} /><td style={td} /><td style={td} /><td style={td} />
                  <td style={td}><b>{fmtInt(t.kw)}</b></td><td style={td}><b>{fmtInt(t.kvar)}</b></td><td style={td}><b>{fmtPf(t.powerFactor)}</b></td>
                  <td style={td}><b>{fmtInt(result.totalKva)}</b></td><td style={td} />
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

/** Every number of transformers tried, in one compact table. */
function Tried({ combos, recommendedCount, opts }: { combos: ComboResult[]; recommendedCount?: number; opts: SizingOptions }) {
  return (
    <div style={{ overflowX: "auto", margin: "12px 0" }}>
      <table style={{ borderCollapse: "collapse" }}>
        <caption style={{ textAlign: "left" }}><b>Numbers of transformers tried</b> (more were added because the loads could not be balanced within the limits)</caption>
        <thead><tr>
          <th style={th}>Transformers</th><th style={th}>Sets</th><th style={th}>Max load (kVA)</th>
          <th style={th}>Average normal loading (%)</th><th style={th}>One out loading (%)</th><th style={th}>Result</th>
        </tr></thead>
        <tbody>
          {combos.map((c) => (
            <tr key={c.count}>
              <td style={td}>{c.count}</td><td style={td}>{c.layout.join(" + ")}</td>
              <td style={td}>{fmtInt(c.chosen.maxLoadKva)}</td>
              <td style={limitCell(c.chosen.normalPct, opts.normalLimit * 100)}>{fmtPct(c.chosen.normalPct)}</td>
              <td style={limitCell(c.chosen.contingencyPct, opts.contingencyLimit * 100)}>{fmtPct(c.chosen.contingencyPct)}</td>
              <td style={verdictCell(c.fits)}>{c.fits ? (c.count === recommendedCount ? "✔ recommended" : "✔ fits") : `✘ ${failureReason(c.chosen)}`}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Combo({ combo: c, recommended, opts, tx }: { combo: ComboResult; recommended: boolean; opts: SizingOptions; tx: string }) {
  const multi = c.sets.length > 1;
  const chain = (names: string[]) => names.join(" — coupler — ");
  const normalLimit = opts.normalLimit * 100;
  const contLimit = opts.contingencyLimit * 100;
  return (
    <section style={{ border: `1px solid ${recommended ? "rgb(40,160,70)" : "#888"}`, borderRadius: 8, padding: "4px 16px 12px", margin: "16px 0" }}>
      <h4 style={{ marginBottom: 4 }}>
        {c.count} transformers{layoutText(c)}: {c.fits ? <>{fmtInt(c.chosen.ratingKva)} kVA each, {tx}</> : <span style={{ color: BAD }}>does not fit ({fmtInt(c.chosen.ratingKva)} kVA shown)</span>}
        {recommended && <span style={{ ...shade(GOOD), borderRadius: 4, padding: "2px 8px", marginLeft: 8 }}>recommended</span>}
        {c.added && <span style={{ border: "1px solid #888", borderRadius: 4, padding: "2px 8px", marginLeft: 8 }}>added: fewer could not be balanced</span>}
      </h4>
      <p style={{ margin: "4px 0" }}>
        {c.incomers} incomers, {c.couplers} bus coupler{c.couplers === 1 ? "" : "s"}
        {multi
          ? <> in {c.sets.length} independent sets (one coupler between neighbouring incomers in a set, none between sets).</>
          : <>: {chain(c.sections.map((s) => s.name))}.</>}
      </p>

      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse" }}>
          <caption style={{ textAlign: "left" }}><b>Rating checked</b></caption>
          <thead><tr>
            <th style={th}>Rating</th><th style={th}>Installed kVA</th><th style={th}>{multi ? "Average normal loading (%)" : "Normal loading (%)"}</th><th style={th}>One out loading (%)</th><th style={th}>Result</th>
          </tr></thead>
          <tbody>
            {c.candidates.map((x) => (
              <tr key={x.ratingKva}>
                <td style={td}>{c.count} × {fmtInt(x.ratingKva)} kVA</td><td style={td}>{fmtInt(x.installedKva)}</td>
                <td style={limitCell(x.normalPct, normalLimit)}>{fmtPct(x.normalPct)}</td>
                <td style={limitCell(x.contingencyPct, contLimit)}>{fmtPct(x.contingencyPct)}</td>
                <td style={verdictCell(x.fits)}>{x.fits ? "✔ fits" : `✘ ${failureReason(x)}`}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {c.mccLoads.length > 0 && (
        <p style={{ margin: "8px 0" }}>
          Total load {fmtInt(c.totalKva)} kVA = load list {fmtInt(c.totalKva - c.mccLoads.reduce((a, l) => a + l.kva, 0))} kVA + {c.mccLoads.length} PMCC/IMCC feeders {fmtInt(c.mccLoads.reduce((a, l) => a + l.kva, 0))} kVA.
        </p>
      )}
      {c.sets.map((set) => (
        <div key={set.index} style={{ marginTop: 12 }}>
          {multi && (
            <h5 style={{ margin: "8px 0 4px" }}>
              Set {set.index}: {set.size} transformers, {set.incomers} incomers, {set.couplers} bus coupler{set.couplers === 1 ? "" : "s"} ({chain(set.sections.map((s) => s.name))}) ·
              load {fmtInt(set.kva)} kVA, limit {fmtInt(set.capacityKva)} kVA
            </h5>
          )}
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <caption style={{ textAlign: "left" }}><b>Load distribution, normal running</b> (limit {fmtPct(normalLimit)} %)</caption>
              <thead><tr>
                <th style={th}>Transformer / bus section</th><th style={th}>Loads</th><th style={th}>kW</th><th style={th}>kvar</th>
                <th style={th}>kVA</th><th style={th}>Loading (%)</th><th style={th}>Normal limit</th>
              </tr></thead>
              <tbody>
                {set.sections.map((s) => (
                  <tr key={s.name}>
                    <td style={{ ...td, ...left }}>
                      <details><summary>{s.name} ({fmtInt(c.chosen.ratingKva)} kVA)</summary>
                        <ul style={{ margin: "4px 0", textAlign: "left" }}>
                          {s.loads.map((l) => (
                            <li key={`${l.rowNumber}-${l.description}`}>
                              {l.description || `Row ${l.rowNumber}`} {l.status && <Tag value={l.status} />}{l.group ? ` [${l.group}]` : ""} – {fmtInt(l.kva)} kVA
                            </li>
                          ))}
                        </ul>
                      </details>
                    </td>
                    <td style={td}>{s.loads.length} ({s.loads.filter((l) => l.status === "W").length} W, {s.loads.filter((l) => l.status === "S").length} S)</td>
                    <td style={td}>{fmtInt(s.kw)}</td><td style={td}>{fmtInt(s.kvar)}</td>
                    <td style={td}>{fmtInt(s.kva)}</td><td style={limitCell(s.loadingPct, normalLimit)}>{fmtPct(s.loadingPct)}</td>
                    <td style={verdictCell(s.withinNormalLimit)}>{s.withinNormalLimit ? "✔ within" : "✘ over"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div style={{ overflowX: "auto", marginTop: 12 }}>
            <table style={{ borderCollapse: "collapse" }}>
              <caption style={{ textAlign: "left" }}><b>One transformer out</b> (remaining units {multi ? "of the set " : ""}share the load; limit {fmtPct(contLimit)} %)</caption>
              <thead><tr>
                <th style={th}>Out</th><th style={th}>Its bus load (kVA)</th><th style={th}>Each remaining (kVA)</th><th style={th}>Loading (%)</th><th style={th}>Result</th>
              </tr></thead>
              <tbody>
                {set.oneOut.map((s) => (
                  <tr key={s.outName}>
                    <td style={{ ...td, ...left }}>{s.outName}</td><td style={td}>{fmtInt(s.movedKva)}</td><td style={td}>{fmtInt(s.eachKva)}</td>
                    <td style={limitCell(s.loadingPct, contLimit)}>{fmtPct(s.loadingPct)}</td>
                    <td style={verdictCell(s.ok)}>{s.ok ? "✔ OK" : "✘ over limit"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ))}
      {c.notes.map((n, i) => <p key={i} style={{ margin: "8px 0 0", color: BAD }}>Note: {n}</p>)}
    </section>
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
