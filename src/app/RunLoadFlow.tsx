"use client";

import { useState } from "react";
import { runLoadFlow, type LoadFlowResult, type Library, type Project } from "@/core";

const cell = { padding: "4px 12px", textAlign: "right" as const };
const head = { ...cell, borderBottom: "1px solid #888" };

export default function RunLoadFlow({ project, library }: { project: Project; library: Library }) {
  const [state, setState] = useState<
    { status: "idle" } | { status: "running" } | { status: "done"; result: LoadFlowResult } | { status: "error"; message: string }
  >({ status: "idle" });

  async function onClick() {
    setState({ status: "running" });
    try {
      setState({ status: "done", result: await runLoadFlow(project, library) });
    } catch (e) {
      setState({ status: "error", message: (e as Error).message });
    }
  }

  return (
    <section>
      <button onClick={onClick} disabled={state.status === "running"} style={{ padding: "8px 16px", fontSize: 16 }}>
        {state.status === "running" ? "Running…" : "Run load flow"}
      </button>
      {state.status === "running" && <p>Calculating. The first run can take several seconds while the server wakes up.</p>}
      {state.status === "error" && <p role="alert" style={{ color: "crimson" }}>Failed: {state.message}</p>}
      {state.status === "done" && <Results result={state.result} project={project} />}
    </section>
  );
}

function Results({ result, project }: { result: LoadFlowResult; project: Project }) {
  const names = new Map([...project.buses, ...project.lines, ...project.transformers].map((e) => [e.id, e.name]));
  const label = (id: string) => (names.get(id) ? `${id} – ${names.get(id)}` : id);
  return (
    <div>
      <p>{result.converged ? "Converged." : "Did NOT converge – results are not reliable."}</p>
      <h4>Bus voltages</h4>
      <table style={{ borderCollapse: "collapse" }}>
        <thead><tr><th style={{ ...head, textAlign: "left" }}>Bus</th><th style={head}>Voltage (pu)</th><th style={head}>Angle (°)</th></tr></thead>
        <tbody>
          {Object.entries(result.buses).map(([id, b]) => (
            <tr key={id}><td style={{ ...cell, textAlign: "left" }}>{label(id)}</td><td style={cell}>{b.vmPu.toFixed(4)}</td><td style={cell}>{b.vaDeg.toFixed(3)}</td></tr>
          ))}
        </tbody>
      </table>
      <h4>Lines</h4>
      <table style={{ borderCollapse: "collapse" }}>
        <thead><tr><th style={{ ...head, textAlign: "left" }}>Line</th><th style={head}>Loading (%)</th><th style={head}>Current (A)</th><th style={head}>Loss (kW)</th></tr></thead>
        <tbody>
          {Object.entries(result.lines).map(([id, l]) => (
            <tr key={id}><td style={{ ...cell, textAlign: "left" }}>{label(id)}</td><td style={cell}>{l.loadingPercent.toFixed(1)}</td><td style={cell}>{(l.currentKa * 1000).toFixed(1)}</td><td style={cell}>{l.lossKw.toFixed(2)}</td></tr>
          ))}
        </tbody>
      </table>
      <h4>Transformers</h4>
      <table style={{ borderCollapse: "collapse" }}>
        <thead><tr><th style={{ ...head, textAlign: "left" }}>Transformer</th><th style={head}>Loading (%)</th><th style={head}>Loss (kW)</th></tr></thead>
        <tbody>
          {Object.entries(result.transformers).map(([id, t]) => (
            <tr key={id}><td style={{ ...cell, textAlign: "left" }}>{label(id)}</td><td style={cell}>{t.loadingPercent.toFixed(1)}</td><td style={cell}>{t.lossKw.toFixed(2)}</td></tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
