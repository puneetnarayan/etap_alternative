import Link from "next/link";
import { EXAMPLE_LIBRARY, findIslands, validateProject } from "@/core";
import { SAMPLE_PROJECT } from "@/core/sample";
import RunLoadFlow from "./RunLoadFlow";

export default function Home() {
  const issues = validateProject(SAMPLE_PROJECT, EXAMPLE_LIBRARY);
  const islands = findIslands(SAMPLE_PROJECT);
  return (
    <main>
      <h1>Open Power Design</h1>
      <p>Foundation preview: data model, equipment library, topology and validation. Load flow runs on the server via pandapower; the SLD editor comes next.</p>
      <h2>Tools</h2>
      <ul>
        <li><Link href="/transformers">Transformer sizing</Link> – drop a load list (Excel) and get the number and rating of transformers.</li>
      </ul>
      <h2>{SAMPLE_PROJECT.name}</h2>
      <ul>
        <li>{SAMPLE_PROJECT.buses.length} buses, {SAMPLE_PROJECT.lines.length} lines, {SAMPLE_PROJECT.transformers.length} transformers</li>
        <li>{islands.length} island(s); energised: {islands.filter((i) => i.energised).length}</li>
      </ul>
      <h3>Validation</h3>
      {issues.length === 0 ? <p>No issues.</p> : (
        <ul>{issues.map((i, n) => <li key={n}>[{i.severity}] {i.message}</li>)}</ul>
      )}
      <h3>Load flow</h3>
      <RunLoadFlow project={SAMPLE_PROJECT} library={EXAMPLE_LIBRARY} />
    </main>
  );
}
