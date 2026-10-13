import { EXAMPLE_LIBRARY, findIslands, validateProject } from "@/core";
import { SAMPLE_PROJECT } from "@/core/sample";
import RunLoadFlow from "../RunLoadFlow";

export const metadata = { title: "Load flow – Open Power Design" };

export default function Page() {
  const issues = validateProject(SAMPLE_PROJECT, EXAMPLE_LIBRARY);
  const islands = findIslands(SAMPLE_PROJECT);
  return (
    <main>
      <h1>Load flow</h1>
      <p>
        Load flow runs on the server with pandapower. This preview uses the built-in sample plant; importing your own network comes with the
        single-line diagram editor.
      </p>
      <h2>{SAMPLE_PROJECT.name}</h2>
      <ul>
        <li>{SAMPLE_PROJECT.buses.length} buses, {SAMPLE_PROJECT.lines.length} lines, {SAMPLE_PROJECT.transformers.length} transformers</li>
        <li>{islands.length} island(s); energised: {islands.filter((i) => i.energised).length}</li>
      </ul>
      <h3>Validation</h3>
      {issues.length === 0 ? <p>No issues.</p> : (
        <ul>{issues.map((i, n) => <li key={n}>[{i.severity}] {i.message}</li>)}</ul>
      )}
      <h3>Run</h3>
      <RunLoadFlow project={SAMPLE_PROJECT} library={EXAMPLE_LIBRARY} />
    </main>
  );
}
