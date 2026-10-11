import type { Library } from "./library";
import type { Project } from "./schema";

/** Shape returned by POST /api/loadflow (see solver/loadflow.py). */
export interface LoadFlowResult {
  converged: boolean;
  buses: Record<string, { vmPu: number; vaDeg: number }>;
  lines: Record<string, { loadingPercent: number; currentKa: number; lossKw: number }>;
  transformers: Record<string, { loadingPercent: number; lossKw: number }>;
}

export async function runLoadFlow(
  project: Project,
  library: Library,
  fetchImpl: typeof fetch = fetch,
): Promise<LoadFlowResult> {
  const res = await fetchImpl("/api/loadflow", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ project, library }),
  });
  const text = await res.text();
  let body: unknown;
  try {
    body = JSON.parse(text);
  } catch {
    throw new Error(`Server returned HTTP ${res.status} (not JSON): ${text.slice(0, 200)}`);
  }
  if (!res.ok) {
    const msg = (body as { error?: string }).error ?? `HTTP ${res.status}`;
    throw new Error(msg);
  }
  return body as LoadFlowResult;
}
