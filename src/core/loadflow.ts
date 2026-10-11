import type { Library } from "./library";
import { ProjectSchema, type Project } from "./schema";
import { validateProject } from "./validate";

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

export type Prepared = { ok: true; project: Project } | { ok: false; problems: string[] };

/** Schema-check and validate a project before sending it to the solver; problems are phrased for end users. */
export function prepareForSolve(project: unknown, library: Library): Prepared {
  const parsed = ProjectSchema.safeParse(project);
  if (!parsed.success) {
    const raw = project as Record<string, { id?: string }[] | undefined>;
    return {
      ok: false,
      problems: parsed.error.issues.map((i) => {
        const [coll, idx, ...rest] = i.path;
        const el = typeof coll === "string" && typeof idx === "number" ? raw?.[coll]?.[idx]?.id : undefined;
        const where = el ? `${el} › ${rest.join(".")}` : i.path.join(".") || "project";
        return `${where}: ${i.message}`;
      }),
    };
  }
  const errors = validateProject(parsed.data, library).filter((i) => i.severity === "error");
  if (errors.length > 0) return { ok: false, problems: errors.map((e) => e.message) };
  return { ok: true, project: parsed.data };
}
