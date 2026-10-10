import { ProjectSchema, SCHEMA_VERSION, type Project } from "./schema";

export type ParseResult =
  | { ok: true; project: Project }
  | { ok: false; errors: string[] };

/** Parse and validate project JSON text. Applies defaults; migrations hook in here when schemaVersion grows. */
export function parseProject(text: string): ParseResult {
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch (e) {
    return { ok: false, errors: [`Invalid JSON: ${(e as Error).message}`] };
  }
  const result = ProjectSchema.safeParse(raw);
  if (!result.success) {
    return {
      ok: false,
      errors: result.error.issues.map((i) => `${i.path.join(".") || "(root)"}: ${i.message}`),
    };
  }
  return { ok: true, project: result.data };
}

export function serializeProject(project: Project): string {
  return JSON.stringify({ ...project, schemaVersion: SCHEMA_VERSION }, null, 2);
}

export function emptyProject(name = "Untitled project"): Project {
  return ProjectSchema.parse({ schemaVersion: SCHEMA_VERSION, name });
}
