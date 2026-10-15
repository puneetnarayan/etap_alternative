/**
 * Which uploaded file is which: the file name starts with PMCC (PMCC load list) or IMCC (IMCC load list), any case,
 * after optional spaces. Several files may be dropped at once into either drop zone; each goes to its own logic.
 */
export type FileKind = "PMCC" | "IMCC";

export function classifyFile(name: string): FileKind | null {
  const m = /^\s*(pmcc|imcc)/i.exec(name);
  return m ? (m[1].toUpperCase() as FileKind) : null;
}

export interface IntakePlan {
  pmcc?: number;
  imcc?: number;
  errors: string[];
  /** True when both a PMCC and an IMCC file were dropped together: the user is asked to confirm. */
  needsConfirmation: boolean;
}

export function planIntake(names: readonly string[]): IntakePlan {
  const errors: string[] = [];
  const idx: Partial<Record<FileKind, number[]>> = {};
  names.forEach((n, i) => {
    const k = classifyFile(n);
    if (!/\.xlsx$/i.test(n)) errors.push(`"${n}": please use an Excel .xlsx file.`);
    else if (!k) errors.push(`"${n}": the file name must start with PMCC or IMCC so the data can be identified.`);
    else (idx[k] ??= []).push(i);
  });
  for (const k of ["PMCC", "IMCC"] as const) {
    if ((idx[k]?.length ?? 0) > 1) errors.push(`More than one ${k} file was dropped (${idx[k]!.map((i) => `"${names[i]}"`).join(", ")}): drop one ${k} file at a time.`);
  }
  const ok = errors.length === 0;
  return { pmcc: ok ? idx.PMCC?.[0] : undefined, imcc: ok ? idx.IMCC?.[0] : undefined, errors, needsConfirmation: ok && idx.PMCC !== undefined && idx.IMCC !== undefined };
}
