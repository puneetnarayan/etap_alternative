/**
 * The design project (PMCCs, IMCCs and their links) kept in this browser (localStorage) and as a .json file.
 * Reads and writes never throw. Older / partial data is migrated by `migrateProject`.
 */
import { migrateProject, type DesignProject } from "@/core";

const KEY = "opd:design-project:v1";

export function loadDesignProject(): DesignProject | null {
  try {
    const t = localStorage.getItem(KEY);
    return t ? migrateProject(JSON.parse(t)) : null;
  } catch {
    return null;
  }
}

export function saveDesignProject(p: DesignProject): boolean {
  try {
    localStorage.setItem(KEY, JSON.stringify(p));
    return true;
  } catch {
    return false;
  }
}
