/**
 * The last load list the user dropped on any page, kept for this browser tab (sessionStorage) so that the
 * MCC Load List tab can reuse it. Reads and writes never throw: private windows may block storage.
 */
const KEY = "opd:last-load-list";

export interface StoredLoadList {
  fileName: string;
  /** The sheet exactly as read (first row = headers). */
  sheet: unknown[][];
}

export function saveLoadList(fileName: string, sheet: readonly (readonly unknown[])[]): void {
  try {
    const plain = sheet.map((r) => r.map((v) => (v instanceof Date ? v.toISOString().slice(0, 10) : v)));
    sessionStorage.setItem(KEY, JSON.stringify({ fileName, sheet: plain } satisfies StoredLoadList));
  } catch {
    /* storage unavailable: the MCC tab will ask for the file instead */
  }
}

export function loadLoadList(): StoredLoadList | null {
  try {
    const t = sessionStorage.getItem(KEY);
    return t ? (JSON.parse(t) as StoredLoadList) : null;
  } catch {
    return null;
  }
}
