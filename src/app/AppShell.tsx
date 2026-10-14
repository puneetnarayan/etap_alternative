"use client";

import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { readSheet } from "read-excel-file/universal";
import { DEFAULT_MCC_ASSUMPTIONS, type MccAssumptions } from "@/core/mcc-breaker";
import { loadLoadList, saveLoadList, type StoredLoadList } from "./loadListStore";

export type MccStatus =
  | { kind: "idle" }
  | { kind: "busy" }
  | { kind: "done"; name: string; rows: number; missing: number }
  | { kind: "no-list" }
  | { kind: "error"; message: string };

interface Shell {
  status: MccStatus;
  /** Assumptions for the MCC workbook (edited on the MCC page; the sidebar click uses the same ones). */
  assumptions: MccAssumptions;
  setAssumptions: (a: MccAssumptions) => void;
  /** The one load list of the app (uploaded in the sidebar, kept for this browser tab). */
  loadList: StoredLoadList | null;
  loadError: string;
  /** Build the MCC workbook from the load list and download it. Only called from a click. */
  generate: () => Promise<void>;
}

const ShellContext = createContext<Shell | null>(null);
export const useShell = () => {
  const c = useContext(ShellContext);
  if (!c) throw new Error("useShell must be used inside AppShell");
  return c;
};

function download(buffer: ArrayBuffer, name: string) {
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:]/g, "").replace("T", "-");

const ITEMS = [
  { href: "/", label: "Home", icon: "🏠" },
  { href: "/transformers", label: "Transformer sizing", icon: "🔌" },
] as const;
const AFTER_MCC = [
  { href: "/bom", label: "BOM", icon: "📋" },
  { href: "/load-flow", label: "Load flow", icon: "⚡" },
] as const;

const MCC_TIP =
  "Purpose: prepares an MCC load schedule with preliminary breaker selections (MCCB / EDO ACB) as an Excel workbook. Click to generate it from your current load list.";

export default function AppShell({ children }: { children: ReactNode }) {
  // A fresh session always starts expanded.
  const [open, setOpen] = useState(true);
  const [status, setStatus] = useState<MccStatus>({ kind: "idle" });
  const [assumptions, setAssumptions] = useState<MccAssumptions>(DEFAULT_MCC_ASSUMPTIONS);
  const [loadList, setLoadList] = useState<StoredLoadList | null>(null);
  const [loadError, setLoadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  // Read after mount: the server render has no browser storage.
  useEffect(() => { setLoadList(loadLoadList()); }, []);
  const loadListRef = useRef(loadList);
  loadListRef.current = loadList;
  const assumptionsRef = useRef(assumptions);
  assumptionsRef.current = assumptions;
  const pathname = usePathname();
  const router = useRouter();

  const generate = useCallback<Shell["generate"]>(async () => {
    const list = loadListRef.current ?? loadLoadList();
    if (!list) {
      setStatus({ kind: "no-list" });
      return;
    }
    setStatus({ kind: "busy" });
    try {
      const { mccWorkbookBuffer } = await import("@/core/mcc-workbook");
      const { buffer, rows, missing } = await mccWorkbookBuffer(list.sheet, list.fileName, assumptionsRef.current);
      const name = `MCC_Load_List_Breaker_Selection_${stamp()}.xlsx`;
      download(buffer, name);
      setStatus({ kind: "done", name, rows, missing });
    } catch (e) {
      setStatus({ kind: "error", message: (e as Error).message });
    }
  }, []);

  async function handleFile(file: File | undefined) {
    if (!file) return;
    setLoadError("");
    if (!/\.xlsx$/i.test(file.name)) {
      setLoadError("Please use an Excel .xlsx file.");
      return;
    }
    try {
      const sheet = (await readSheet(await file.arrayBuffer())).map((r) => [...r]);
      saveLoadList(file.name, sheet);
      setLoadList({ fileName: file.name, sheet });
      setStatus({ kind: "idle" });
    } catch (e) {
      setLoadError(`Could not read that file: ${(e as Error).message}`);
    }
  }

  const onMccClick = () => {
    router.push("/mcc-load-list");
    void generate();
  };

  const ctx = useMemo(() => ({ status, generate, assumptions, setAssumptions, loadList, loadError }), [status, generate, assumptions, loadList, loadError]);
  const link = (active: boolean) => ({
    display: "flex", alignItems: "center", gap: 10, padding: "5px 8px", borderRadius: 6, textDecoration: "none",
    color: "inherit", background: active ? "rgba(128,128,128,0.22)" : "transparent", cursor: "pointer",
    border: "none", font: "inherit", width: "100%", textAlign: "left" as const,
  });

  return (
    <ShellContext.Provider value={ctx}>
      <div style={{ display: "flex", minHeight: "100vh" }}>
        <nav
          aria-label="Main"
          style={{ width: open ? 220 : 48, flexShrink: 0, borderRight: "1px solid rgba(128,128,128,0.4)", padding: 8, boxSizing: "border-box" }}
        >
          <button
            type="button" aria-label={open ? "Collapse sidebar" : "Expand sidebar"} aria-expanded={open}
            onClick={() => setOpen((o) => !o)}
            style={{ ...link(false), width: 36, justifyContent: "center", fontSize: 20, marginBottom: 8 }}
          >☰</button>
          <div
            role="button" tabIndex={0} title="Load list (.xlsx): used by every calculation in the app"
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => { e.preventDefault(); setDragging(false); void handleFile(e.dataTransfer.files[0]); }}
            onClick={() => fileInput.current?.click()}
            onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") fileInput.current?.click(); }}
            style={{
              border: "2px dashed #888", borderRadius: 8, padding: open ? "8px 6px" : "4px 0", textAlign: "center", cursor: "pointer",
              fontSize: 13, marginBottom: 8, overflowWrap: "anywhere", background: dragging ? "#C6EFCE" : "transparent", color: dragging ? "#000" : "inherit",
            }}
          >
            {open
              ? (loadList ? <>Load list: <b>{loadList.fileName}</b><br />Drop or click to replace</> : <>Load list (.xlsx)<br />Drag and drop here, or click to choose</>)
              : <span aria-label="Load list upload">📂</span>}
            <input ref={fileInput} type="file" accept=".xlsx" hidden onChange={(e) => { void handleFile(e.target.files?.[0]); e.target.value = ""; }} />
          </div>
          {loadError && open && <div role="alert" style={{ color: "#c0392b", fontSize: 13, padding: "0 4px 8px" }}>{loadError}</div>}
          {ITEMS.map((i) => (
            <Link key={i.href} href={i.href} style={link(pathname === i.href)} title={i.label}>
              <span aria-hidden>{i.icon}</span>{open && <span>{i.label}</span>}
            </Link>
          ))}
          <button
            type="button" onClick={onMccClick} title={MCC_TIP} aria-label="MCC Load List / Breaker Selection"
            style={link(pathname === "/mcc-load-list")}
          >
            <span aria-hidden>🗄️</span>{open && <span>MCC Load List / Breaker Selection</span>}
          </button>
          {AFTER_MCC.map((i) => (
            <Link key={i.href} href={i.href} style={link(pathname === i.href)} title={i.label}>
              <span aria-hidden>{i.icon}</span>{open && <span>{i.label}</span>}
            </Link>
          ))}
          <div aria-live="polite" style={{ fontSize: 13, padding: "6px 10px", overflowWrap: "anywhere" }}>
            {status.kind === "busy" && (open ? "Generating workbook…" : "…")}
            {status.kind === "done" && open && `✔ ${status.name} downloaded`}
            {status.kind === "no-list" && open && "No load list yet: add one at the top of this sidebar."}
            {status.kind === "error" && open && <span role="alert" style={{ color: "#c0392b" }}>Could not generate: {status.message}</span>}
          </div>
        </nav>
        <div style={{ flex: 1, minWidth: 0, padding: "0.75rem 1rem" }}>
          <div style={{ maxWidth: 1100, margin: "0 auto" }}>{children}</div>
        </div>
      </div>
    </ShellContext.Provider>
  );
}
