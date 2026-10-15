"use client";

import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { readSheet } from "read-excel-file/universal";
import { DEFAULT_MCC_ASSUMPTIONS, type MccAssumptions } from "@/core/mcc-breaker";
import { DEFAULT_ACB_MAX_LOADING, DEFAULT_INCOMER_ACB_A, INCOMER_ACB_OPTIONS, appendFeedersToSheet, designImccs, imccFeederRows, parseImccLoadList, planIntake, type ImccDesignSet, type ImccFeederRow } from "@/core";
import { loadLoadList, saveLoadList, type StoredLoadList } from "./loadListStore";

export type TabKey = "transformers" | "bom" | "imcc";

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
  /**
   * The one load list of the app (uploaded in the sidebar, kept for this browser tab). With the "Automatic PMCC feeders"
   * switch ON it also holds one outgoing feeder per IMCC incomer; with the switch OFF it is exactly the uploaded list.
   */
  loadList: StoredLoadList | null;
  /** The IMCC tab's link to the rest of the app: the incomer feeders it generates and the switch. */
  imccFeeders: readonly ImccFeederRow[];
  /** The IMCC load list (file name starting with IMCC) and its design; null until one is added. */
  imccList: StoredLoadList | null;
  imccDesign: ImccDesignSet | null;
  imccError: string;
  /** Takes dropped / chosen files: a PMCC file and / or an IMCC file, identified by their names (docs/imcc.md). */
  intake: (files: File[], opts?: { confirmed?: boolean }) => Promise<void>;
  /** One line per file handled by the last intake. */
  intakeNotice: string;
  /** Maximum loading of an ACB (fraction, default 0.9), adjustable in the IMCC GUI. */
  acbMaxLoading: number;
  setAcbMaxLoading: (v: number) => void;
  /** The incomer ACB size chosen in the IMCC GUI (A): 630, 800, 1000 (default) or 1250. */
  incomerAcbA: number;
  setIncomerAcbA: (v: number) => void;
  /**
   * A click on a sidebar tab asks the page of that tab to download its Excel (the data is already calculated when the
   * file is added). The page runs it once and clears the request.
   */
  pendingDownload: TabKey | null;
  clearDownload: () => void;
  autoFeeders: boolean;
  setAutoFeeders: (on: boolean) => void;
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

const IMCC_KEY = "opd:imcc:v1";
const IMCC_LIST_KEY = "opd:imcc-list:v1";
// The IMCC tab is a different colour (teal) from the PMCC tabs (blue).
const IMCC_IDLE = "#E3F4F1";
const IMCC_ACTIVE = "#BFE6DF";
const IMCC_BAR = "#009688";
const PMCC_ACTIVE = "#D6E6F7";
const PMCC_BAR = "#4A86C5";
const PMCC_TABS = [{ href: "/transformers", label: "Transformer sizing", icon: "🔌" }] as const;

const MCC_TIP =
  "Purpose: prepares an MCC load schedule with preliminary breaker selections (MCCB / EDO ACB) as an Excel workbook. Click to generate it from your current load list.";

export default function AppShell({ children }: { children: ReactNode }) {
  // A fresh session always starts expanded.
  const [open, setOpen] = useState(true);
  const [status, setStatus] = useState<MccStatus>({ kind: "idle" });
  const [assumptions, setAssumptions] = useState<MccAssumptions>(DEFAULT_MCC_ASSUMPTIONS);
  const [rawList, setLoadList] = useState<StoredLoadList | null>(null);
  const [imccList, setImccList] = useState<StoredLoadList | null>(null);
  const [imccError, setImccError] = useState("");
  const [intakeNotice, setIntakeNotice] = useState("");
  const [pendingDownload, setPendingDownload] = useState<TabKey | null>(null);
  const [autoFeeders, setAutoFeedersState] = useState(true); // ON by default
  const [acbMaxLoading, setAcbMaxLoadingState] = useState(DEFAULT_ACB_MAX_LOADING);
  const [incomerAcbA, setIncomerAcbState] = useState<number>(DEFAULT_INCOMER_ACB_A);
  const [loadError, setLoadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  // Read after mount: the server render has no browser storage.
  useEffect(() => {
    setLoadList(loadLoadList());
    try { const t = sessionStorage.getItem(IMCC_LIST_KEY); if (t) setImccList(JSON.parse(t) as StoredLoadList); } catch { /* none */ }
    try {
      const t = localStorage.getItem(IMCC_KEY);
      if (t) { const v = JSON.parse(t) as { auto?: boolean; acb?: number; inc?: number }; setAutoFeedersState(v.auto !== false); if (typeof v.acb === "number" && v.acb > 0 && v.acb <= 1) setAcbMaxLoadingState(v.acb); if ((INCOMER_ACB_OPTIONS as readonly number[]).includes(v.inc as number)) setIncomerAcbState(v.inc as number); }
    } catch { /* defaults */ }
  }, []);
  const persistImcc = (auto: boolean, acb: number, inc: number) => { try { localStorage.setItem(IMCC_KEY, JSON.stringify({ auto, acb, inc })); } catch { /* ignore */ } };
  const setAutoFeeders = (on: boolean) => { setAutoFeedersState(on); persistImcc(on, acbMaxLoading, incomerAcbA); };
  const setAcbMaxLoading = (v: number) => { setAcbMaxLoadingState(v); persistImcc(autoFeeders, v, incomerAcbA); };
  const setIncomerAcbA = (v: number) => { setIncomerAcbState(v); persistImcc(autoFeeders, acbMaxLoading, v); };
  // The IMCC design is computed here so that a list added anywhere (sidebar or IMCC tab) is designed and linked at once.
  const imccLoads = useMemo(() => (imccList ? parseImccLoadList(imccList.sheet) : null), [imccList]);
  const imccDesign = useMemo<ImccDesignSet | null>(() => (imccLoads && imccLoads.errors.length === 0 ? designImccs(imccLoads.loads, assumptions, { acbMaxLoading, maxIncomerA: incomerAcbA }) : null), [imccLoads, assumptions, acbMaxLoading, incomerAcbA]);
  const imccFeeders = useMemo<readonly ImccFeederRow[]>(() => (imccDesign ? imccFeederRows(imccDesign) : []), [imccDesign]);
  useEffect(() => { setImccError(imccLoads?.errors.join(" ") ?? ""); }, [imccLoads]);
  // Switch ON: the IMCC incomer feeders are rows of the load list (transformer sizing, MCC and BOM see them). OFF: nothing changes.
  const loadList = useMemo<StoredLoadList | null>(
    () => (rawList && autoFeeders && imccFeeders.length ? { ...rawList, sheet: appendFeedersToSheet(rawList.sheet, imccFeeders) } : rawList),
    [rawList, autoFeeders, imccFeeders],
  );
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

  async function intake(files: File[], opts: { confirmed?: boolean } = {}) {
    setLoadError(""); setIntakeNotice("");
    if (files.length === 0) return;
    const plan = planIntake(files.map((f) => f.name));
    if (plan.errors.length) { setLoadError(plan.errors.join(" ")); return; }
    if (plan.needsConfirmation && !opts.confirmed) {
      const ok = window.confirm(
        `Both a PMCC and an IMCC file were dropped:\n\n  PMCC: ${files[plan.pmcc!].name}\n  IMCC: ${files[plan.imcc!].name}\n\nProcess both? The PMCC file goes to the PMCC design (transformer sizing, MCC, BOM) and the IMCC file to the IMCC design.`,
      );
      if (!ok) return;
    }
    const lines: string[] = [];
    try {
      if (plan.pmcc !== undefined) {
        const f = files[plan.pmcc];
        const sheet = (await readSheet(await f.arrayBuffer())).map((r) => [...r]);
        saveLoadList(f.name, sheet);
        setLoadList({ fileName: f.name, sheet });
        setStatus({ kind: "idle" });
        lines.push(`PMCC: ${f.name}`);
      }
      if (plan.imcc !== undefined) {
        const f = files[plan.imcc];
        const sheet = (await readSheet(await f.arrayBuffer())).map((r) => [...r]);
        const parsed = parseImccLoadList(sheet);
        if (parsed.errors.length) throw new Error(`IMCC file "${f.name}": ${parsed.errors.join(" ")}`);
        setImccList({ fileName: f.name, sheet });
        try { sessionStorage.setItem(IMCC_LIST_KEY, JSON.stringify({ fileName: f.name, sheet })); } catch { /* ignore */ }
        lines.push(`IMCC: ${f.name}`);
      }
      setIntakeNotice(lines.join(" · "));
    } catch (e) {
      setLoadError(`Could not read the file: ${(e as Error).message}`);
      if (lines.length) setIntakeNotice(lines.join(" · "));
    }
  }

  /** Loads the two sample files (PMCC and IMCC) and shows the output: the transformer sizing with the IMCC incomers (switch permitting). */
  async function useSamples() {
    setLoadError("");
    try {
      const names = ["PMCC_sample_load_list.xlsx", "IMCC_sample_load_list.xlsx"];
      const files = await Promise.all(names.map(async (n) => {
        const res = await fetch(`${window.location.origin}/${n}`);
        if (!res.ok) throw new Error(`could not load ${n} (${res.status})`);
        return new File([await res.blob()], n, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      }));
      await intake(files, { confirmed: true }); // the click is the confirmation
      router.push("/transformers");
    } catch (e) {
      setLoadError(`Could not load the sample files: ${(e as Error).message}`);
    }
  }

  const onMccClick = () => {
    router.push("/mcc-load-list");
    void generate();
  };

  const ctx = useMemo(() => ({ status, generate, assumptions, setAssumptions, loadList, loadError, imccFeeders, imccList, imccDesign, imccError, intake, intakeNotice, autoFeeders, setAutoFeeders, acbMaxLoading, setAcbMaxLoading, incomerAcbA, setIncomerAcbA, pendingDownload, clearDownload: () => setPendingDownload(null) }), [status, generate, assumptions, loadList, loadError, imccFeeders, imccList, imccDesign, imccError, intakeNotice, autoFeeders, acbMaxLoading, incomerAcbA, pendingDownload]); // eslint-disable-line react-hooks/exhaustive-deps
  // Sidebar tabs: pastel, with a coloured bar on the active one.
  const tab = (active: boolean, activeBg: string, bar = "#B8AE96", idleBg = "transparent") => ({
    display: "flex", alignItems: "center", gap: 10, padding: "6px 8px", borderRadius: 8, textDecoration: "none", color: "#222",
    background: active ? activeBg : idleBg, cursor: "pointer", border: "none", borderLeft: `4px solid ${active ? bar : "transparent"}`,
    font: "inherit", width: "100%", textAlign: "left" as const, marginBottom: 2, fontWeight: active ? 600 : 400, boxSizing: "border-box" as const,
  });
  const heading = { fontSize: 11, letterSpacing: 0.6, textTransform: "uppercase" as const, color: "#7a6f58", padding: "8px 6px 2px" };

  return (
    <ShellContext.Provider value={ctx}>
      <div style={{ display: "flex", minHeight: "100vh" }}>
        <nav
          aria-label="Main"
          style={{ width: open ? 236 : 52, flexShrink: 0, borderRight: "1px solid #E4DCCB", background: "#FBF8F2", padding: 8, boxSizing: "border-box" }}
        >
          <button
            type="button" aria-label={open ? "Collapse sidebar" : "Expand sidebar"} aria-expanded={open}
            onClick={() => setOpen((o) => !o)}
            style={{ ...tab(false, "#EDE6D8"), width: 36, justifyContent: "center", fontSize: 20, marginBottom: 6 }}
          >☰</button>

          {open && <div style={heading}>Input files</div>}
          <div
            role="button" tabIndex={0} title="Drop a PMCC and/or an IMCC load list (.xlsx). The file name must start with PMCC or IMCC."
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => { e.preventDefault(); setDragging(false); void intake(Array.from(e.dataTransfer.files)); }}
            onClick={() => fileInput.current?.click()}
            onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") fileInput.current?.click(); }}
            style={{
              border: "2px dashed #C9BC9E", borderRadius: 8, padding: open ? "8px 6px" : "4px 0", textAlign: "center", cursor: "pointer",
              fontSize: 13, marginBottom: 6, overflowWrap: "anywhere", background: dragging ? "#D5EFDD" : "#FFFDF7", color: "#222",
            }}
          >
            {open
              ? (
                <>
                  <div>PMCC: {loadList ? <b>{loadList.fileName}</b> : "none"}</div>
                  <div>IMCC: {imccList ? <b>{imccList.fileName}</b> : "none"}</div>
                  <div style={{ opacity: 0.75 }}>Drop or click to add or replace (name starts with PMCC or IMCC)</div>
                </>
              )
              : <span aria-label="Load list upload">📂</span>}
            <input ref={fileInput} type="file" accept=".xlsx" multiple hidden onChange={(e) => { void intake(Array.from(e.target.files ?? [])); e.target.value = ""; }} />
          </div>
          {loadError && open && <div role="alert" style={{ color: "#b3261e", fontSize: 13, padding: "0 4px 6px" }}>{loadError}</div>}
          <button type="button" onClick={() => void useSamples()} title="Load the PMCC and IMCC sample files and show the output" style={{ ...tab(false, "#DDEBFA"), justifyContent: "center", border: "1px solid #B9D2EE", fontSize: 13 }}>
            <span aria-hidden>▶</span>{open && <span><b>Use sample files</b></span>}
          </button>
          {open && (
            <div style={{ fontSize: 12, padding: "4px 6px 0" }}>
              Samples:{" "}
              <a href="/PMCC_sample_load_list.xlsx" download>PMCC</a> · <a href="/IMCC_sample_load_list.xlsx" download>IMCC</a>
            </div>
          )}

          {open && <div style={heading}>General</div>}
          <Link href="/" style={tab(pathname === "/", "#E6E1D6")} title="Home"><span aria-hidden>🏠</span>{open && <span>Home</span>}</Link>

          {open && <div style={heading}>PMCC design</div>}
          {PMCC_TABS.map((i) => (
            <Link key={i.href} href={i.href} onClick={() => setPendingDownload("transformers")} style={tab(pathname === i.href, PMCC_ACTIVE, PMCC_BAR)} title={`${i.label}: click to open and download its Excel report`} aria-current={pathname === i.href ? "page" : undefined}>
              <span aria-hidden>{i.icon}</span>{open && <span>{i.label}</span>}
            </Link>
          ))}
          <button
            type="button" onClick={onMccClick} title={MCC_TIP} aria-label="MCC Load List / Breaker Selection"
            style={tab(pathname === "/mcc-load-list", PMCC_ACTIVE, PMCC_BAR)}
          >
            <span aria-hidden>🗄️</span>{open && <span>MCC Load List / Breaker Selection</span>}
          </button>
          <Link href="/bom" onClick={() => setPendingDownload("bom")} style={tab(pathname === "/bom", PMCC_ACTIVE, PMCC_BAR)} title="BOM: click to open and download its Excel" aria-current={pathname === "/bom" ? "page" : undefined}>
            <span aria-hidden>📋</span>{open && <span>BOM</span>}
          </Link>

          {open && <div style={heading}>IMCC design</div>}
          <Link href="/imcc" onClick={() => setPendingDownload("imcc")} style={tab(pathname === "/imcc", IMCC_ACTIVE, IMCC_BAR, IMCC_IDLE)} title="IMCC Design: click to open and download its Excel" aria-current={pathname === "/imcc" ? "page" : undefined}>
            <span aria-hidden>🧩</span>{open && <span>IMCC Design</span>}
          </Link>
          <button
            type="button" role="switch" aria-checked={autoFeeders} onClick={() => setAutoFeeders(!autoFeeders)}
            title={`Automatic PMCC feeders: ${autoFeeders ? "ON" : "OFF"}. ON: each IMCC incomer becomes an outgoing feeder of the PMCC load list (transformer sizing is recalculated). OFF: the IMCC design is independent. Click to switch.`}
            style={{ ...tab(false, IMCC_ACTIVE, IMCC_BAR, IMCC_IDLE), fontSize: 13, marginTop: 4 }}
          >
            <span className="blink" aria-hidden style={{ width: 14, height: 14, borderRadius: 7, flexShrink: 0, background: autoFeeders ? "#43a047" : "#e53935", boxShadow: `0 0 6px ${autoFeeders ? "#43a047" : "#e53935"}` }} />
            {open && <span className="blink">Automatic PMCC feeders: <b>{autoFeeders ? "ON" : "OFF"}</b></span>}
          </button>

          {open && <div style={heading}>Input formats</div>}
          <Link href="/format/pmcc" style={tab(pathname === "/format/pmcc", PMCC_ACTIVE, PMCC_BAR)} title="The PMCC load list format"><span aria-hidden>📄</span>{open && <span>PMCC input format</span>}</Link>
          <Link href="/format/imcc" style={tab(pathname === "/format/imcc", IMCC_ACTIVE, IMCC_BAR, IMCC_IDLE)} title="The IMCC load list format"><span aria-hidden>📄</span>{open && <span>IMCC input format</span>}</Link>

          {open && <div style={heading}>Analysis</div>}
          <Link href="/load-flow" style={tab(pathname === "/load-flow", "#E6E1D6")} title="Load flow"><span aria-hidden>⚡</span>{open && <span>Load flow</span>}</Link>

          <div aria-live="polite" style={{ fontSize: 13, padding: "6px 10px", overflowWrap: "anywhere" }}>
            {status.kind === "busy" && (open ? "Generating workbook…" : "…")}
            {status.kind === "done" && open && `✔ ${status.name} downloaded`}
            {status.kind === "no-list" && open && "No load list yet: add one at the top of this sidebar."}
            {status.kind === "error" && open && <span role="alert" style={{ color: "#b3261e" }}>Could not generate: {status.message}</span>}
          </div>
        </nav>
        <div style={{ flex: 1, minWidth: 0, padding: "0.75rem 1rem" }}>
          <div style={{ maxWidth: 1100, margin: "0 auto" }}>{children}</div>
        </div>
      </div>
    </ShellContext.Provider>
  );
}
