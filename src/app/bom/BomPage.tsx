"use client";

import { useEffect, useMemo, useState } from "react";
import { REPORT_COLORS, bomItems, mccRows, type BomItem } from "@/core";
import { useShell } from "../AppShell";
import type { StoredLoadList } from "../loadListStore";

const BORDER = `1px solid ${REPORT_COLORS.border}`;
const th = { padding: "2px 6px", border: BORDER, background: REPORT_COLORS.header, color: "#000", fontWeight: "bold" as const, textAlign: "center" as const };
const td = { padding: "2px 6px", border: BORDER, textAlign: "center" as const, verticalAlign: "top" as const };
const left = { textAlign: "left" as const };

type Status = { kind: "idle" } | { kind: "busy" } | { kind: "done"; name: string; sheets: number } | { kind: "error"; message: string };

const stamp = () => new Date().toISOString().slice(0, 16).replace(/[-:]/g, "").replace("T", "-");

export default function BomPage() {
  const { assumptions, loadList: stored, pendingDownload, clearDownload } = useShell();
  const [status, setStatus] = useState<Status>({ kind: "idle" });

  const mcc = useMemo(() => (stored ? mccRows(stored.sheet, assumptions).rows : null), [stored, assumptions]);
  const items = useMemo(() => (mcc ? bomItems(mcc) : null), [mcc]);
  // The current behind each item (from its first load): entered or calculated from kW, and the design current that selected the breaker.
  const currentOf = (i: BomItem) => mcc?.find((r) => r.rowNumber === i.rows[0])?.current;

  // Clicking the BOM tab in the sidebar downloads the BOM workbook of the (already calculated) items.
  useEffect(() => {
    if (pendingDownload !== "bom") return;
    clearDownload();
    if (stored && items && items.length) void download(stored, items);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingDownload]);

  async function download(list: StoredLoadList, its: BomItem[]) {
    setStatus({ kind: "busy" });
    try {
      // The template is a static file of the site. One retry (cache-busted) covers a dropped connection.
      const tplUrl = `${window.location.origin}/bom/unit_bom_template.xlsx`;
      let res: Response;
      try {
        res = await fetch(tplUrl);
      } catch {
        try {
          res = await fetch(`${tplUrl}?retry=${Date.now()}`, { cache: "no-store" });
        } catch {
          throw new Error(`the unit-BOM template could not be fetched from ${tplUrl} (network error or blocked by the browser or an extension). Reload the page and try again.`);
        }
      }
      if (!res.ok) throw new Error(`could not load the unit-BOM template (${res.status}) from ${tplUrl}`);
      const { buildBomWorkbook } = await import("@/core/bom-workbook");
      const out = await buildBomWorkbook(await res.arrayBuffer(), its, list.fileName);
      const name = `BOM_${stamp()}.xlsx`;
      const url = URL.createObjectURL(new Blob([out.buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10_000);
      setStatus({ kind: "done", name, sheets: out.sheets });
    } catch (e) {
      setStatus({ kind: "error", message: (e as Error).message });
    }
  }

  const actions = items?.filter((i) => i.action).length ?? 0;
  return (
    <div>
      <p>
        {stored ? <>Load list: <b>{stored.fileName}</b>. </> : <>No load list yet: add one at the top of the sidebar. </>}
        <button type="button" disabled={!stored || !items || status.kind === "busy"} onClick={() => stored && items && void download(stored, items)}>
          Download the BOM workbook
        </button>
      </p>
      <div aria-live="polite">
        {status.kind === "busy" && <p>Preparing the BOM workbook…</p>}
        {status.kind === "done" && <p>✔ Downloaded <b>{status.name}</b> ({status.sheets} unit-BOM sheets plus an Index). A browser cannot launch Excel by itself: open the file from your downloads list.</p>}
        {status.kind === "error" && <p role="alert" style={{ color: "crimson" }}>Could not create the BOM workbook: {status.message}</p>}
      </div>

      {items && (
        <section>
          <h3>Unique items (type + rating)</h3>
          <p>
            <b>{items.length}</b> unique items · {items.filter((i) => i.sheets.length).length} with a unit-BOM sheet ·{" "}
            <span style={{ background: actions ? "#FF0000" : "transparent", color: actions ? "#fff" : undefined, padding: "0 6px" }}>{actions} action items</span>
          </p>
          <p style={{ fontSize: 14 }}>
            <b>Class</b> is empty on purpose: its logic will be given later (reminder). Red = action item, to be reviewed. A plain 6 or 10 A MPCB / MCB feeder needs no BOM sheet.
            The sheet layouts of the unit BOMs are not changed: only the rating text is set.
          </p>
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse" }}>
              <thead><tr>
                <th style={th}>SL</th><th style={th}>Type</th><th style={th}>Breaker / rating</th><th style={th}>Class</th>
                <th style={th}>Motor kW</th><th style={th}>Current used (A)</th><th style={th}>Calculated from kW (A)</th><th style={th}>Design current (A)</th><th style={th}>Quantity</th><th style={th}>Unit-BOM sheet(s)</th><th style={th}>Action</th>
              </tr></thead>
              <tbody>
                {items.map((i, n) => (
                  <tr key={i.key}>
                    <td style={td}>{n + 1}</td><td style={{ ...td, ...left }}>{i.type}</td>
                    <td style={{ ...td, whiteSpace: "nowrap" }}>{i.breaker || "(review)"}</td><td style={td} />
                    <td style={td}>{i.kw ?? ""}</td>
                    <td style={td} title={currentOf(i)?.source}>{currentOf(i)?.usedA === undefined ? "" : Math.round(currentOf(i)!.usedA!)}</td>
                    <td style={td}>{currentOf(i)?.calculatedA === undefined ? "" : Math.round(currentOf(i)!.calculatedA!)}</td>
                    <td style={td}>{currentOf(i)?.designA === undefined ? "" : Math.round(currentOf(i)!.designA!)}</td><td style={td}>{i.count}</td>
                    <td style={{ ...td, ...left, minWidth: 260 }}>{i.noSheet ? "No BOM sheet required" : i.sheets.map((s) => s.trim()).join("; ") || "(none)"}</td>
                    <td style={{ ...td, ...left, minWidth: 320, fontSize: 13, ...(i.action ? { background: "#FF0000", color: "#fff", fontWeight: "bold" } : {}) }}>{i.action ?? ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p style={{ fontSize: 14 }}>
            Bus couplers are not in a load list; they follow the transformer sets and will be linked later. Selections are preliminary and need engineering approval.
          </p>
        </section>
      )}
    </div>
  );
}
