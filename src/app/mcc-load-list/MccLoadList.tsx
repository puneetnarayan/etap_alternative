"use client";

import { useEffect, useRef, useState } from "react";
import { readSheet } from "read-excel-file/universal";
import { loadLoadList } from "../loadListStore";
import { useShell } from "../AppShell";

export default function MccLoadList() {
  const { status, generate } = useShell();
  const [dragging, setDragging] = useState(false);
  const [error, setError] = useState("");
  const input = useRef<HTMLInputElement>(null);
  // Read after mount (and after each generate): the server render has no browser storage.
  const [stored, setStored] = useState<{ fileName: string } | null>(null);
  useEffect(() => { setStored(loadLoadList()); }, [status]);

  async function handleFile(file: File | undefined) {
    if (!file) return;
    setError("");
    if (!/\.xlsx$/i.test(file.name)) {
      setError("Please use an Excel .xlsx file.");
      return;
    }
    try {
      const sheet = await readSheet(await file.arrayBuffer());
      await generate({ fileName: file.name, sheet });
    } catch (e) {
      setError(`Could not read that file: ${(e as Error).message}`);
    }
  }

  return (
    <div>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); void handleFile(e.dataTransfer.files[0]); }}
        onClick={() => input.current?.click()}
        style={{ border: "2px dashed #888", borderRadius: 8, padding: "28px 16px", textAlign: "center", cursor: "pointer", background: dragging ? "#C6EFCE" : "transparent" }}
      >
        Drag and drop a load list (.xlsx) here, or click to choose a file. The workbook is created and downloaded straight away.
        <input ref={input} type="file" accept=".xlsx" hidden onChange={(e) => { void handleFile(e.target.files?.[0]); e.target.value = ""; }} />
      </div>
      {error && <p role="alert" style={{ color: "#c0392b" }}>{error}</p>}
      <p>
        {stored ? <>Current load list: <b>{stored.fileName}</b> (the last one used in this tab). </> : <>No load list has been used in this tab yet. </>}
        <button type="button" disabled={!stored || status.kind === "busy"} onClick={() => void generate()}>
          Generate the workbook again
        </button>
      </p>
      <div aria-live="polite">
        {status.kind === "busy" && <p>Generating the workbook…</p>}
        {status.kind === "done" && (
          <p>
            ✔ Downloaded <b>{status.name}</b> ({status.rows} loads{status.missing > 0 ? `, ${status.missing} without enough data for a current` : ""}).
            A browser cannot launch Excel by itself: open the downloaded file from your downloads list.
          </p>
        )}
        {status.kind === "no-list" && <p>No load list yet: drop one above, or drop one on the Transformer sizing page first.</p>}
        {status.kind === "error" && <p role="alert" style={{ color: "#c0392b" }}>Could not generate the workbook: {status.message}</p>}
      </div>
    </div>
  );
}
