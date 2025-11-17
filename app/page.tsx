"use client";

import { useMemo, useState } from "react";

export default function HomePage() {
  const [file, setFile] = useState<File | null>(null);
  const [instruction, setInstruction] = useState("");
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<string>("");

  const canRun = useMemo(() => !!file && instruction.trim().length > 0 && !loading, [file, instruction, loading]);

  async function onRun() {
    if (!file) return;
    setLoading(true);
    setMessage("");
    try {
      const form = new FormData();
      form.append("file", file);
      form.append("instruction", instruction);
      const res = await fetch("/api/transform", { method: "POST", body: form });
      if (!res.ok) {
        const txt = await res.text();
        throw new Error(txt || `HTTP ${res.status}`);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const outName = (file.name.replace(/\.xlsx$/i, "") || "output") + "_updated.xlsx";
      a.download = outName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setMessage("? File updated. Download started.");
    } catch (e: any) {
      setMessage(`? ${e?.message || "Failed"}`);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="container">
      <div className="card">
        <div className="badge">Excel AI Agent</div>
        <h1 className="h1">Excel ko AI se update karein</h1>
        <p className="sub">Upload Excel (.xlsx) aur natural language mein instruction dein (Hindi/English). Agent aapke hisaab se file update karega.</p>

        <div className="row" style={{ marginBottom: 12 }}>
          <input className="input" type="file" accept=".xlsx" onChange={(e) => setFile(e.target.files?.[0] || null)} />
        </div>

        <div className="row" style={{ marginBottom: 12 }}>
          <textarea
            className="input"
            rows={5}
            placeholder={
              "Udaharan: 'Sheet1 me Total column add karo jo Price + Tax ho' ya 'Sales sheet ka naam Revenue rakho' ya 'Quantity > 10 wali rows rakhna, baaki hatao'"
            }
            value={instruction}
            onChange={(e) => setInstruction(e.target.value)}
          />
        </div>

        <div className="row" style={{ alignItems: "center" }}>
          <button className="button" disabled={!canRun} onClick={onRun}>{loading ? "Processing..." : "Run Agent"}</button>
          <div className="hint">Supported: rename sheet/column, add column (sum), delete column, filter rows, sort, set values</div>
        </div>

        {message && (
          <p className={message.startsWith("?") ? "success" : "error"} style={{ marginTop: 12 }}>{message}</p>
        )}

        <div className="footer">Data local hi process hota hai server function me; file download ke liye turant ready.</div>
      </div>
    </div>
  );
}
