import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const FIELDS = [
  { key: "car_name",        label: "車名" },
  { key: "grade",           label: "グレード" },
  { key: "year",            label: "年式" },
  { key: "month",           label: "月" },
  { key: "color",           label: "色" },
  { key: "chassis_number",  label: "車台番号" },
  { key: "model_code",      label: "型式" },
  { key: "engine_cc",       label: "排気量(cc)" },
  { key: "fuel",            label: "燃料" },
  { key: "shift",           label: "シフト" },
  { key: "doors",           label: "ドア数" },
  { key: "mileage",         label: "走行距離(km)" },
  { key: "score",           label: "評価点" },
  { key: "interior_score",  label: "内装評価" },
  { key: "auction_house",   label: "オークション会場" },
  { key: "lot_number",      label: "出品番号" },
  { key: "purchase_price",  label: "落札価格(円)" },
  { key: "tax",             label: "消費税(円)" },
  { key: "self_tax",        label: "自動車税(円)" },
  { key: "recycle_fee",     label: "リサイクル料(円)" },
  { key: "auction_fee",     label: "落札手数料(円)" },
  { key: "navi",            label: "ナビ" },
  { key: "tv",              label: "TV" },
  { key: "sr",              label: "SR" },
  { key: "leather",         label: "革シート" },
  { key: "etc",             label: "ETC" },
  { key: "condition_notes", label: "状態・傷メモ", wide: true },
];

const SYSTEM_PROMPT = `日本の中古車オークションシートを読み取り、JSONのみ返してください。説明文やMarkdownは一切不要です。
{
  "car_name": "車名（メーカー含む 例: スズキ クロスビー）",
  "grade": "グレード",
  "year": "年式（西暦4桁 例: 2021）",
  "month": "初度登録月（数字のみ 例: 7）",
  "color": "色",
  "chassis_number": "車台番号（英数字）",
  "model_code": "型式",
  "engine_cc": "排気量（数字のみ）",
  "fuel": "燃料種別（ガソリン/ディーゼル/ハイブリッド/電気）",
  "shift": "シフト（AT/MT/CVT/DAT等）",
  "doors": "ドア数（数字のみ）",
  "mileage": "走行距離（数字のみ）",
  "score": "評価点（数字 例: 4.5）",
  "interior_score": "内装評価（A/B/C等）",
  "auction_house": "会場名",
  "lot_number": "出品番号",
  "purchase_price": "落札価格（数字のみ）",
  "tax": "消費税（数字のみ）",
  "self_tax": "自動車税（数字のみ）",
  "recycle_fee": "リサイクル料（数字のみ）",
  "auction_fee": "落札手数料（数字のみ）",
  "navi": "有/無",
  "tv": "有/無",
  "sr": "有/無",
  "leather": "有/無",
  "etc": "有/無",
  "condition_notes": "傷・汚れ・修復歴など（カンマ区切り）"
}`;

async function readOneImage(img) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1500,
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content: [
        { type: "image", source: { type: "base64", media_type: img.mimeType, data: img.base64 } },
        { type: "text", text: "読み取ってJSONのみ返してください。" }
      ]}]
    })
  });
  const data = await res.json();
  const text = (data.content || []).map(c => c.text || "").join("");
  return JSON.parse(text.replace(/```json|```/g, "").trim());
}

function EditableCard({ result, index, onChange, onDelete }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState({ ...result });
  const save = () => { onChange(index, draft); setEditing(false); };
  const cancel = () => { setDraft({ ...result }); setEditing(false); };

  return (
    <div style={{ background: "white", borderRadius: 12, overflow: "hidden", marginBottom: 14, boxShadow: "0 1px 6px rgba(0,0,0,0.09)" }}>
      <div style={{ background: result._ok ? "#1a2744" : "#dc2626", color: "white", padding: "10px 14px", display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
        {result._preview && <img src={result._preview} alt="" style={{ width: 48, height: 36, objectFit: "cover", borderRadius: 4, flexShrink: 0 }} />}
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 700, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            {result._ok ? `${result.car_name || "車名不明"} ${result.grade || ""}` : "❌ 読み取りエラー"}
          </div>
          <div style={{ fontSize: 11, opacity: 0.7, marginTop: 1 }}>
            {result._ok ? `${result.auction_house || ""}${result.lot_number ? " #" + result.lot_number : ""}` : result._err || ""}
          </div>
        </div>
        {result._ok && result.score && <span style={{ background: "#f59e0b", borderRadius: 6, padding: "3px 8px", fontWeight: 700, fontSize: 12 }}>評価 {result.score}</span>}
        {result._ok && !editing && (
          <button onClick={() => { setDraft({ ...result }); setEditing(true); }}
            style={{ background: "rgba(255,255,255,0.2)", color: "white", border: "none", borderRadius: 6, padding: "5px 10px", cursor: "pointer", fontSize: 12 }}>✏️ 編集</button>
        )}
        {editing && <>
          <button onClick={save} style={{ background: "#22c55e", color: "white", border: "none", borderRadius: 6, padding: "5px 10px", cursor: "pointer", fontSize: 12, fontWeight: 700 }}>✓ 保存</button>
          <button onClick={cancel} style={{ background: "rgba(255,255,255,0.15)", color: "white", border: "none", borderRadius: 6, padding: "5px 8px", cursor: "pointer", fontSize: 12 }}>取消</button>
        </>}
        <button onClick={() => onDelete(index)}
          style={{ background: "rgba(255,255,255,0.12)", color: "white", border: "none", borderRadius: 6, padding: "5px 8px", cursor: "pointer", fontSize: 12 }}>削除</button>
      </div>
      {result._ok && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(155px, 1fr))" }}>
          {FIELDS.map(({ key, label, wide }) => (
            <div key={key} style={{ padding: "9px 13px", borderRight: "1px solid #f1f5f9", borderBottom: "1px solid #f1f5f9", gridColumn: wide ? "1 / -1" : undefined }}>
              <div style={{ fontSize: 10, color: "#94a3b8", fontWeight: 600, marginBottom: 3 }}>{label}</div>
              {editing ? (
                <input value={draft[key] || ""} onChange={e => setDraft(d => ({ ...d, [key]: e.target.value }))}
                  style={{ width: "100%", border: "1px solid #93c5fd", borderRadius: 5, padding: "4px 7px", fontSize: 13, fontFamily: "inherit", outline: "none", background: "white" }} />
              ) : (
                <div style={{ fontSize: 13, color: result[key] ? "#1e293b" : "#d1d5db", fontWeight: result[key] ? 500 : 400, minHeight: 20 }}>
                  {result[key] || "—"}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [xlsxData, setXlsxData] = useState(null);
  const [images, setImages] = useState([]);
  const [results, setResults] = useState([]);
  const [loadingSet, setLoadingSet] = useState(new Set());
  const xlsxRef = useRef();
  const imgRef = useRef();

  const handleXlsx = useCallback((files) => {
    const f = Array.from(files).find(f => f.name.match(/\.xlsx?$/i));
    if (!f) return;
    const reader = new FileReader();
    reader.onload = e => {
      const buffer = e.target.result;
      const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
      const sheetName = wb.SheetNames.includes("27期") ? "27期" : wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const range = XLSX.utils.decode_range(ws["!ref"]);
      let lastNum = 0, lastRow = 1;
      for (let r = range.e.r; r >= 0; r--) {
        const cell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
        if (cell && typeof cell.v === "string" && /^R\d+-\d+$/i.test(cell.v)) {
          lastNum = parseInt(cell.v.split("-")[1]);
          lastRow = r + 1;
          break;
        }
      }
      setXlsxData({ lastNum, lastRow, name: f.name });
      setResults([]);
    };
    reader.readAsArrayBuffer(f);
  }, []);

  const handleFiles = useCallback((files) => {
    const valid = Array.from(files).filter(f => f.type.startsWith("image/"));
    if (!valid.length) return;
    Promise.all(valid.map(file => new Promise(res => {
      const reader = new FileReader();
      reader.onload = e => {
        const dataUrl = e.target.result;
        res({ name: file.name, base64: dataUrl.split(",")[1], preview: dataUrl, mimeType: dataUrl.split(";")[0].replace("data:", "") });
      };
      reader.readAsDataURL(file);
    }))).then(imgs => setImages(prev => [...prev, ...imgs]));
  }, []);

  const analyze = async () => {
    if (!images.length) return;
    const toAnalyze = [...images];
    setImages([]);
    setLoadingSet(new Set(toAnalyze.map((_, i) => i)));
    await Promise.all(toAnalyze.map(async (img, i) => {
      try {
        const parsed = await readOneImage(img);
        setResults(prev => [...prev, { ...parsed, _preview: img.preview, _ok: true }]);
      } catch(e) {
        setResults(prev => [...prev, { _preview: img.preview, _name: img.name, _ok: false, _err: e.message }]);
      }
      setLoadingSet(prev => { const s = new Set(prev); s.delete(i); return s; });
    }));
  };

  const handleDownload = () => {
    const valid = results.filter(r => r._ok && (r.car_name || r.chassis_number));
    if (!valid.length) return;

    const MAX_COL = 52;
    const colLabels = [];
    for (let i = 1; i <= MAX_COL; i++) {
      let s = "", n = i;
      while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); }
      colLabels.push(s);
    }

    const ci = letter => {
      let n = 0;
      for (const c of letter.toUpperCase()) n = n * 26 + c.charCodeAt(0) - 64;
      return n - 1;
    };

    const esc = v => {
      const s = String(v ?? "");
      return s.includes(",") || s.includes('"') || s.includes("\n") ? `"${s.replace(/"/g, '""')}"` : s;
    };

    let num = (xlsxData?.lastNum ?? 0) + 1;
    const rows = [];

    for (const r of valid) {
      const id = `R27-${String(num).padStart(3, "0")}`;
      const carName = [r.car_name, r.grade].filter(Boolean).join("　");
      const yearMonth = r.year ? `${r.year}${r.month ? "/" + r.month : ""}` : "";

      const row1 = Array(MAX_COL).fill("");
      row1[ci("A")]  = id;
      row1[ci("D")]  = yearMonth;
      row1[ci("F")]  = carName;
      row1[ci("J")]  = r.score ? (parseFloat(r.score) || r.score) : "";
      row1[ci("K")]  = r.purchase_price ? parseInt(r.purchase_price) : "";
      row1[ci("M")]  = r.tax ? parseInt(r.tax) : "";
      row1[ci("O")]  = r.self_tax ? parseInt(r.self_tax) : "";
      row1[ci("Q")]  = r.recycle_fee ? parseInt(r.recycle_fee) : "";
      row1[ci("S")]  = r.auction_fee ? parseInt(r.auction_fee) : "";
      row1[ci("AC")] = "予定";

      const row2 = Array(MAX_COL).fill("");
      row2[ci("B")]  = r.auction_house || "";
      row2[ci("C")]  = r.lot_number || "";
      row2[ci("D")]  = r.color || "";
      row2[ci("F")]  = r.chassis_number || "";
      row2[ci("I")]  = r.mileage ? parseInt(r.mileage) : "";
      row2[ci("AC")] = "着";

      rows.push(row1);
      rows.push(row2);
      num++;
    }

    const csv = "\uFEFF" + colLabels.join(",") + "\n" + rows.map(r => r.map(esc).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `追記データ_R27-${String((xlsxData?.lastNum ?? 0) + 1).padStart(3, "0")}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const isLoading = loadingSet.size > 0;
  const okResults = results.filter(r => r._ok);
  const nextNum = (xlsxData?.lastNum ?? 0) + 1;
  const canDownload = !isLoading && okResults.length > 0;

  return (
    <div style={{ fontFamily: "sans-serif", background: "#f0f4f8", minHeight: "100vh", fontSize: 14 }}>
      <div style={{ background: "#1a2744", color: "white", padding: "14px 20px", display: "flex", alignItems: "center", gap: 10, position: "sticky", top: 0, zIndex: 10, boxShadow: "0 2px 8px rgba(0,0,0,0.2)" }}>
        <span style={{ fontSize: 22 }}>🚗</span>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 16 }}>在庫帳 オークションシート読み取り</div>
          <div style={{ fontSize: 11, opacity: 0.6 }}>読み取り → 編集 → CSVでDL</div>
        </div>
        {isLoading && <span style={{ background: "#f59e0b", borderRadius: 20, padding: "4px 12px", fontSize: 12, fontWeight: 700 }}>⏳ {loadingSet.size}件処理中</span>}
        {canDownload && (
          <button onClick={handleDownload} style={{ background: "#22c55e", color: "white", border: "none", borderRadius: 8, padding: "8px 14px", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>
            ⬇ CSVでDL
          </button>
        )}
      </div>

      <div style={{ padding: 20, maxWidth: 860, margin: "0 auto" }}>
        <div style={{ background: "white", borderRadius: 12, padding: 16, marginBottom: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)", border: `2px solid ${xlsxData ? "#22c55e" : "#e2e8f0"}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ background: xlsxData ? "#22c55e" : "#94a3b8", color: "white", borderRadius: "50%", width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12 }}>
              {xlsxData ? "✓" : "A"}
            </div>
            <div style={{ fontWeight: 700, color: "#1a2744", flex: 1 }}>在庫帳Excel（連番検出用・どちらが先でもOK）</div>
            {xlsxData && <button onClick={() => setXlsxData(null)} style={{ background: "none", border: "1px solid #e2e8f0", color: "#64748b", borderRadius: 6, padding: "3px 10px", cursor: "pointer", fontSize: 12 }}>変更</button>}
          </div>
          {xlsxData ? (
            <div style={{ marginTop: 10, background: "#f0fdf4", borderRadius: 8, padding: "10px 14px" }}>
              <div style={{ fontWeight: 700, color: "#166534", fontSize: 13 }}>{xlsxData.name}</div>
              <div style={{ fontSize: 12, color: "#16a34a", marginTop: 1 }}>最終: R27-{String(xlsxData.lastNum).padStart(3,"0")} → 次回: R27-{String(nextNum).padStart(3,"0")} から</div>
            </div>
          ) : (
            <div onClick={() => xlsxRef.current.click()} style={{ marginTop: 12, border: "2px dashed #cbd5e1", borderRadius: 10, padding: 18, textAlign: "center", cursor: "pointer", background: "#f8fafc" }}>
              <input ref={xlsxRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => { handleXlsx(e.target.files); e.target.value = ""; }} />
              <div style={{ fontSize: 26, marginBottom: 4 }}>📊</div>
              <div style={{ fontWeight: 700, color: "#1a2744", fontSize: 13, marginBottom: 2 }}>在庫帳.xlsx を選択</div>
              <div style={{ fontSize: 11, color: "#94a3b8" }}>「27期」シートから連番を自動検出</div>
            </div>
          )}
        </div>

        <div style={{ background: "white", borderRadius: 12, padding: 16, marginBottom: 16, boxShadow: "0 1px 4px rgba(0,0,0,0.08)", border: "2px solid #93c5fd" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
            <div style={{ background: okResults.length > 0 ? "#22c55e" : "#1a2744", color: "white", borderRadius: "50%", width: 26, height: 26, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12 }}>
              {okResults.length > 0 ? "✓" : "B"}
            </div>
            <div style={{ fontWeight: 700, color: "#1a2744" }}>オークションシートの写真（どちらが先でもOK）</div>
          </div>
          <div onClick={() => imgRef.current.click()} onDrop={e => { e.preventDefault(); handleFiles(e.dataTransfer.files); }} onDragOver={e => e.preventDefault()}
            style={{ border: "2px dashed #93c5fd", borderRadius: 10, padding: 20, textAlign: "center", cursor: "pointer", background: "#f8fafc" }}>
            <input ref={imgRef} type="file" accept="image/*" multiple style={{ display: "none" }} onChange={e => { handleFiles(e.target.files); e.target.value = ""; }} />
            <div style={{ fontSize: 28, marginBottom: 6 }}>📸</div>
            <div style={{ fontWeight: 700, color: "#1a2744", fontSize: 13, marginBottom: 2 }}>写真を選択（複数可）</div>
            <div style={{ fontSize: 11, color: "#94a3b8" }}>JPG・PNG / 複数は並列処理</div>
          </div>
          {images.length > 0 && (
            <>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(90px, 1fr))", gap: 8, marginTop: 12 }}>
                {images.map((img, i) => (
                  <div key={i} style={{ position: "relative", borderRadius: 8, overflow: "hidden", background: "#e2e8f0" }}>
                    <img src={img.preview} alt="" style={{ width: "100%", height: 70, objectFit: "cover", display: "block" }} />
                    <div style={{ padding: "3px 5px", fontSize: 9, color: "#64748b", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{img.name}</div>
                    <button onClick={() => setImages(p => p.filter((_, j) => j !== i))}
                      style={{ position: "absolute", top: 3, right: 3, background: "rgba(0,0,0,0.55)", color: "white", border: "none", borderRadius: "50%", width: 18, height: 18, cursor: "pointer", fontSize: 10, lineHeight: "18px", textAlign: "center" }}>✕</button>
                  </div>
                ))}
              </div>
              <button onClick={analyze} style={{ background: "#1a2744", color: "white", border: "none", borderRadius: 10, padding: 13, fontWeight: 700, fontSize: 14, cursor: "pointer", width: "100%", marginTop: 12, fontFamily: "inherit" }}>
                🔍 {images.length}枚を並列読み取り開始
              </button>
            </>
          )}
        </div>

        {isLoading && (
          <div style={{ background: "white", borderRadius: 12, padding: "14px 18px", marginBottom: 14, display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ fontSize: 22 }}>⏳</div>
            <div>
              <div style={{ fontWeight: 700, color: "#1a2744", fontSize: 13 }}>AI読み取り中...</div>
              <div style={{ fontSize: 12, color: "#64748b" }}>残り {loadingSet.size}枚 / 完了したものから順に表示</div>
            </div>
          </div>
        )}

        {canDownload && (
          <div style={{ background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 10, padding: "12px 16px", marginBottom: 14, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <span>📋</span>
            <div style={{ flex: 1, fontSize: 13, color: "#166534" }}><strong>{okResults.length}台</strong>分のCSVをダウンロードできます（R27-{String(nextNum).padStart(3,"0")}〜）</div>
            <button onClick={handleDownload} style={{ background: "#22c55e", color: "white", border: "none", borderRadius: 8, padding: "8px 16px", fontWeight: 700, cursor: "pointer", fontSize: 13 }}>⬇ CSVでDL</button>
          </div>
        )}

        {results.length === 0 && images.length === 0 && !isLoading && (
          <div style={{ textAlign: "center", padding: "30px 20px", color: "#94a3b8" }}>
            <div style={{ fontSize: 36, marginBottom: 8 }}>📋</div>
            <div>写真をアップロードして読み取り開始</div>
          </div>
        )}

        {results.map((r, i) => (
          <EditableCard key={i} result={r} index={i}
            onChange={(idx, newData) => setResults(p => p.map((x, j) => j === idx ? { ...newData, _preview: x._preview, _ok: x._ok } : x))}
            onDelete={(idx) => setResults(p => p.filter((_, j) => j !== idx))}
          />
        ))}

        {canDownload && (
          <button onClick={handleDownload} style={{ background: "#22c55e", color: "white", border: "none", borderRadius: 10, padding: 14, fontWeight: 700, fontSize: 15, cursor: "pointer", width: "100%", marginTop: 4, fontFamily: "inherit" }}>
            ⬇ {okResults.length}台分をCSVでDL（R27-{String(nextNum).padStart(3,"0")}〜）
          </button>
        )}
      </div>
    </div>
  );
}
