import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";

const WORKER_URL = "https://listing-expander-api.alibaba-amy-h.workers.dev";

async function callLLM(prompt, imageBase64, imageType, maxTokens) {
  const res = await fetch(WORKER_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, imageBase64, imageType, maxTokens: maxTokens || 8000 }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message || "API 錯誤");
  const text = (data.content || []).map((x) => x.text || "").join("");
  let clean = text.replace(/```json[\s\S]*?```|```[\s\S]*?```/g, (m) => m.replace(/```json|```/g, "")).trim();
  clean = clean.replace(/^```json|^```|```$/g, "").trim();

  // 找到第一個 [ 開始
  const start = clean.indexOf("[");
  if (start === -1) throw new Error("AI 回傳格式錯誤，請重試");
  clean = clean.slice(start);

  // 嘗試直接 parse
  try { return clean; } catch (_) {}

  // 截斷修復：找最後一個完整的 }
  const lastBrace = clean.lastIndexOf("}");
  if (lastBrace > 0) clean = clean.slice(0, lastBrace + 1) + "]";
  else clean = clean + '"}]';

  return clean;
}

const STYLE = {
  app: { fontFamily: "sans-serif", maxWidth: 900, margin: "0 auto", padding: 24, color: "#222" },
  h1: { color: "#e8650a", marginBottom: 4 },
  sub: { color: "#666", marginBottom: 24 },
  card: { background: "#fff", border: "1px solid #ddd", borderRadius: 8, padding: 20, marginBottom: 16 },
  label: { display: "block", fontWeight: 600, marginBottom: 6 },
  input: { width: "100%", padding: "8px 12px", border: "1px solid #ccc", borderRadius: 6, fontSize: 14, boxSizing: "border-box" },
  textarea: { width: "100%", padding: "8px 12px", border: "1px solid #ccc", borderRadius: 6, fontSize: 14, boxSizing: "border-box", resize: "vertical" },
  btn: { padding: "10px 24px", borderRadius: 6, border: "none", fontWeight: 600, fontSize: 15, cursor: "pointer" },
  btnOrange: { background: "#e8650a", color: "#fff" },
  btnGray: { background: "#f0f0f0", color: "#555" },
  btnGreen: { background: "#27ae60", color: "#fff" },
  tag: { display: "inline-block", background: "#fff3e0", color: "#e8650a", borderRadius: 4, padding: "2px 8px", fontSize: 12, margin: "2px" },
  resultCard: { background: "#fafafa", border: "1px solid #e0e0e0", borderRadius: 8, padding: 16, marginBottom: 12 },
  num: { background: "#e8650a", color: "#fff", borderRadius: "50%", width: 24, height: 24, display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, marginRight: 8 },
};

export default function App() {
  const [productName, setProductName] = useState("");
  const [productDesc, setProductDesc] = useState("");
  const [imageFile, setImageFile] = useState(null);
  const [imageBase64, setImageBase64] = useState(null);
  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef();

  function handleImage(e) {
    const file = e.target.files[0];
    if (!file) return;
    setImageFile(file);
    const reader = new FileReader();
    reader.onload = (ev) => {
      const b64 = ev.target.result.split(",")[1];
      setImageBase64(b64);
    };
    reader.readAsDataURL(file);
  }

  async function handleExpand() {
    if (!productName.trim()) { setError("請輸入商品名稱"); return; }
    setError("");
    setLoading(true);
    setResults([]);

    const prompt = `You are an Alibaba.com listing expert. Generate exactly 10 differentiated product listings in JSON array format.

Product: ${productName}
${productDesc ? `Details: ${productDesc}` : ""}

Rules:
- Each listing targets a different buyer segment or use case
- Title: 50-120 chars with keywords
- Keywords: array of 5 English search terms
- Description: 80-150 chars
- Category: suggested Alibaba category
- Selling_point: one key differentiator

Return ONLY a JSON array, no markdown, no explanation:
[{"title":"...","keywords":["k1","k2","k3","k4","k5"],"description":"...","category":"...","selling_point":"..."},...]`;

    try {
      const raw = await callLLM(prompt, imageBase64, imageFile?.type, 4000);
      const parsed = JSON.parse(raw);
      setResults(parsed);
    } catch (err) {
      setError("生成失敗：" + err.message);
    } finally {
      setLoading(false);
    }
  }

  function exportExcel() {
    const rows = results.map((r, i) => ({
      "編號": i + 1,
      "Title（標題）": r.title,
      "Keywords（關鍵字）": (r.keywords || []).join(", "),
      "Description（描述）": r.description,
      "Category（分類）": r.category,
      "Selling Point（賣點）": r.selling_point,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 6 }, { wch: 60 }, { wch: 40 }, { wch: 80 }, { wch: 30 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Listings");
    XLSX.writeFile(wb, `${productName}_listings.xlsx`);
  }

  return (
    <div style={STYLE.app}>
      <h1 style={STYLE.h1}>阿里巴巴 Listing 擴張器</h1>
      <p style={STYLE.sub}>輸入 1 個商品 → AI 自動生成 10 個差異化 listing，擴大曝光觸及更多買家</p>

      <div style={STYLE.card}>
        <label style={STYLE.label}>商品名稱 *</label>
        <input
          style={STYLE.input}
          placeholder="例：Stainless Steel Water Bottle"
          value={productName}
          onChange={(e) => setProductName(e.target.value)}
        />

        <label style={{ ...STYLE.label, marginTop: 16 }}>商品描述（選填，越詳細越準確）</label>
        <textarea
          style={{ ...STYLE.textarea, minHeight: 80 }}
          placeholder="例：500ml 雙層保溫，適合戶外運動，BPA-free，OEM 可接受..."
          value={productDesc}
          onChange={(e) => setProductDesc(e.target.value)}
        />

        <label style={{ ...STYLE.label, marginTop: 16 }}>商品圖片（選填）</label>
        <input ref={fileRef} type="file" accept="image/*" onChange={handleImage} style={{ display: "none" }} />
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <button style={{ ...STYLE.btn, ...STYLE.btnGray }} onClick={() => fileRef.current.click()}>
            選擇圖片
          </button>
          {imageFile && <span style={{ color: "#27ae60", fontSize: 14 }}>✓ {imageFile.name}</span>}
        </div>

        {error && <div style={{ color: "#c0392b", marginTop: 12, fontSize: 14 }}>{error}</div>}

        <div style={{ marginTop: 20 }}>
          <button
            style={{ ...STYLE.btn, ...STYLE.btnOrange, opacity: loading ? 0.7 : 1 }}
            onClick={handleExpand}
            disabled={loading}
          >
            {loading ? "AI 生成中..." : "生成 10 個 Listing"}
          </button>
        </div>
      </div>

      {results.length > 0 && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <h2 style={{ margin: 0 }}>生成結果（{results.length} 個）</h2>
            <button style={{ ...STYLE.btn, ...STYLE.btnGreen }} onClick={exportExcel}>
              下載 Excel
            </button>
          </div>

          {results.map((r, i) => (
            <div key={i} style={STYLE.resultCard}>
              <div style={{ display: "flex", alignItems: "flex-start", marginBottom: 8 }}>
                <span style={STYLE.num}>{i + 1}</span>
                <strong style={{ fontSize: 15, lineHeight: "24px" }}>{r.title}</strong>
              </div>

              <div style={{ marginBottom: 8 }}>
                {(r.keywords || []).map((kw, j) => (
                  <span key={j} style={STYLE.tag}>{kw}</span>
                ))}
              </div>

              <p style={{ margin: "8px 0", fontSize: 14, color: "#444", lineHeight: 1.6 }}>{r.description}</p>

              <div style={{ display: "flex", gap: 16, fontSize: 13, color: "#666", marginTop: 8 }}>
                <span>📁 {r.category}</span>
                <span>⭐ {r.selling_point}</span>
              </div>
            </div>
          ))}
        </>
      )}
    </div>
  );
}
