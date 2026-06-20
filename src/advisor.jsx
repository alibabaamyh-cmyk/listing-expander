import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import ReactDOM from "react-dom/client";

const WORKER_URL = "https://listing-expander-api.alibaba-amy-h.workers.dev";

async function callLLM(prompt, fileBase64, fileMime, maxTokens) {
  var res;
  try {
    res = await fetch(WORKER_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt, imageBase64: fileBase64, imageType: fileMime, maxTokens: maxTokens || 4000 })
    });
  } catch (e) {
    throw new Error("網路連線失敗，請重新整理後再試（" + e.message + "）");
  }
  if (!res.ok) throw new Error("伺服器錯誤 " + res.status);
  var data = await res.json();
  if (data.error) throw new Error(data.error.message || "API 錯誤");
  var text = (data.content || []).map(function(x) { return x.text || ""; }).join("");
  var clean = text.replace(/```json\s*|```/g, "").trim();
  var fb = clean.indexOf("{"), fbr = clean.indexOf("[");
  var start = fb === -1 ? fbr : fbr === -1 ? fb : Math.min(fb, fbr);
  if (start > 0) clean = clean.slice(start);
  try { JSON.parse(clean); return clean; } catch (_) {}
  if (clean.startsWith("[")) {
    var depth = 0, inStr = false, esc = false, lastEnd = -1;
    for (var i = 0; i < clean.length; i++) {
      var c = clean[i];
      if (esc) { esc = false; continue; }
      if (c === "\\" && inStr) { esc = true; continue; }
      if (c === '"') { inStr = !inStr; continue; }
      if (inStr) continue;
      if (c === "{" || c === "[") depth++;
      if (c === "}" || c === "]") { depth--; if (depth === 1 && c === "}") lastEnd = i; }
    }
    return lastEnd > 0 ? clean.slice(0, lastEnd + 1) + "]" : "[]";
  }
  return clean;
}

// ─── 顏色系統 ────────────────────────────────────────────
const C = {
  bg: "#F8FAFC",
  navy: "#1B2A4A",
  orange: "#FF6B35",
  green: "#16A34A",
  card: "#fff",
  border: "#E5E7EB",
  text: "#1e293b",
  muted: "#64748b",
  inp: { width: "100%", padding: "10px 13px", borderRadius: 9, border: "1.5px solid #E5E7EB", fontSize: 14, outline: "none", boxSizing: "border-box", color: "#1e293b", fontFamily: "inherit", background: "#fff" }
};

function Card({ children, style }) {
  return <div style={{ background: C.card, borderRadius: 14, padding: 24, boxShadow: "0 1px 6px rgba(0,0,0,0.07)", marginBottom: 20, ...style }}>{children}</div>;
}

function SectionTitle({ children }) {
  return <div style={{ fontWeight: 800, fontSize: 16, color: C.navy, marginBottom: 14, paddingBottom: 10, borderBottom: "2px solid " + C.orange }}>{children}</div>;
}

function Btn({ onClick, disabled, color, children, style }) {
  var bg = color === "orange" ? C.orange : color === "green" ? C.green : color === "ghost" ? "#F3F4F6" : C.navy;
  var fg = color === "ghost" ? C.text : "#fff";
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ padding: "11px 22px", background: disabled ? "#E5E7EB" : bg, color: disabled ? "#9CA3AF" : fg, border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer", ...style }}>
      {children}
    </button>
  );
}

// ════════════════════════════════════════════════════════
// MAIN ADVISOR APP
// ════════════════════════════════════════════════════════
export default function Advisor() {
  // ── Step 狀態 ──
  const [step, setStep] = useState("intake"); // intake | analyzing | result-many | result-few

  // ── Step 1 intake ──
  const [productRange, setProductRange] = useState(""); // "<10" | "10-50" | "50-200" | "200+"
  const [hasWebsite, setHasWebsite] = useState(null); // true | false | null
  const [htmlContent, setHtmlContent] = useState("");
  const [pdfBase64, setPdfBase64] = useState(null);
  const [pdfName, setPdfName] = useState("");
  const [loading, setLoading] = useState(false);
  const [loadingSecs, setLoadingSecs] = useState(0);
  const [error, setError] = useState("");
  const fileRef = useRef();

  // ── Step 2 analysis result ──
  const [analysis, setAnalysis] = useState(null);
  // analysis = { path: "many"|"few", summary, productCount, estimatedListings, products: [{name}], strategy: {dim1Label, dim2Label, dim3Label} }

  // ── Step 3b expansion ──
  const [products, setProducts] = useState([]); // [{name, dim1, dim2, dim3, dim1Options:[], dim2Options:[], dim3Options:[], listings:[], generating:false, done:false}]
  const [exportMsg, setExportMsg] = useState("");

  // 計時器
  useEffect(function () {
    if (!loading) { setLoadingSecs(0); return; }
    setLoadingSecs(0);
    var t = setInterval(function () { setLoadingSecs(function (s) { return s + 1; }); }, 1000);
    return function () { clearInterval(t); };
  }, [loading]);

  // ── PDF 上傳 ──
  function handlePdf(e) {
    var f = e.target.files[0];
    if (!f) return;
    setPdfName(f.name);
    var reader = new FileReader();
    reader.onload = function (ev) {
      var b64 = ev.target.result.split(",")[1];
      setPdfBase64(b64);
    };
    reader.readAsDataURL(f);
  }

  // ── Step 1 → AI 分析 ──
  async function analyze() {
    if (!productRange) { setError("請先選擇產品數量範圍"); return; }
    setError(""); setLoading(true); setStep("analyzing");

    var contextParts = [
      "你是阿里巴巴國際站電商顧問，請分析以下商家資訊並給出發品策略建議。",
      "\n商家產品數量範圍：" + productRange,
      hasWebsite ? "\n商家有自己的官網。" : "\n商家沒有官網。",
      htmlContent.trim() ? "\n\n商家網站 HTML 內容（節錄）：\n" + htmlContent.slice(0, 8000) : "",
      pdfBase64 ? "\n（已提供 PDF 型錄，請一併分析）" : "",
    ].join("");

    var prompt = contextParts + `

請根據商家情況，判斷屬於哪種路徑：

路徑A：商家品項超過200隻，應建議使用現有平台工具（阿里後台AI發品 / Accio Work）
路徑B：商家品項不多，需要裂變策略擴充到200隻以上

請只回傳 JSON，格式如下：
{
  "path": "many" 或 "few",
  "summary": "一段話說明商家現況與建議方向（繁體中文，2-3句）",
  "productCount": 商家目前大約有幾隻產品（整數），
  "estimatedListings": 透過裂變策略預計可達到幾隻（整數，路徑A填0）,
  "hasWebsite": true或false,
  "hasCatalog": true或false,
  "products": [從商家提供資訊中提取的產品名稱列表，最多20筆，格式：{"name":"產品名稱"}],
  "strategy": {
    "dim1Label": "裂變維度1的名稱（例：規格/成分/材質）",
    "dim1Examples": ["範例1","範例2","範例3"],
    "dim2Label": "裂變維度2的名稱（例：應用場景/目標受眾）",
    "dim2Examples": ["範例1","範例2","範例3"],
    "dim3Label": "裂變維度3的名稱（例：B2B服務模式）",
    "dim3Examples": ["範例1","範例2","範例3"]
  }
}`;

    try {
      var raw = await callLLM(prompt, pdfBase64, pdfBase64 ? "application/pdf" : null, 6000);
      var result = JSON.parse(raw);
      setAnalysis(result);

      if (result.path === "many") {
        setStep("result-many");
      } else {
        var prods = (result.products || []).map(function (p) {
          return {
            name: p.name || "",
            dim1: "",
            dim2: "",
            dim3: "",
            dim1Options: result.strategy ? result.strategy.dim1Examples || [] : [],
            dim2Options: result.strategy ? result.strategy.dim2Examples || [] : [],
            dim3Options: result.strategy ? result.strategy.dim3Examples || [] : [],
            listings: [],
            generating: false,
            done: false,
            error: ""
          };
        });
        // 若 AI 沒提取到產品，給一個空白列
        if (!prods.length) prods = [{ name: "", dim1: "", dim2: "", dim3: "", dim1Options: [], dim2Options: [], dim3Options: [], listings: [], generating: false, done: false, error: "" }];
        setProducts(prods);
        setStep("result-few");
      }
    } catch (e) {
      setError("分析失敗：" + e.message);
      setStep("intake");
    } finally {
      setLoading(false);
    }
  }

  // ── 單一產品擴品 ──
  async function expandProduct(idx) {
    var p = products[idx];
    if (!p.name.trim()) return;
    setProducts(function (prev) {
      var next = [...prev];
      next[idx] = { ...next[idx], generating: true, error: "", listings: [] };
      return next;
    });

    var dim1 = p.dim1.trim() || (p.dim1Options[0] || "");
    var dim2 = p.dim2.trim() || (p.dim2Options[0] || "");
    var dim3 = p.dim3.trim() || (p.dim3Options[0] || "");
    var strategy = analysis ? analysis.strategy : {};

    var prompt = [
      "你是阿里巴巴國際站電商專家。請根據以下裂變策略，為這個產品生成20組差異化的Listing。\n",
      "產品名稱：" + p.name + "\n",
      (strategy.dim1Label || "規格/成分") + "（維度1）：" + dim1 + "\n",
      (strategy.dim2Label || "應用場景/受眾") + "（維度2）：" + dim2 + "\n",
      (strategy.dim3Label || "B2B服務模式") + "（維度3）：" + dim3 + "\n\n",
      "請將3個維度交叉組合，生成20組標題各不相同的差異化Listing。\n",
      "每組包含：title_en（英文標題，150字元內）、title_zh（中文標題）、keywords_en（5個英文關鍵詞，逗號分隔）、keywords_zh（5個中文關鍵詞，頓號分隔）、dimension（本組主打的維度說明）\n\n",
      "只回傳JSON陣列，不要其他文字：\n",
      '[{"title_en":"","title_zh":"","keywords_en":"","keywords_zh":"","dimension":""}]'
    ].join("");

    try {
      var raw = await callLLM(prompt, null, null, 16000);
      var listings = JSON.parse(raw);
      setProducts(function (prev) {
        var next = [...prev];
        next[idx] = { ...next[idx], generating: false, done: true, listings: listings };
        return next;
      });
    } catch (e) {
      setProducts(function (prev) {
        var next = [...prev];
        next[idx] = { ...next[idx], generating: false, error: "生成失敗：" + e.message };
        return next;
      });
    }
  }

  function addProduct() {
    var strategy = analysis ? analysis.strategy : {};
    setProducts(function (prev) {
      return [...prev, {
        name: "", dim1: "", dim2: "", dim3: "",
        dim1Options: strategy ? strategy.dim1Examples || [] : [],
        dim2Options: strategy ? strategy.dim2Examples || [] : [],
        dim3Options: strategy ? strategy.dim3Examples || [] : [],
        listings: [], generating: false, done: false, error: ""
      }];
    });
  }

  function updateProduct(idx, field, value) {
    setProducts(function (prev) {
      var next = [...prev];
      next[idx] = { ...next[idx], [field]: value };
      return next;
    });
  }

  function exportAllToExcel() {
    var rows = [["產品名稱", "維度1", "維度2", "維度3", "英文標題", "中文標題", "英文關鍵詞", "中文關鍵詞", "維度說明"]];
    products.forEach(function (p) {
      if (!p.listings.length) return;
      p.listings.forEach(function (l) {
        rows.push([p.name, p.dim1, p.dim2, p.dim3, l.title_en, l.title_zh, l.keywords_en, l.keywords_zh, l.dimension]);
      });
    });
    var ws = XLSX.utils.aoa_to_sheet(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "裂變發品");
    XLSX.writeFile(wb, "advisor_listings.xlsx");
    setExportMsg("✓ 已下載");
    setTimeout(function () { setExportMsg(""); }, 2500);
  }

  var totalListings = products.reduce(function (s, p) { return s + p.listings.length; }, 0);
  var doneCount = products.filter(function (p) { return p.done; }).length;

  // ════════ RENDER ════════
  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Inter, PingFang TC, sans-serif" }}>

      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#1B2A4A,#2D4270)", padding: "16px 32px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 26 }}>🚀</span>
          <div>
            <div style={{ color: "#fff", fontSize: 18, fontWeight: 800 }}>阿里巴巴智慧發品顧問</div>
            <div style={{ color: "#8BA0C4", fontSize: 12 }}>AI 分析商家現況 → 量身訂製發品策略</div>
          </div>
        </div>
        <a href="../" style={{ color: "#8BA0C4", fontSize: 12, textDecoration: "none" }}>← 回到 Listing 擴張器</a>
      </div>

      <div style={{ maxWidth: 860, margin: "0 auto", padding: "32px 24px" }}>

        {/* ══ STEP 1: INTAKE ══ */}
        {(step === "intake" || step === "analyzing") && (
          <div>
            <Card>
              <SectionTitle>📋 第一步：告訴我你的狀況</SectionTitle>

              <div style={{ marginBottom: 20 }}>
                <div style={{ fontWeight: 600, fontSize: 14, color: C.text, marginBottom: 10 }}>你目前大約有多少個產品可上架？</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 10 }}>
                  {[
                    { v: "<10", label: "10 隻以下", desc: "品項很少，需要擴品策略" },
                    { v: "10-50", label: "10–50 隻", desc: "有一定品項，可擴大布局" },
                    { v: "50-200", label: "50–200 隻", desc: "品項中等，選擇性擴充" },
                    { v: "200+", label: "200 隻以上", desc: "品項豐富，需整理上架策略" },
                  ].map(function (opt) {
                    return (
                      <div key={opt.v} onClick={function () { setProductRange(opt.v); }}
                        style={{ border: "2px solid " + (productRange === opt.v ? C.orange : C.border), borderRadius: 10, padding: "14px 16px", cursor: "pointer", background: productRange === opt.v ? "#FFF7ED" : "#fff", transition: "all 0.15s" }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: productRange === opt.v ? C.orange : C.text }}>{opt.label}</div>
                        <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{opt.desc}</div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div style={{ marginBottom: 20 }}>
                <div style={{ fontWeight: 600, fontSize: 14, color: C.text, marginBottom: 10 }}>你有自己的官網嗎？</div>
                <div style={{ display: "flex", gap: 10 }}>
                  {[{ v: true, label: "✅ 有官網" }, { v: false, label: "❌ 沒有官網" }].map(function (opt) {
                    return (
                      <div key={String(opt.v)} onClick={function () { setHasWebsite(opt.v); }}
                        style={{ flex: 1, border: "2px solid " + (hasWebsite === opt.v ? C.orange : C.border), borderRadius: 10, padding: "12px 16px", cursor: "pointer", background: hasWebsite === opt.v ? "#FFF7ED" : "#fff", textAlign: "center", fontWeight: 600, fontSize: 14, color: hasWebsite === opt.v ? C.orange : C.text, transition: "all 0.15s" }}>
                        {opt.label}
                      </div>
                    );
                  })}
                </div>
              </div>

              {hasWebsite && (
                <div style={{ marginBottom: 20 }}>
                  <div style={{ fontWeight: 600, fontSize: 14, color: C.text, marginBottom: 6 }}>貼上官網首頁的 HTML 原始碼（可選）</div>
                  <div style={{ fontSize: 12, color: C.muted, marginBottom: 8 }}>在瀏覽器按右鍵 → 「檢視頁面原始碼」→ 全選複製</div>
                  <textarea value={htmlContent} onChange={function (e) { setHtmlContent(e.target.value); }}
                    placeholder="貼上 HTML 原始碼，AI 會自動分析你的產品線..."
                    style={{ ...C.inp, height: 120, resize: "vertical", fontFamily: "monospace", fontSize: 12 }} />
                </div>
              )}

              <div style={{ marginBottom: 20 }}>
                <div style={{ fontWeight: 600, fontSize: 14, color: C.text, marginBottom: 6 }}>上傳 PDF 產品型錄（可選）</div>
                <div onClick={function () { fileRef.current.click(); }}
                  style={{ border: "2px dashed " + C.border, borderRadius: 10, padding: "20px", textAlign: "center", cursor: "pointer", background: pdfBase64 ? "#F0FDF4" : "#FAFAFA" }}>
                  <div style={{ fontSize: 22, marginBottom: 6 }}>{pdfBase64 ? "📄" : "📎"}</div>
                  <div style={{ fontSize: 13, color: pdfBase64 ? C.green : C.muted }}>{pdfName || "點擊上傳 PDF 型錄"}</div>
                </div>
                <input ref={fileRef} type="file" accept=".pdf" style={{ display: "none" }} onChange={handlePdf} />
              </div>

              {error && <div style={{ background: "#FFF1F2", color: "#E11D48", padding: "10px 14px", borderRadius: 8, fontSize: 13, marginBottom: 14 }}>{error}</div>}

              <Btn onClick={analyze} disabled={loading || !productRange} color="orange" style={{ width: "100%" }}>
                {loading ? ("⏳ AI 分析中… " + loadingSecs + "s") : "🔍 開始分析，給我發品建議 →"}
              </Btn>
            </Card>
          </div>
        )}

        {/* ══ RESULT: 品項多（200+）══ */}
        {step === "result-many" && analysis && (
          <div>
            <Card>
              <SectionTitle>📊 分析結果</SectionTitle>
              <div style={{ background: "#F0F9FF", border: "1px solid #BAE6FD", borderRadius: 10, padding: 16, marginBottom: 20 }}>
                <div style={{ fontWeight: 700, color: "#0369A1", marginBottom: 6 }}>✅ 你的品項很豐富！</div>
                <div style={{ fontSize: 14, color: C.text, lineHeight: 1.7 }}>{analysis.summary}</div>
              </div>

              <div style={{ fontWeight: 700, fontSize: 15, color: C.navy, marginBottom: 14 }}>建議發品路徑</div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                <Card style={{ border: "2px solid #BFDBFE", margin: 0 }}>
                  <div style={{ fontWeight: 700, color: "#1D4ED8", marginBottom: 8 }}>路徑 A：有官網</div>
                  <div style={{ fontSize: 13, color: C.text, lineHeight: 1.7, marginBottom: 16 }}>
                    使用阿里後台 AI 發品功能，或 Accio Work 一鍵貼上產品網址批次上架，效率最高。
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <a href="https://seller.alibaba.com" target="_blank" style={{ display: "block", padding: "10px 14px", background: "#1D4ED8", color: "#fff", borderRadius: 8, textAlign: "center", fontSize: 13, fontWeight: 700, textDecoration: "none" }}>→ 阿里後台 AI 發品</a>
                    <a href="https://work.accio.com" target="_blank" style={{ display: "block", padding: "10px 14px", background: "#EFF6FF", color: "#1D4ED8", borderRadius: 8, textAlign: "center", fontSize: 13, fontWeight: 700, textDecoration: "none", border: "1px solid #BFDBFE" }}>→ Accio Work</a>
                  </div>
                </Card>

                <Card style={{ border: "2px solid #D1FAE5", margin: 0 }}>
                  <div style={{ fontWeight: 700, color: C.green, marginBottom: 8 }}>路徑 B：有 PDF 型錄</div>
                  <div style={{ fontSize: 13, color: C.text, lineHeight: 1.7, marginBottom: 16 }}>
                    AI 從你的型錄中萃取產品資料，整理成阿里標準 Excel 格式，直接匯入後台批次上架。
                  </div>
                  <Btn color="green" style={{ width: "100%" }} onClick={function () { alert("此功能開發中，敬請期待！"); }}>
                    → 從型錄產出 Excel（即將推出）
                  </Btn>
                </Card>
              </div>
            </Card>

            <Btn color="ghost" onClick={function () { setStep("intake"); setAnalysis(null); }} style={{ marginTop: 4 }}>← 重新填寫</Btn>
          </div>
        )}

        {/* ══ RESULT: 品項少，需裂變 ══ */}
        {step === "result-few" && analysis && (
          <div>
            {/* 摘要 */}
            <Card>
              <SectionTitle>🎯 發品策略分析</SectionTitle>
              <div style={{ background: "#FFF7ED", border: "1px solid #FED7AA", borderRadius: 10, padding: 16, marginBottom: 12 }}>
                <div style={{ fontWeight: 700, color: C.orange, marginBottom: 6, fontSize: 15 }}>
                  你目前約有 {analysis.productCount} 隻產品，透過裂變策略預計可發展出 {analysis.estimatedListings || analysis.productCount * 20}+ 隻 listing！
                </div>
                <div style={{ fontSize: 14, color: C.text, lineHeight: 1.7 }}>{analysis.summary}</div>
              </div>

              {/* 裂變框架說明 */}
              {analysis.strategy && (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
                  {[
                    { label: analysis.strategy.dim1Label, examples: analysis.strategy.dim1Examples, color: "#4F46E5", bg: "#EEF2FF" },
                    { label: analysis.strategy.dim2Label, examples: analysis.strategy.dim2Examples, color: "#0369A1", bg: "#F0F9FF" },
                    { label: analysis.strategy.dim3Label, examples: analysis.strategy.dim3Examples, color: C.green, bg: "#F0FDF4" },
                  ].map(function (d, i) {
                    return (
                      <div key={i} style={{ background: d.bg, borderRadius: 10, padding: 14 }}>
                        <div style={{ fontWeight: 700, fontSize: 13, color: d.color, marginBottom: 8 }}>維度 {i + 1}｜{d.label}</div>
                        {(d.examples || []).map(function (ex, j) {
                          return <div key={j} style={{ fontSize: 12, color: C.text, padding: "3px 8px", background: "#fff", borderRadius: 5, marginBottom: 4, border: "1px solid " + d.color + "44" }}>{ex}</div>;
                        })}
                      </div>
                    );
                  })}
                </div>
              )}
            </Card>

            {/* 產品清單 */}
            <Card>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
                <SectionTitle style={{ marginBottom: 0, paddingBottom: 0, borderBottom: "none" }}>📦 產品清單（最多 5 隻）</SectionTitle>
                {products.length < 5 && (
                  <Btn color="ghost" onClick={addProduct} style={{ fontSize: 13, padding: "7px 14px" }}>+ 新增產品</Btn>
                )}
              </div>

              {products.map(function (p, idx) {
                return (
                  <div key={idx} style={{ border: "1px solid " + C.border, borderRadius: 12, padding: 18, marginBottom: 14, background: p.done ? "#F0FDF4" : "#FAFAFA" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                      <div style={{ width: 26, height: 26, borderRadius: 7, background: p.done ? C.green : C.navy, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 800, flexShrink: 0 }}>
                        {p.done ? "✓" : idx + 1}
                      </div>
                      <input value={p.name} onChange={function (e) { updateProduct(idx, "name", e.target.value); }}
                        placeholder="產品名稱（例：塑膠射出模具）"
                        style={{ ...C.inp, fontWeight: 600, fontSize: 14 }} />
                      {products.length > 1 && (
                        <button onClick={function () { setProducts(function (prev) { return prev.filter(function (_, i) { return i !== idx; }); }); }}
                          style={{ background: "none", border: "none", color: "#9CA3AF", cursor: "pointer", fontSize: 18, flexShrink: 0 }}>✕</button>
                      )}
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginBottom: 12 }}>
                      {[
                        { field: "dim1", label: (analysis.strategy && analysis.strategy.dim1Label) || "維度1", opts: p.dim1Options },
                        { field: "dim2", label: (analysis.strategy && analysis.strategy.dim2Label) || "維度2", opts: p.dim2Options },
                        { field: "dim3", label: (analysis.strategy && analysis.strategy.dim3Label) || "維度3", opts: p.dim3Options },
                      ].map(function (d) {
                        return (
                          <div key={d.field}>
                            <div style={{ fontSize: 12, fontWeight: 600, color: C.muted, marginBottom: 5 }}>{d.label}</div>
                            <input value={p[d.field]} onChange={function (e) { updateProduct(idx, d.field, e.target.value); }}
                              placeholder={d.opts.slice(0, 2).join(" / ") || "填入裂變內容"}
                              style={{ ...C.inp, fontSize: 12 }} />
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginTop: 5 }}>
                              {d.opts.map(function (opt, oi) {
                                return (
                                  <span key={oi} onClick={function () { updateProduct(idx, d.field, p[d.field] ? p[d.field] + " / " + opt : opt); }}
                                    style={{ fontSize: 11, padding: "2px 8px", background: "#EEF2FF", color: "#4F46E5", borderRadius: 4, cursor: "pointer" }}>
                                    + {opt}
                                  </span>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {p.error && <div style={{ color: "#E11D48", fontSize: 12, marginBottom: 8 }}>{p.error}</div>}

                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                      {p.done && <span style={{ fontSize: 13, color: C.green, fontWeight: 600 }}>✅ 已生成 {p.listings.length} 組 Listing</span>}
                      {!p.done && <span />}
                      <Btn onClick={function () { expandProduct(idx); }} disabled={p.generating || !p.name.trim()} color={p.done ? "green" : "orange"} style={{ fontSize: 13, padding: "8px 18px" }}>
                        {p.generating ? ("⏳ " + loadingSecs + "s…") : p.done ? "🔄 重新生成" : "⚡ AI 擴品"}
                      </Btn>
                    </div>
                  </div>
                );
              })}
            </Card>

            {/* 匯出 */}
            {totalListings > 0 && (
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#fff", border: "1px solid " + C.border, borderRadius: 12, padding: "14px 20px" }}>
                <div style={{ fontSize: 14, color: C.text }}>
                  共 <strong>{doneCount}</strong> 個產品完成，合計 <strong>{totalListings}</strong> 組 Listing
                </div>
                <Btn onClick={exportAllToExcel} color="navy">
                  {exportMsg || "⬇ 匯出 Excel"}
                </Btn>
              </div>
            )}

            <Btn color="ghost" onClick={function () { setStep("intake"); setAnalysis(null); setProducts([]); }} style={{ marginTop: 16 }}>← 重新填寫</Btn>
          </div>
        )}

      </div>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<Advisor />);
