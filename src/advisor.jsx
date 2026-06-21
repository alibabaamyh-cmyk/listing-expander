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
  var text = (data.content || []).map(function (x) { return x.text || ""; }).join("");
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

const C = {
  bg: "#F8FAFC", navy: "#1B2A4A", orange: "#FF6B35", green: "#16A34A",
  card: "#fff", border: "#E5E7EB", text: "#1e293b", muted: "#64748b",
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
      style={{ padding: "11px 22px", background: disabled ? "#E5E7EB" : bg, color: disabled ? "#9CA3AF" : fg, border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer", transition: "background 0.15s", ...style }}>
      {children}
    </button>
  );
}

// ════════════════════════════════════════════════════════
// STEPS:
//   intake → [analyzing-path] → result-many
//                             → product-input → [analyzing-dims] → result-few
// ════════════════════════════════════════════════════════
export default function Advisor() {
  const [step, setStep] = useState("intake");

  // Step 1: intake
  const [productRange, setProductRange] = useState("");
  const [hasWebsite, setHasWebsite] = useState(null);
  const [htmlContent, setHtmlContent] = useState("");
  const [pdfBase64, setPdfBase64] = useState(null);
  const [pdfName, setPdfName] = useState("");
  const fileRef = useRef();

  // loading
  const [loading, setLoading] = useState(false);
  const [loadingSecs, setLoadingSecs] = useState(0);
  const [error, setError] = useState("");

  // Step 2 path result
  const [pathResult, setPathResult] = useState(null);
  // { path:"many"|"few", summary, hasWebsite, hasCatalog }

  // Step 2b: 商家自填產品清單（最多5隻）
  const [inputProducts, setInputProducts] = useState([{ name: "" }]);

  // Step 3: 維度分析結果 + 擴品
  const [strategy, setStrategy] = useState(null);
  // { dim1Label, dim1Examples[], dim2Label, dim2Examples[], dim3Label, dim3Examples[], estimatedListings, strategySummary }
  const [products, setProducts] = useState([]);
  // [{name, dim1, dim2, dim3, listings:[], generating:false, done:false, error:""}]

  const [exportMsg, setExportMsg] = useState("");

  // 計時器
  useEffect(function () {
    if (!loading) { setLoadingSecs(0); return; }
    setLoadingSecs(0);
    var t = setInterval(function () { setLoadingSecs(function (s) { return s + 1; }); }, 1000);
    return function () { clearInterval(t); };
  }, [loading]);

  function handlePdf(e) {
    var f = e.target.files[0];
    if (!f) return;
    setPdfName(f.name);
    var reader = new FileReader();
    reader.onload = function (ev) { setPdfBase64(ev.target.result.split(",")[1]); };
    reader.readAsDataURL(f);
  }

  // ── Step 1 → 判斷路徑 ──
  async function analyzePath() {
    if (!productRange) { setError("請先選擇產品數量範圍"); return; }
    setError(""); setLoading(true);

    var ctx = [
      "商家產品數量範圍：" + productRange,
      hasWebsite ? "有自己的官網。" : "沒有官網。",
      htmlContent.trim() ? "網站HTML（節錄）：\n" + htmlContent.slice(0, 3000) : "",
      pdfBase64 ? "（已附上PDF型錄）" : "",
    ].filter(Boolean).join("\n");

    var prompt = "你是阿里巴巴國際站電商顧問。\n" + ctx + "\n\n" +
      "判斷商家屬於哪個路徑：\n" +
      "路徑A（many）：品項200隻以上，建議用平台工具整批上架\n" +
      "路徑B（few）：品項不足，需要裂變策略擴充到200隻\n\n" +
      "只回傳JSON：\n" +
      '{"path":"many或few","summary":"2句話說明現況與建議方向（繁體中文）","hasWebsite":true或false,"hasCatalog":true或false}';

    try {
      var raw = await callLLM(prompt, pdfBase64, pdfBase64 ? "application/pdf" : null, 2000);
      var result = JSON.parse(raw);
      setPathResult(result);
      if (result.path === "many") {
        setStep("result-many");
      } else {
        setStep("product-input");
      }
    } catch (e) {
      setError("分析失敗：" + e.message);
    } finally {
      setLoading(false);
    }
  }

  // ── Step 2b → AI 分析具體產品，給客製化維度 ──
  async function analyzeDimensions() {
    var names = inputProducts.map(function (p) { return p.name.trim(); }).filter(Boolean);
    if (!names.length) { setError("請至少填寫一個產品名稱"); return; }
    setError(""); setLoading(true);

    var ctx = [
      "商家品項數量範圍：" + productRange,
      pdfBase64 ? "（已附上PDF型錄）" : "",
      htmlContent.trim() ? "網站HTML（節錄）：\n" + htmlContent.slice(0, 2000) : "",
    ].filter(Boolean).join("\n");

    var prompt = "你是阿里巴巴國際站電商顧問，專精跨境B2B裂變策略。\n\n" +
      "商家背景：\n" + ctx + "\n\n" +
      "商家的產品：\n" + names.map(function (n, i) { return (i + 1) + ". " + n; }).join("\n") + "\n\n" +
      "請針對這些具體產品，設計最適合的3個裂變維度。每個維度給3-5個實際可用的範例（不要太抽象）。\n\n" +
      "只回傳JSON：\n" +
      '{"strategySummary":"針對這些產品的裂變策略說明（繁體中文，2句）","estimatedListings":預計可達到的listing總數（整數）,' +
      '"dim1Label":"維度1名稱","dim1Examples":["具體範例1","具體範例2","具體範例3"],' +
      '"dim2Label":"維度2名稱","dim2Examples":["具體範例1","具體範例2","具體範例3"],' +
      '"dim3Label":"維度3名稱","dim3Examples":["具體範例1","具體範例2","具體範例3"]}';

    try {
      var raw = await callLLM(prompt, pdfBase64, pdfBase64 ? "application/pdf" : null, 3000);
      var result = JSON.parse(raw);
      setStrategy(result);
      var prods = names.map(function (name) {
        return {
          name: name,
          dim1: (result.dim1Examples || [])[0] || "",
          dim2: (result.dim2Examples || [])[0] || "",
          dim3: (result.dim3Examples || [])[0] || "",
          listings: [], generating: false, done: false, error: ""
        };
      });
      setProducts(prods);
      setStep("result-few");
    } catch (e) {
      setError("分析失敗：" + e.message);
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

    var prompt = [
      "你是阿里巴巴國際站電商專家。請根據以下裂變策略，為這個產品生成20組差異化Listing。\n",
      "產品名稱：" + p.name + "\n",
      (strategy && strategy.dim1Label ? strategy.dim1Label : "規格/成分") + "（維度1）：" + p.dim1 + "\n",
      (strategy && strategy.dim2Label ? strategy.dim2Label : "應用場景/受眾") + "（維度2）：" + p.dim2 + "\n",
      (strategy && strategy.dim3Label ? strategy.dim3Label : "B2B服務模式") + "（維度3）：" + p.dim3 + "\n\n",
      "將3個維度交叉組合，生成20組標題各不相同的差異化Listing。\n",
      "每組包含：title_en（英文標題150字元內）、title_zh（中文標題）、keywords_en（5個英文關鍵詞逗號分隔）、keywords_zh（5個中文關鍵詞頓號分隔）、dimension（本組主打的維度說明，10字以內）\n\n",
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

  function updateProduct(idx, field, value) {
    setProducts(function (prev) { var n = [...prev]; n[idx] = { ...n[idx], [field]: value }; return n; });
  }

  function exportAllToExcel() {
    var rows = [["產品名稱", "維度1", "維度2", "維度3", "英文標題", "中文標題", "英文關鍵詞", "中文關鍵詞", "維度說明"]];
    products.forEach(function (p) {
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

  // ════ RENDER ════
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

      <div style={{ maxWidth: 820, margin: "0 auto", padding: "32px 24px" }}>

        {/* ══ STEP 1: INTAKE ══ */}
        {step === "intake" && (
          <Card>
            <SectionTitle>📋 告訴我你的現況</SectionTitle>

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
                <div style={{ fontWeight: 600, fontSize: 14, color: C.text, marginBottom: 6 }}>貼上官網首頁 HTML 原始碼（可選）</div>
                <div style={{ fontSize: 12, color: C.muted, marginBottom: 8 }}>瀏覽器右鍵 → 「檢視頁面原始碼」→ 全選複製</div>
                <textarea value={htmlContent} onChange={function (e) { setHtmlContent(e.target.value); }}
                  placeholder="貼上 HTML 原始碼，AI 會分析你的產品線..."
                  style={{ ...C.inp, height: 110, resize: "vertical", fontFamily: "monospace", fontSize: 12 }} />
              </div>
            )}

            <div style={{ marginBottom: 24 }}>
              <div style={{ fontWeight: 600, fontSize: 14, color: C.text, marginBottom: 6 }}>上傳 PDF 產品型錄（可選）</div>
              <div onClick={function () { fileRef.current.click(); }}
                style={{ border: "2px dashed " + C.border, borderRadius: 10, padding: "18px", textAlign: "center", cursor: "pointer", background: pdfBase64 ? "#F0FDF4" : "#FAFAFA" }}>
                <div style={{ fontSize: 22, marginBottom: 4 }}>{pdfBase64 ? "📄" : "📎"}</div>
                <div style={{ fontSize: 13, color: pdfBase64 ? C.green : C.muted }}>{pdfName || "點擊上傳 PDF 型錄"}</div>
              </div>
              <input ref={fileRef} type="file" accept=".pdf" style={{ display: "none" }} onChange={handlePdf} />
            </div>

            {error && <div style={{ background: "#FFF1F2", color: "#E11D48", padding: "10px 14px", borderRadius: 8, fontSize: 13, marginBottom: 14 }}>{error}</div>}

            <Btn onClick={analyzePath} disabled={loading || !productRange} color="orange" style={{ width: "100%" }}>
              {loading ? ("⏳ 分析中… " + loadingSecs + "s") : "下一步 →"}
            </Btn>
          </Card>
        )}

        {/* ══ STEP 2b: 商家輸入產品 ══ */}
        {step === "product-input" && (
          <div>
            {pathResult && (
              <div style={{ background: "#FFF7ED", border: "1px solid #FED7AA", borderRadius: 12, padding: 16, marginBottom: 20 }}>
                <div style={{ fontWeight: 700, color: C.orange, marginBottom: 4 }}>📊 初步分析</div>
                <div style={{ fontSize: 14, color: C.text, lineHeight: 1.7 }}>{pathResult.summary}</div>
              </div>
            )}

            <Card>
              <SectionTitle>📦 你的產品有哪些？</SectionTitle>
              <div style={{ fontSize: 14, color: C.muted, marginBottom: 16 }}>
                填入你想上架的產品名稱（最多 5 個），AI 會針對這些具體產品給你客製化的裂變維度建議。
              </div>

              {inputProducts.map(function (p, idx) {
                return (
                  <div key={idx} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                    <div style={{ width: 28, height: 28, borderRadius: 7, background: C.navy, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 800, flexShrink: 0 }}>{idx + 1}</div>
                    <input value={p.name}
                      onChange={function (e) {
                        var v = e.target.value;
                        setInputProducts(function (prev) { var n = [...prev]; n[idx] = { name: v }; return n; });
                      }}
                      placeholder={"產品名稱（例：不鏽鋼內六角螺絲）"}
                      style={{ ...C.inp }} />
                    {inputProducts.length > 1 && (
                      <button onClick={function () { setInputProducts(function (prev) { return prev.filter(function (_, i) { return i !== idx; }); }); }}
                        style={{ background: "none", border: "none", color: "#9CA3AF", cursor: "pointer", fontSize: 20, flexShrink: 0 }}>✕</button>
                    )}
                  </div>
                );
              })}

              {inputProducts.length < 5 && (
                <button onClick={function () { setInputProducts(function (prev) { return [...prev, { name: "" }]; }); }}
                  style={{ background: "none", border: "1.5px dashed " + C.border, borderRadius: 9, padding: "9px 18px", fontSize: 13, color: C.muted, cursor: "pointer", width: "100%", marginTop: 4 }}>
                  + 新增產品
                </button>
              )}

              {error && <div style={{ background: "#FFF1F2", color: "#E11D48", padding: "10px 14px", borderRadius: 8, fontSize: 13, margin: "14px 0" }}>{error}</div>}

              <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
                <Btn color="ghost" onClick={function () { setStep("intake"); setError(""); }}>← 返回</Btn>
                <Btn color="orange" disabled={loading || !inputProducts.some(function (p) { return p.name.trim(); })}
                  onClick={analyzeDimensions} style={{ flex: 1 }}>
                  {loading ? ("⏳ AI 分析中… " + loadingSecs + "s（針對你的產品設計裂變策略）") : "⚡ AI 幫我設計裂變策略 →"}
                </Btn>
              </div>
            </Card>
          </div>
        )}

        {/* ══ RESULT-MANY ══ */}
        {step === "result-many" && pathResult && (
          <div>
            <Card>
              <SectionTitle>📊 分析結果</SectionTitle>
              <div style={{ background: "#F0F9FF", border: "1px solid #BAE6FD", borderRadius: 10, padding: 16, marginBottom: 20 }}>
                <div style={{ fontWeight: 700, color: "#0369A1", marginBottom: 6 }}>✅ 你的品項很豐富！</div>
                <div style={{ fontSize: 14, color: C.text, lineHeight: 1.7 }}>{pathResult.summary}</div>
              </div>

              <div style={{ fontWeight: 700, fontSize: 15, color: C.navy, marginBottom: 14 }}>建議發品路徑</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                <Card style={{ border: "2px solid #BFDBFE", margin: 0 }}>
                  <div style={{ fontWeight: 700, color: "#1D4ED8", marginBottom: 8 }}>路徑 A：有官網</div>
                  <div style={{ fontSize: 13, color: C.text, lineHeight: 1.7, marginBottom: 14 }}>使用阿里後台 AI 發品，或 Accio Work 一鍵貼網址批次上架。</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <a href="https://seller.alibaba.com" target="_blank" style={{ display: "block", padding: "9px 14px", background: "#1D4ED8", color: "#fff", borderRadius: 8, textAlign: "center", fontSize: 13, fontWeight: 700, textDecoration: "none" }}>→ 阿里後台 AI 發品</a>
                    <a href="https://work.accio.com" target="_blank" style={{ display: "block", padding: "9px 14px", background: "#EFF6FF", color: "#1D4ED8", borderRadius: 8, textAlign: "center", fontSize: 13, fontWeight: 700, textDecoration: "none", border: "1px solid #BFDBFE" }}>→ Accio Work</a>
                  </div>
                </Card>
                <Card style={{ border: "2px solid #D1FAE5", margin: 0 }}>
                  <div style={{ fontWeight: 700, color: C.green, marginBottom: 8 }}>路徑 B：有 PDF 型錄</div>
                  <div style={{ fontSize: 13, color: C.text, lineHeight: 1.7, marginBottom: 14 }}>AI 從型錄萃取產品資料，整理成阿里標準 Excel 格式直接匯入。</div>
                  <Btn color="green" style={{ width: "100%", fontSize: 13 }} onClick={function () { alert("此功能開發中，敬請期待！"); }}>→ 從型錄產出 Excel（即將推出）</Btn>
                </Card>
              </div>
            </Card>
            <Btn color="ghost" onClick={function () { setStep("intake"); setPathResult(null); }}>← 重新填寫</Btn>
          </div>
        )}

        {/* ══ RESULT-FEW: 裂變擴品 ══ */}
        {step === "result-few" && strategy && (
          <div>
            {/* 策略摘要 */}
            <Card>
              <SectionTitle>🎯 你的專屬裂變策略</SectionTitle>
              <div style={{ background: "#FFF7ED", border: "1px solid #FED7AA", borderRadius: 10, padding: 16, marginBottom: 16 }}>
                <div style={{ fontWeight: 700, color: C.orange, fontSize: 15, marginBottom: 6 }}>
                  預計可發展出 {strategy.estimatedListings || products.length * 20}+ 隻 listing！
                </div>
                <div style={{ fontSize: 14, color: C.text, lineHeight: 1.7 }}>{strategy.strategySummary}</div>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
                {[
                  { label: strategy.dim1Label, examples: strategy.dim1Examples, color: "#4F46E5", bg: "#EEF2FF" },
                  { label: strategy.dim2Label, examples: strategy.dim2Examples, color: "#0369A1", bg: "#F0F9FF" },
                  { label: strategy.dim3Label, examples: strategy.dim3Examples, color: C.green, bg: "#F0FDF4" },
                ].map(function (d, i) {
                  return (
                    <div key={i} style={{ background: d.bg, borderRadius: 10, padding: 14 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: d.color, marginBottom: 8 }}>維度 {i + 1}｜{d.label}</div>
                      {(d.examples || []).map(function (ex, j) {
                        return <div key={j} style={{ fontSize: 12, color: C.text, padding: "3px 8px", background: "#fff", borderRadius: 5, marginBottom: 4, border: "1px solid " + d.color + "33" }}>{ex}</div>;
                      })}
                    </div>
                  );
                })}
              </div>
            </Card>

            {/* 產品個別擴品 */}
            <Card>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
                <SectionTitle style={{ margin: 0, border: "none", padding: 0 }}>📦 逐一擴品</SectionTitle>
                {totalListings > 0 && (
                  <Btn onClick={exportAllToExcel} color="navy" style={{ fontSize: 13, padding: "8px 16px" }}>
                    {exportMsg || ("⬇ 匯出 Excel（" + totalListings + " 組）")}
                  </Btn>
                )}
              </div>

              {products.map(function (p, idx) {
                return (
                  <div key={idx} style={{ border: "1px solid " + C.border, borderRadius: 12, padding: 18, marginBottom: 14, background: p.done ? "#F0FDF4" : "#FAFAFA" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
                      <div style={{ width: 26, height: 26, borderRadius: 7, background: p.done ? C.green : C.navy, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 800, flexShrink: 0 }}>
                        {p.done ? "✓" : idx + 1}
                      </div>
                      <div style={{ fontWeight: 700, fontSize: 15, color: C.text }}>{p.name}</div>
                      {p.done && <span style={{ fontSize: 12, color: C.green, fontWeight: 600 }}>已生成 {p.listings.length} 組</span>}
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginBottom: 12 }}>
                      {[
                        { field: "dim1", label: strategy.dim1Label, opts: strategy.dim1Examples || [], color: "#4F46E5" },
                        { field: "dim2", label: strategy.dim2Label, opts: strategy.dim2Examples || [], color: "#0369A1" },
                        { field: "dim3", label: strategy.dim3Label, opts: strategy.dim3Examples || [], color: C.green },
                      ].map(function (d) {
                        return (
                          <div key={d.field}>
                            <div style={{ fontSize: 12, fontWeight: 600, color: d.color, marginBottom: 5 }}>{d.label}</div>
                            <input value={p[d.field]} onChange={function (e) { updateProduct(idx, d.field, e.target.value); }}
                              placeholder="填入或點選建議"
                              style={{ ...C.inp, fontSize: 12, borderColor: d.color + "66" }} />
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginTop: 5 }}>
                              {d.opts.map(function (opt, oi) {
                                var active = p[d.field] === opt;
                                return (
                                  <span key={oi} onClick={function () { updateProduct(idx, d.field, opt); }}
                                    style={{ fontSize: 11, padding: "3px 8px", background: active ? d.color : "#fff", color: active ? "#fff" : d.color, borderRadius: 4, cursor: "pointer", border: "1px solid " + d.color + "55", transition: "all 0.1s" }}>
                                    {opt}
                                  </span>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {p.error && <div style={{ color: "#E11D48", fontSize: 12, marginBottom: 8 }}>{p.error}</div>}

                    <div style={{ display: "flex", justifyContent: "flex-end" }}>
                      <Btn onClick={function () { expandProduct(idx); }} disabled={p.generating} color={p.done ? "green" : "orange"} style={{ fontSize: 13, padding: "8px 20px" }}>
                        {p.generating ? ("⏳ " + loadingSecs + "s…（約需 30-60 秒）") : p.done ? "🔄 重新生成" : "⚡ AI 擴品"}
                      </Btn>
                    </div>
                  </div>
                );
              })}
            </Card>

            <Btn color="ghost" onClick={function () { setStep("product-input"); setStrategy(null); setProducts([]); }}>← 修改產品清單</Btn>
          </div>
        )}

      </div>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<Advisor />);
