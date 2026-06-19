import { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ─── 常數 ───────────────────────────────────────────────
const MARKETS = ["北美市場","歐洲市場","東南亞市場","中東市場","拉丁美洲","非洲市場"];
const ADVANTAGES = ["價格優勢","品質保證","客製化能力","快速交貨","研發實力","國際認證齊全"];
const STYLES = ["現代簡約","專業企業","工業硬核","溫馨居家","科技未來","大氣極簡"];
const LISTING_MARKETS = ["全球通用","歐美市場","東南亞市場","中東市場","日韓市場"];
const PRESETS = [
  { label:"🏢 專業企業", bg:"#ffffff", primary:"#0052D9", text:"#1e293b", accent:"#FAAD14" },
  { label:"💻 科技未來", bg:"#111827", primary:"#3b82f6", text:"#f8fafc", accent:"#10b981" },
  { label:"🌿 環保永續", bg:"#ffffff", primary:"#10b981", text:"#374151", accent:"#f59e0b" },
  { label:"🚀 阿里熱情", bg:"#fdfdfd", primary:"#ff6b35", text:"#2d3436", accent:"#d63031" },
];
const DIMENSION_TAGS = {
  "應用場景":{ bg:"#EEF2FF", color:"#4F46E5", label:"場景" },
  "買家身份":{ bg:"#FFF7ED", color:"#C2410C", label:"買家" },
  "規格重點":{ bg:"#F0FDF4", color:"#16A34A", label:"規格" },
  "地區市場":{ bg:"#FDF4FF", color:"#9333EA", label:"市場" },
  "賣點角度":{ bg:"#FFF1F2", color:"#E11D48", label:"賣點" },
};

// ─── 共用元件 ────────────────────────────────────────────
function Tag({ type }) {
  const t = DIMENSION_TAGS[type] || { bg:"#F3F4F6", color:"#6B7280", label:type };
  return (
    <span style={{ background:t.bg, color:t.color, fontSize:11, fontWeight:700, padding:"2px 8px", borderRadius:4 }}>
      {t.label}
    </span>
  );
}

function CopyBtn({ text, label }) {
  const [copied, setCopied] = useState(false);
  const lbl = label || "複製";
  return (
    <button
      onClick={() => { navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 1500); }}
      style={{ background: copied ? "#D1FAE5" : "#F3F4F6", color: copied ? "#065F46" : "#374151", border:"none", borderRadius:6, padding:"4px 10px", fontSize:12, cursor:"pointer", fontWeight:500 }}
    >
      {copied ? "✓ 已複製" : lbl}
    </button>
  );
}

function Field({ label, hint, children }) {
  return (
    <div style={{ marginBottom:16 }}>
      <div style={{ fontSize:13, fontWeight:600, color:"#1e293b", marginBottom: hint ? 3 : 6 }}>{label}</div>
      {hint && <div style={{ fontSize:11, color:"#94a3b8", marginBottom:6, lineHeight:1.5 }}>{hint}</div>}
      {children}
    </div>
  );
}

function CheckGroup({ items, checked, onChange, cols }) {
  const c = cols || 2;
  return (
    <div style={{ display:"grid", gridTemplateColumns:"repeat("+c+", 1fr)", gap:8, marginTop:8 }}>
      {items.map(item => (
        <label key={item} style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer", fontSize:13, color:"#475569" }}>
          <input type="checkbox" checked={!!checked[item]} onChange={() => onChange(item)} style={{ width:16, height:16, accentColor:"#ff6b35", cursor:"pointer" }} />
          {item}
        </label>
      ))}
    </div>
  );
}

function OtherField({ checked, onCheck, value, onChange, placeholder }) {
  return (
    <div>
      <label style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer", fontSize:13, color:"#475569", marginTop:8 }}>
        <input type="checkbox" checked={checked} onChange={onCheck} style={{ width:16, height:16, accentColor:"#ff6b35", cursor:"pointer" }} />
        其它
      </label>
      {checked && (
        <input value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} style={{ width:"100%", padding:"8px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", marginTop:6 }} />
      )}
    </div>
  );
}

function StepBar({ step }) {
  const steps = ["填寫產品資訊","確認產品優勢","生成 Listing"];
  return (
    <div style={{ display:"flex", alignItems:"center", marginBottom:20 }}>
      {steps.map((s, i) => {
        const active = i + 1 === step;
        const done = i + 1 < step;
        return (
          <div key={i} style={{ display:"flex", alignItems:"center", flex: i < steps.length - 1 ? 1 : "none" }}>
            <div style={{ display:"flex", alignItems:"center", gap:7 }}>
              <div style={{ width:26, height:26, borderRadius:"50%", background: done ? "#16A34A" : active ? "#ff6b35" : "#E5E7EB", color: done || active ? "#fff" : "#9CA3AF", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:800 }}>
                {done ? "✓" : i + 1}
              </div>
              <span style={{ fontSize:12, fontWeight: active ? 700 : 400, color: active ? "#1e293b" : done ? "#16A34A" : "#9CA3AF", whiteSpace:"nowrap" }}>{s}</span>
            </div>
            {i < steps.length - 1 && <div style={{ flex:1, height:2, margin:"0 10px", background: done ? "#16A34A" : "#E5E7EB" }} />}
          </div>
        );
      })}
    </div>
  );
}

function ListingCard({ item, index, isChecked, onToggle }) {
  const full = "【" + item.dimension + "】\n標題EN：" + item.title_en + "\n標題中：" + item.title_zh + "\n關鍵詞EN：" + item.keywords_en + "\n關鍵詞中：" + item.keywords_zh + "\n屬性：\n" + (item.attributes||[]).map(a => "  " + a.name + "：" + a.value).join("\n") + "\n圖片Prompt：\n" + item.image_prompt;
  return (
    <div style={{ background: isChecked ? "#fff" : "#F9FAFB", border: isChecked ? "1px solid #E5E7EB" : "1px solid #E5E7EB", borderRadius:12, padding:"18px 20px", marginBottom:14, boxShadow:"0 1px 4px rgba(0,0,0,0.05)", opacity: isChecked ? 1 : 0.55, transition:"opacity 0.15s" }}>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <label style={{ display:"flex", alignItems:"center", cursor:"pointer" }}>
            <input type="checkbox" checked={!!isChecked} onChange={onToggle} style={{ width:18, height:18, accentColor:"#16A34A", cursor:"pointer", marginRight:4 }} />
          </label>
          <span style={{ width:26, height:26, borderRadius:7, background:"#1B2A4A", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:800 }}>A{index + 1}</span>
          <Tag type={item.dimension} />
          <span style={{ color:"#6B7280", fontSize:12 }}>{item.angle}</span>
        </div>
        <CopyBtn text={full} label="複製全組" />
      </div>
      <div style={{ marginBottom:10 }}>
        <div style={{ fontSize:10, color:"#9CA3AF", fontWeight:700, textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:5 }}>標題</div>
        <div style={{ fontSize:13, fontWeight:700, color:"#111827", marginBottom:3 }}>{item.title_en}</div>
        <div style={{ fontSize:12, color:"#6B7280" }}>{item.title_zh}</div>
      </div>
      <div style={{ marginBottom:10 }}>
        <div style={{ fontSize:10, color:"#9CA3AF", fontWeight:700, textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:5 }}>關鍵詞</div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:4, marginBottom:4 }}>
          {(item.keywords_en||"").split(",").map((k, i) => <span key={i} style={{ background:"#EEF2FF", color:"#4F46E5", fontSize:11, padding:"2px 7px", borderRadius:4 }}>{k.trim()}</span>)}
        </div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
          {(item.keywords_zh||"").split("，").map((k, i) => <span key={i} style={{ background:"#FFF8F0", color:"#D97706", fontSize:11, padding:"2px 7px", borderRadius:4 }}>{k.trim()}</span>)}
        </div>
      </div>
      <div style={{ marginBottom:10 }}>
        <div style={{ fontSize:10, color:"#9CA3AF", fontWeight:700, textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:5 }}>屬性</div>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(190px, 1fr))", gap:5 }}>
          {(item.attributes||[]).map((a, i) => (
            <div key={i} style={{ background:"#F9FAFB", border:"1px solid #E5E7EB", borderRadius:6, padding:"5px 9px", fontSize:12, display:"flex", justifyContent:"space-between", gap:8 }}>
              <span style={{ color:"#9CA3AF" }}>{a.name}</span>
              <span style={{ color:"#111827", fontWeight:700 }}>{a.value}</span>
            </div>
          ))}
        </div>
      </div>
      <div>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:5 }}>
          <div style={{ fontSize:10, color:"#9CA3AF", fontWeight:700, textTransform:"uppercase", letterSpacing:"0.05em" }}>圖片 Prompt</div>
          <CopyBtn text={item.image_prompt || ""} label="複製Prompt" />
        </div>
        <div style={{ background:"#F8F9FF", border:"1px solid #E0E7FF", borderRadius:8, padding:"9px 11px", fontSize:11, color:"#4B5563", lineHeight:1.7, fontFamily:"monospace" }}>{item.image_prompt}</div>
      </div>
    </div>
  );
}

// ─── API 呼叫（透過 Cloudflare Worker → Gemini）────────────
const WORKER_URL = "https://listing-expander-api.alibaba-amy-h.workers.dev";

async function callLLM(prompt, imageBase64, imageType, maxTokens) {
  var tokens = maxTokens || 2000;
  var res = await fetch(WORKER_URL, {
    method:"POST",
    headers:{ "Content-Type":"application/json" },
    body: JSON.stringify({ prompt, imageBase64, imageType, maxTokens: tokens })
  });
  var data = await res.json();
  if (data.error) throw new Error(data.error.message || "API 錯誤");
  var text = (data.content || []).map(function(x) { return x.text || ""; }).join("") || "";
  var clean = text.replace(/```json|```/g, "").trim();
  var start = clean.indexOf("[");
  if (start > 0) clean = clean.slice(start);
  if (!clean.endsWith("]") && clean.includes("[")) {
    var lastBrace = clean.lastIndexOf("}");
    clean = lastBrace > 0 ? clean.slice(0, lastBrace + 1) + "]" : clean + "]";
  }
  return clean;
}

function buildListingPrompt(pName, pDesc, pAdv, pMkt, selected, companyCtx) {
  return [
    "你是阿里巴巴國際站電商專家，幫助B2B商家優化產品listing以避免重複鋪貨。\n",
    companyCtx ? "【公司背景資訊（請納入考量）】\n" + companyCtx + "\n" : "",
    "商家產品資訊：\n",
    "- 產品名稱：" + pName + "\n",
    "- 產品描述：" + (pDesc || "（未提供）") + "\n",
    "- 商家自填優勢/特色：" + (pAdv || "（未提供）") + "\n",
    "- 目標市場：" + pMkt + "\n",
    selected ? "- 商家確認具備的優勢與賣點：" + selected + "\n" : "",
    "\n【屬性設計規則】\n",
    "先定義6-8個固定屬性名稱（英文），10組中完全一致，只有屬性值隨角度變化。\n",
    "錯誤示範：A1用Material，A2用Main Material 禁止\n",
    "\n請根據5個差異化維度生成10組Listing（每維度2組）：\n",
    "維度：應用場景、買家身份、規格重點、地區市場、賣點角度\n",
    "\n嚴格規則：\n",
    "1. 所有內容必須基於商家實際資訊，不可捏造不存在的認證或特性\n",
    "2. 10組標題、關鍵詞必須明顯不同\n",
    "3. 屬性名稱10組完全一致，屬性值隨角度切換\n",
    "4. 屬性名稱與屬性值全部英文\n",
    selected ? "5. 只使用商家確認具備的賣點\n" : "",
    "\n每組包含：dimension、angle（10字內）、title_en（150字元內）、title_zh、keywords_en（5-8個逗號分隔）、keywords_zh（5-8個頓號分隔）、attributes（name+value 6-8個）、image_prompt（英文攝影風格，900x900 pixels square format）\n",
    "\n只回傳JSON陣列，不要任何其他文字：\n",
    '[{"dimension":"應用場景","angle":"工廠自動化","title_en":"...","title_zh":"...","keywords_en":"...","keywords_zh":"...","attributes":[{"name":"Material","value":"..."}],"image_prompt":"..., 900x900 pixels, square format"}]'
  ].join("");
}

function buildDetailPrompt(pName, pDesc, pAdv, pSpecs, pPkg, pMoq, pLeadtime, pCustomize, pPayment, pBuyerType, pMkt, selected, companyCtx) {
  return [
    "你是阿里巴巴國際站資深電商文案專家，請為以下產品撰寫一份完整的英文商品詳情頁內容。\n",
    companyCtx ? "【公司背景資訊】\n" + companyCtx + "\n" : "",
    "【產品資訊】\n",
    "- 產品名稱：" + pName + "\n",
    "- 產品描述：" + (pDesc || "未提供") + "\n",
    "- 規格尺寸：" + (pSpecs || "未提供") + "\n",
    "- 包裝方式：" + (pPkg || "未提供") + "\n",
    "- MOQ：" + (pMoq || "未提供") + "\n",
    "- 交期：" + (pLeadtime || "未提供") + "\n",
    "- 客製化服務：" + (pCustomize || "未提供") + "\n",
    "- 付款方式：" + (pPayment || "未提供") + "\n",
    "- 目標買家類型：" + (pBuyerType || "未提供") + "\n",
    "- 目標市場：" + pMkt + "\n",
    "- 核心優勢與賣點：" + (selected || pAdv || "未提供") + "\n",
    "\n【輸出規則 - 嚴格遵守】\n",
    "1. 所有內容必須基於商家實際提供資訊，不可捏造數字、認證或客戶案例\n",
    "2. 未提供的欄位用 [Please fill in: ___] 標示\n",
    "3. 禁止在輸出內容中提及 Gold Supplier / Verified Supplier / 平台年資\n",
    "4. 禁止出現任何 Emoji 或顏文字\n",
    "5. 全程英文輸出\n",
    "\n請依照以下七個區塊結構輸出：\n",
    "\n块① Product Title\n公式：核心賣點 + 核心關鍵字 + 材質/特性 + 目標用途，目標長度90-110字元\n",
    "\n块② Keywords\n5-10個，逗號分隔\n",
    "\n块③ Product Description（2000字元內）\nCVP核心價值主張 / KSP關鍵賣點矩陣（FABE）/ ICP理想客戶畫像 / TSP信任保障 / CTA行動號召\n",
    "\n块④ Specifications\n表格格式，包含材質、尺寸、重量、顏色、MOQ、認證等\n",
    "\n块⑤ Company Profile（2000字元內）\nP-C-S-A模型：Positioning / Credibility / Strength / Action\n",
    "\n块⑥ FAQ（6-8題）\nQ: 買家視角英文問題 / A: 簡潔回答不超過3句\n",
    "\n块⑦ Information Gap List\n列出仍缺少的關鍵資訊"
  ].join("");
}

function exportToExcel(allResults, filename, extraData) {
  var ed = extraData || {};
  var placeOfOrigin = ed.placeOfOrigin || "";
  var brandName = ed.brandName || "";
  var category = ed.category || "";
  var unit = ed.unit || "Pieces";
  var currency = ed.currency || "USD";
  var pricingType = ed.pricingType || "Tiered pricing";
  var inventory = ed.inventory || "";
  var spuId = ed.spuId || "";
  var tieredPrices = ed.tieredPrices || [{price:"",moq:""},{price:"",moq:""},{price:"",moq:""}];
  var grossWeight = ed.grossWeight || "";
  var itemLength = ed.itemLength || "";
  var itemWidth = ed.itemWidth || "";
  var itemHeight = ed.itemHeight || "";
  var hsCode = ed.hsCode || "";
  var shippingTemplate = ed.shippingTemplate || "";
  var shippingLeadtime = ed.shippingLeadtime || "";
  var logisticsAttr = ed.logisticsAttr || "Ordinary goods";

  var maxAttr = 0;
  allResults.forEach(function(p) {
    (p.listings || []).forEach(function(r) {
      if ((r.attributes || []).length > maxAttr) maxAttr = r.attributes.length;
    });
  });

  var attrHeaders = [];
  for (var i = 1; i <= maxAttr; i++) {
    attrHeaders.push("Product attribute name " + i);
    attrHeaders.push("Product attribute value " + i);
  }

  var refHeaders = ["[參考] 產品名稱","[參考] 標題(EN)","[參考] 標題(中)","[參考] 關鍵詞(EN)","[參考] 關鍵詞(中)","[參考] 圖片Prompt"];
  var headers = ["* Product title","* Product image 1","Product image 2","Product image 3","Product image 4","Product image 5","Product image 6","Product description","Place of origin","Brand name","Category"].concat(attrHeaders).concat(["Sell product by","Batch quantity","Unit","Currency","Pricing type","Inventory","SPU ID (Model number)","SKU code","SKU pricing","Tiered pricing 1","MOQ 1","Tiered pricing 2","MOQ 2","Tiered pricing 3","MOQ 3","* Gross weight (KG)","* Shipping template name","* Shipping quantity 1","* Estimated shipping lead time 1","* Item length (cm)","* Item width (cm)","* Item height (cm)","* Logistics attributes","HS Code"]).concat(refHeaders);

  var rows = [headers];
  allResults.forEach(function(p) {
    (p.listings || []).forEach(function(r, ri) {
      var attrCells = [];
      for (var ai = 0; ai < maxAttr; ai++) {
        var a = (r.attributes || [])[ai];
        attrCells.push(a ? a.name : "");
        attrCells.push(a ? a.value : "");
      }
      var t1 = tieredPrices[0] || {};
      var t2 = tieredPrices[1] || {};
      var t3 = tieredPrices[2] || {};
      var row = [r.title_en,"","","","","","","",placeOfOrigin,brandName,category].concat(attrCells).concat([
        "unit","",unit,currency,pricingType,inventory,
        spuId ? spuId + "-" + String(ri+1).padStart(3,"0") : "","",
        pricingType === "SKU pricing" ? (ed.skuPrice || "") : "",
        t1.price||"",t1.moq||"",t2.price||"",t2.moq||"",t3.price||"",t3.moq||"",
        grossWeight,shippingTemplate,"1",shippingLeadtime||"",
        itemLength,itemWidth,itemHeight,logisticsAttr,hsCode,
        p.productName||r.title_en,r.title_en,r.title_zh,r.keywords_en,r.keywords_zh,r.image_prompt
      ]);
      rows.push(row);
    });
  });

  var ws = XLSX.utils.aoa_to_sheet(rows);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Alibaba default template");
  XLSX.writeFile(wb, filename);
}

// ════════════════════════════════════════════════════════
// MAIN APP
// ════════════════════════════════════════════════════════
export default function App() {
  var IS = useState;
  var IR = useRef;

  // API

  // Tabs
  const [activeTab, setActiveTab] = IS("company");
  const [companyUnlocked, setCompanyUnlocked] = IS(false);

  // 公司資料
  const [companyName, setCompanyName] = IS("");
  const [website, setWebsite] = IS("");
  const [productCats, setProductCats] = IS([""]);
  const [keywords, setKeywords] = IS("");
  const [markets, setMarkets] = IS({});
  const [marketOther, setMarketOther] = IS("");
  const [marketOtherChecked, setMarketOtherChecked] = IS(false);
  const [advantages, setAdvantages] = IS({});
  const [advantageOther, setAdvantageOther] = IS("");
  const [advantageOtherChecked, setAdvantageOtherChecked] = IS(false);
  const [identity, setIdentity] = IS("");
  const [customService, setCustomService] = IS("");
  const [sampleService, setSampleService] = IS("");
  const [pricing, setPricing] = IS({});
  const [styleChoice, setStyleChoice] = IS("現代簡約");
  const [colors, setColors] = IS({ bg:"#ffffff", primary:"#ff6b35", text:"#1e293b", accent:"#ec4899" });

  // 工廠
  const [factoryProcess, setFactoryProcess] = IS("");
  const [factoryExhibition, setFactoryExhibition] = IS("");
  const [factoryTeam, setFactoryTeam] = IS("");
  const [factoryRD, setFactoryRD] = IS({});
  const [factoryRDOther, setFactoryRDOther] = IS("");
  const [factoryRDOtherChecked, setFactoryRDOtherChecked] = IS(false);
  const [factoryCert, setFactoryCert] = IS({});
  const [factoryCertOther, setFactoryCertOther] = IS("");
  const [factoryCertOtherChecked, setFactoryCertOtherChecked] = IS(false);
  const [factoryOrderFlex, setFactoryOrderFlex] = IS("");
  const [factoryBuyer, setFactoryBuyer] = IS({});
  const [factoryBuyerOther, setFactoryBuyerOther] = IS("");
  const [factoryBuyerOtherChecked, setFactoryBuyerOtherChecked] = IS(false);

  // 貿易商
  const [tradingService, setTradingService] = IS({});
  const [tradingServiceOther, setTradingServiceOther] = IS("");
  const [tradingServiceOtherChecked, setTradingServiceOtherChecked] = IS(false);
  const [tradingTeam, setTradingTeam] = IS({});
  const [tradingTeamOther, setTradingTeamOther] = IS("");
  const [tradingTeamOtherChecked, setTradingTeamOtherChecked] = IS(false);
  const [tradingBuyer, setTradingBuyer] = IS({});
  const [tradingBuyerOther, setTradingBuyerOther] = IS("");
  const [tradingBuyerOtherChecked, setTradingBuyerOtherChecked] = IS(false);

  // 旺舖輸出
  const [wangpuPrompt, setWangpuPrompt] = IS("");

  // Listing
  const [lStep, setLStep] = IS(1);
  const [productName, setProductName] = IS("");
  const [description, setDescription] = IS("");
  const [prodAdvantages, setProdAdvantages] = IS("");
  const [lMarket, setLMarket] = IS("全球通用");
  const [imageFile, setImageFile] = IS(null);
  const [imagePreview, setImagePreview] = IS(null);
  const [imageBase64, setImageBase64] = IS(null);
  const [suggestions, setSuggestions] = IS(null);
  const [checked, setChecked] = IS({});
  const [lLoading1, setLLoading1] = IS(false);
  const [lLoading2, setLLoading2] = IS(false);
  const [lError, setLError] = IS("");
  const [exportMsg, setExportMsg] = IS("");
  const [detailPage, setDetailPage] = IS("");
  const [showDetail, setShowDetail] = IS(false);

  // 累積批次：每個 batch = { id, productName, round, listings, checkedMap, collapsed }
  const [batches, setBatches] = IS([]);
  const [expandMore, setExpandMore] = IS(false);
  const [moreNote, setMoreNote] = IS("");
  const [moreLoading, setMoreLoading] = IS(false);

  // 商品詳情頁欄位
  const [prodSpecs, setProdSpecs] = IS("");
  const [prodPackaging, setProdPackaging] = IS("");
  const [prodMoq, setProdMoq] = IS("");
  const [prodLeadtime, setProdLeadtime] = IS("");
  const [prodCustomize, setProdCustomize] = IS("");
  const [prodPayment, setProdPayment] = IS("");
  const [prodBuyerType, setProdBuyerType] = IS("");

  // 阿里上架欄位
  const [placeOfOrigin, setPlaceOfOrigin] = IS("");
  const [brandName, setBrandName] = IS("");
  const [category, setCategory] = IS("");
  const [unit, setUnit] = IS("Pieces");
  const [currency, setCurrency] = IS("USD");
  const [pricingType, setPricingType] = IS("Tiered pricing");
  const [inventory, setInventory] = IS("");
  const [spuId, setSpuId] = IS("");
  const [skuPrice, setSkuPrice] = IS("");
  const [tieredPrices, setTieredPrices] = IS([{price:"",moq:""},{price:"",moq:""},{price:"",moq:""}]);
  const [grossWeight, setGrossWeight] = IS("");
  const [itemLength, setItemLength] = IS("");
  const [itemWidth, setItemWidth] = IS("");
  const [itemHeight, setItemHeight] = IS("");
  const [hsCode, setHsCode] = IS("");
  const [shippingTemplate, setShippingTemplate] = IS("");
  const [shippingLeadtime, setShippingLeadtime] = IS("");
  const [logisticsAttr, setLogisticsAttr] = IS("Ordinary goods");

  const fileRef = IR();

  // ── helpers ──
  const inp = { width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", color:"#1e293b", fontFamily:"inherit" };
  const tog = (obj, setObj, key) => setObj(p => Object.assign({}, p, { [key]: !p[key] }));

  const buildCtx = () => {
    if (!companyUnlocked) return "";
    var mktList = Object.keys(markets).filter(k => markets[k]);
    if (marketOtherChecked && marketOther.trim()) mktList.push(marketOther.trim());
    var advList = Object.keys(advantages).filter(k => advantages[k]);
    if (advantageOtherChecked && advantageOther.trim()) advList.push(advantageOther.trim());
    var certCtxList = Object.keys(factoryCert).filter(k => factoryCert[k]);
    if (factoryCertOtherChecked && factoryCertOther.trim()) certCtxList.push(factoryCertOther.trim());
    return [
      companyName && "公司名稱：" + companyName,
      identity && "身份定位：" + identity,
      productCats.filter(Boolean).length && "主營產品：" + productCats.filter(Boolean).join("、"),
      mktList.length && "目標市場：" + mktList.join("、"),
      advList.length && "公司核心優勢：" + advList.join("、"),
      certCtxList.length && "公司認證：" + certCtxList.join("、"),
    ].filter(Boolean).join("\n");
  };

  const genWangpu = () => {
    if (!companyName.trim()) return;
    var mktList = Object.keys(markets).filter(k => markets[k]);
    if (marketOtherChecked && marketOther.trim()) mktList.push(marketOther.trim());
    var mkt = mktList.join("、");
    var advList = Object.keys(advantages).filter(k => advantages[k]);
    if (advantageOtherChecked && advantageOther.trim()) advList.push(advantageOther.trim());
    var adv = advList.join("、");
    var cats = productCats.filter(Boolean);
    var pricingStr = Object.keys(pricing).filter(k => pricing[k]).join("、");
    var identityDetails = "";
    if (identity === "工廠直營" || identity === "工貿一體") {
      var rdL = Object.keys(factoryRD).filter(k => factoryRD[k]);
      if (factoryRDOtherChecked && factoryRDOther.trim()) rdL.push(factoryRDOther.trim());
      var certL = Object.keys(factoryCert).filter(k => factoryCert[k]);
      if (factoryCertOtherChecked && factoryCertOther.trim()) certL.push(factoryCertOther.trim());
      var buyerL = Object.keys(factoryBuyer).filter(k => factoryBuyer[k]);
      if (factoryBuyerOtherChecked && factoryBuyerOther.trim()) buyerL.push(factoryBuyerOther.trim());
      identityDetails = "生產流程：" + (factoryProcess || "未填") + "\n研發技術：" + (rdL.join("、") || "未選") + "\n品質認證：" + (certL.join("、") || "未選") + "\n展會：" + (factoryExhibition || "未填") + "\n團隊：" + (factoryTeam || "未填") + "\n訂單彈性：" + (factoryOrderFlex || "未填") + "\n目標買家：" + (buyerL.join("、") || "未選");
    } else if (identity === "貿易商") {
      var svcL = Object.keys(tradingService).filter(k => tradingService[k]);
      if (tradingServiceOtherChecked && tradingServiceOther.trim()) svcL.push(tradingServiceOther.trim());
      var tmL = Object.keys(tradingTeam).filter(k => tradingTeam[k]);
      if (tradingTeamOtherChecked && tradingTeamOther.trim()) tmL.push(tradingTeamOther.trim());
      var buyerTL = Object.keys(tradingBuyer).filter(k => tradingBuyer[k]);
      if (tradingBuyerOtherChecked && tradingBuyerOther.trim()) buyerTL.push(tradingBuyerOther.trim());
      identityDetails = "服務加值：" + (svcL.join("、") || "未選") + "\n團隊服務力：" + (tmL.join("、") || "未選") + "\n目標買家：" + (buyerTL.join("、") || "未選");
    }
    var prompt = [
      "你現在是一名阿里巴巴國際站視覺研究顧問 + B2B採購行為專家。\n\n",
      "【第一步：旺舖顧問分析】\n",
      "請依據以下商家資訊分析市場買方需求和競業狀況：\n",
      "• 公司名稱：" + companyName + (website ? "\n• 官網：" + website : "") + "\n",
      "• 主營產品：" + cats.join("、") + "\n",
      keywords ? "• 產品關鍵詞：" + keywords + "\n" : "",
      "• 目標市場：" + (mkt || "未選") + "\n",
      "• 核心優勢：" + (adv || "未選") + "\n",
      "• 定製服務：" + (customService || "未選") + "　樣品服務：" + (sampleService || "未選") + "\n",
      "• 身份定位：" + (identity || "未選") + "\n",
      "• 定位細節：\n" + identityDetails + "\n",
      "• 定價策略：" + (pricingStr || "未選") + "\n\n",
      "【第二步：旺舖OBS文案生成（中英對照）】\n",
      "(1) 關於我們 / About Us\n(2) 產品分類 / Product Categories\n(3) 公司優勢 / Why Choose Us\n",
      "(4) 生產流程 / Production Process（工廠/工貿適用）\n(5) 認證 / Certificates\n(6) 展會 / Exhibitions\n(7) 團隊 / Our Team\n(8) 項目案例 / Project Examples\n\n",
      "【第三步：商品明細頁建議】\n",
      "針對主營類目：" + cats.join("、") + "\n",
      "分別提供模板A（B2B OEM/ODM客製化）和模板B（RTS現貨/批發商）的商品詳情頁建議。\n\n",
      "【第四步：旺舖簡報結構（NotebookLM Resource）】\n",
      "視覺風格：" + styleChoice + "\n",
      "色彩系統：背景" + colors.bg + "，主色" + colors.primary + "，文字" + colors.text + "，強調色" + colors.accent + "\n",
      "請產出可放入NotebookLM的完整英文簡報結構。"
    ].join("");
    setWangpuPrompt(prompt);
    setCompanyUnlocked(true);
  };

  const analyzeProduct = async () => {
    if (!productName.trim()) { setLError("請填寫產品名稱"); return; }
    setLError(""); setLLoading1(true);
    var ctx = buildCtx();
    var prompt = [
      "你是B2B跨境電商專家。\n",
      ctx ? "公司背景：\n" + ctx + "\n" : "",
      "產品名稱：" + productName + "，描述：" + (description || "未提供") + "，優勢：" + (prodAdvantages || "未提供") + "，市場：" + lMarket + "\n\n",
      "請列出此產品類別通常應具備的項目，分三類，每類3-8個具體項目。\n",
      "只回傳JSON，不要任何其他文字：\n",
      '{"industry_summary":"一句話說明產業和主要買家","advantages":[{"id":"a1","label":"可量身客製化規格","hint":"說明"}],"certifications":[{"id":"c1","label":"CE認證","hint":"說明"}],"selling_points":[{"id":"s1","label":"工廠直供","hint":"說明"}]}'
    ].join("");
    try {
      var rawText = await callLLM(prompt, imageBase64, imageFile ? imageFile.type : null, 2000);
      var parsed = JSON.parse(rawText.replace(/```json|```/g, "").trim());
      setSuggestions(parsed);
      var init = {};
      [].concat(parsed.advantages || [], parsed.certifications || [], parsed.selling_points || []).forEach(function(x) { init[x.id] = true; });
      setChecked(init);
      setLStep(2);
    } catch(e) { setLError("分析失敗：" + e.message); }
    finally { setLLoading1(false); }
  };

  const generateListings = async () => {
    setLError(""); setLLoading2(true); setDetailPage(""); setShowDetail(false);
    var sel = suggestions ? [].concat(
      (suggestions.advantages || []).filter(function(x) { return checked[x.id]; }),
      (suggestions.certifications || []).filter(function(x) { return checked[x.id]; }),
      (suggestions.selling_points || []).filter(function(x) { return checked[x.id]; })
    ).map(function(x) { return x.label; }).join("、") : "";
    var ctx = buildCtx();
    try {
      var listingPrompt = buildListingPrompt(productName, description, prodAdvantages, lMarket, sel, ctx);
      var detailPrompt = buildDetailPrompt(productName, description, prodAdvantages, prodSpecs, prodPackaging, prodMoq, prodLeadtime, prodCustomize, prodPayment, prodBuyerType, lMarket, sel, ctx);
      var results = await Promise.all([
        callLLM(listingPrompt, imageBase64, imageFile ? imageFile.type : null, 8000),
        callLLM(detailPrompt, imageBase64, imageFile ? imageFile.type : null, 6000).catch(function() { return "（商品詳情頁生成失敗）"; })
      ]);
      var newListings = JSON.parse(results[0]);
      var checkedMap = {};
      newListings.forEach(function(_, i) { checkedMap[i] = true; });
      var newBatch = {
        id: Date.now() + "-" + Math.random().toString(36).slice(2,7),
        productName: productName,
        round: 1,
        listings: newListings,
        checkedMap: checkedMap,
        collapsed: false,
        detailPage: typeof results[1] === "string" ? results[1] : "",
        showDetail: false
      };
      setBatches(function(prev) { return prev.concat([newBatch]); });
      setLStep(3);
    } catch(e) { setLError("生成失敗：" + e.message); }
    finally { setLLoading2(false); }
  };

  // 再擴10組：基於目前/調整後的產品資訊，對同一產品再生成一輪，疊加進 batches
  const generateMoreForSameProduct = async () => {
    setLError(""); setMoreLoading(true);
    var sel = suggestions ? [].concat(
      (suggestions.advantages || []).filter(function(x) { return checked[x.id]; }),
      (suggestions.certifications || []).filter(function(x) { return checked[x.id]; }),
      (suggestions.selling_points || []).filter(function(x) { return checked[x.id]; })
    ).map(function(x) { return x.label; }).join("、") : "";
    var ctx = buildCtx();
    var extraAdv = prodAdvantages + (moreNote.trim() ? ("；補充：" + moreNote.trim()) : "");
    try {
      var listingPrompt = buildListingPrompt(productName, description, extraAdv, lMarket, sel, ctx);
      var rawText = await callLLM(listingPrompt, imageBase64, imageFile ? imageFile.type : null, 8000);
      var newListings = JSON.parse(rawText);
      var checkedMap = {};
      newListings.forEach(function(_, i) { checkedMap[i] = true; });
      setBatches(function(prev) {
        var sameProductRounds = prev.filter(function(b) { return b.productName === productName; }).length;
        var newBatch = {
          id: Date.now() + "-" + Math.random().toString(36).slice(2,7),
          productName: productName,
          round: sameProductRounds + 1,
          listings: newListings,
          checkedMap: checkedMap,
          collapsed: false,
          detailPage: "",
          showDetail: false
        };
        return prev.concat([newBatch]);
      });
      setExpandMore(false);
      setMoreNote("");
    } catch(e) { setLError("生成失敗：" + e.message); }
    finally { setMoreLoading(false); }
  };

  const toggleBatchCollapse = (batchId) => {
    setBatches(function(prev) { return prev.map(function(b) { return b.id === batchId ? Object.assign({}, b, { collapsed: !b.collapsed }) : b; }); });
  };

  const toggleBatchDetail = (batchId) => {
    setBatches(function(prev) { return prev.map(function(b) { return b.id === batchId ? Object.assign({}, b, { showDetail: !b.showDetail }) : b; }); });
  };

  const toggleListingChecked = (batchId, idx) => {
    setBatches(function(prev) { return prev.map(function(b) {
      if (b.id !== batchId) return b;
      var nc = Object.assign({}, b.checkedMap);
      nc[idx] = !nc[idx];
      return Object.assign({}, b, { checkedMap: nc });
    }); });
  };

  const getSelectedExportData = () => {
    return batches.map(function(b) {
      var selectedListings = b.listings.filter(function(_, i) { return b.checkedMap[i]; });
      return { productName: b.productName, listings: selectedListings };
    }).filter(function(b) { return b.listings.length > 0; });
  };

  const totalSelectedCount = () => {
    var c = 0;
    batches.forEach(function(b) { Object.keys(b.checkedMap).forEach(function(k) { if (b.checkedMap[k]) c++; }); });
    return c;
  };

  const handleImage = (file) => {
    if (!file) return;
    setImageFile(file);
    var reader = new FileReader();
    reader.onload = function(e) { setImagePreview(e.target.result); setImageBase64(e.target.result.split(",")[1]); };
    reader.readAsDataURL(file);
  };

  const resetProduct = () => {
    setLStep(1); setSuggestions(null); setDetailPage(""); setShowDetail(false);
    setProductName(""); setDescription(""); setProdAdvantages(""); setProdSpecs(""); setProdPackaging("");
    setProdMoq(""); setProdLeadtime(""); setProdCustomize(""); setProdPayment(""); setProdBuyerType("");
    setImageFile(null); setImagePreview(null); setImageBase64(null); setLMarket("全球通用");
    setBatches([]); setExpandMore(false); setMoreNote("");
  };

  const resetProductOnly = () => {
    setLStep(1); setSuggestions(null); setDetailPage(""); setShowDetail(false);
    setProductName(""); setDescription(""); setProdAdvantages(""); setProdSpecs(""); setProdPackaging("");
    setProdMoq(""); setProdLeadtime(""); setProdCustomize(""); setProdPayment(""); setProdBuyerType("");
    setImageFile(null); setImagePreview(null); setImageBase64(null);
    setExpandMore(false); setMoreNote("");
    // batches 保留，新產品的結果會疊加進去
  };

  const getExtraData = () => ({
    placeOfOrigin, brandName, category, unit, currency, pricingType, inventory, spuId, skuPrice,
    tieredPrices, grossWeight, itemLength, itemWidth, itemHeight, hsCode, shippingTemplate, shippingLeadtime, logisticsAttr
  });


      {/* Tabs */}
      <div style={{ background:"#fff", borderBottom:"1px solid #E5E7EB", padding:"0 32px", display:"flex", gap:0 }}>
        {tabs.map(function(t) {
          return (
            <button key={t.key} onClick={() => !t.locked && setActiveTab(t.key)} style={{ padding:"14px 24px", border:"none", background:"none", cursor: t.locked ? "not-allowed" : "pointer", borderBottom: activeTab === t.key ? "3px solid #ff6b35" : "3px solid transparent", display:"flex", alignItems:"center", gap:10, opacity: t.locked ? 0.4 : 1, transition:"all 0.15s" }}>
              <div>
                <div style={{ fontSize:14, fontWeight: activeTab === t.key ? 700 : 500, color: activeTab === t.key ? "#ff6b35" : "#374151" }}>{t.label}</div>
                <div style={{ fontSize:11, color: companyUnlocked && t.key === "listing" ? "#16A34A" : "#9CA3AF" }}>{t.sub}</div>
              </div>
              {t.locked && <span style={{ fontSize:14 }}>🔒</span>}
            </button>
          );
        })}
      </div>

      <div style={{ maxWidth:1300, margin:"0 auto", padding:"28px 24px" }}>

        {/* ══ 公司實力 Tab ══ */}
        {activeTab === "company" && (
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:24 }}>
            <div>
              {/* Stage 1 */}
              <div style={{ background:"#fff", borderRadius:14, padding:24, boxShadow:"0 1px 6px rgba(0,0,0,0.07)", marginBottom:20 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:18, paddingBottom:12, borderBottom:"2px solid #ff6b35" }}>
                  <div style={{ width:28, height:28, borderRadius:"50%", background:"linear-gradient(135deg,#ff6b35,#f7931e)", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800 }}>1</div>
                  <div style={{ fontWeight:700, color:"#1e293b", fontSize:15 }}>基礎資料</div>
                </div>
                <Field label="公司名稱（英文）*">
                  <input value={companyName} onChange={e => setCompanyName(e.target.value)} placeholder="例：Shenzhen Smart Tech Co., Ltd." style={inp} />
                </Field>
                <Field label="官網網址（選填）">
                  <input value={website} onChange={e => setWebsite(e.target.value)} placeholder="https://www.example.com" style={inp} />
                </Field>
                <Field label="主營產品分類 *" hint="每個分類一個輸入框，描述越詳細越好">
                  {productCats.map(function(cat, i) {
                    return (
                      <div key={i} style={{ display:"flex", gap:8, marginBottom:8 }}>
                        <input value={cat} onChange={function(e) { var n = productCats.slice(); n[i] = e.target.value; setProductCats(n); }} placeholder={"例：保養品精華液"} style={{ flex:1, padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit" }} />
                        {productCats.length > 1
                          ? <button onClick={() => setProductCats(productCats.filter(function(_, j) { return j !== i; }))} style={{ padding:"0 10px", background:"#FEE2E2", color:"#EF4444", border:"none", borderRadius:7, cursor:"pointer", fontWeight:700 }}>−</button>
                          : <button onClick={() => setProductCats(productCats.concat([""]))} style={{ padding:"0 10px", background:"#D1FAE5", color:"#059669", border:"none", borderRadius:7, cursor:"pointer", fontWeight:700 }}>+</button>
                        }
                      </div>
                    );
                  })}
                  {productCats.length > 1 && <button onClick={() => setProductCats(productCats.concat([""]))} style={{ fontSize:12, color:"#3B5BDB", background:"none", border:"none", cursor:"pointer", padding:0 }}>+ 新增分類</button>}
                </Field>
                <Field label="產品關鍵詞" hint="逗號分隔">
                  <input value={keywords} onChange={e => setKeywords(e.target.value)} placeholder="例：skin regeneration, anti-aging serum" style={inp} />
                </Field>
                <Field label="目標市場 *">
                  <CheckGroup items={MARKETS} checked={markets} onChange={k => tog(markets, setMarkets, k)} cols={3} />
                  <OtherField checked={marketOtherChecked} onCheck={() => setMarketOtherChecked(function(p) { return !p; })} value={marketOther} onChange={setMarketOther} placeholder="請說明其它市場" />
                </Field>
                <Field label="核心優勢">
                  <CheckGroup items={ADVANTAGES} checked={advantages} onChange={k => tog(advantages, setAdvantages, k)} cols={3} />
                  <OtherField checked={advantageOtherChecked} onCheck={() => setAdvantageOtherChecked(function(p) { return !p; })} value={advantageOther} onChange={setAdvantageOther} placeholder="請說明其它優勢" />
                </Field>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
                  <Field label="定製服務？">
                    {["是","否"].map(function(v) {
                      return (
                        <label key={v} style={{ display:"flex", alignItems:"center", gap:6, cursor:"pointer", fontSize:13, color:"#475569", marginTop:6 }}>
                          <input type="radio" name="custom" value={v} checked={customService === v} onChange={() => setCustomService(v)} style={{ accentColor:"#ff6b35" }} /> {v}
                        </label>
                      );
                    })}
                  </Field>
                  <Field label="樣品服務？">
                    {["是","否"].map(function(v) {
                      return (
                        <label key={v} style={{ display:"flex", alignItems:"center", gap:6, cursor:"pointer", fontSize:13, color:"#475569", marginTop:6 }}>
                          <input type="radio" name="sample" value={v} checked={sampleService === v} onChange={() => setSampleService(v)} style={{ accentColor:"#ff6b35" }} /> {v}
                        </label>
                      );
                    })}
                  </Field>
                </div>
              </div>

              {/* Stage 2 */}
              <div style={{ background:"#fff", borderRadius:14, padding:24, boxShadow:"0 1px 6px rgba(0,0,0,0.07)", marginBottom:20 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:18, paddingBottom:12, borderBottom:"2px solid #ff6b35" }}>
                  <div style={{ width:28, height:28, borderRadius:"50%", background:"linear-gradient(135deg,#ff6b35,#f7931e)", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800 }}>2</div>
                  <div style={{ fontWeight:700, color:"#1e293b", fontSize:15 }}>定位細節</div>
                </div>
                <Field label="身份定位 *">
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginTop:8 }}>
                    {["工廠直營","貿易商","工貿一體","品牌商"].map(function(v) {
                      return (
                        <label key={v} style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer", padding:"9px 12px", borderRadius:8, border:"1.5px solid " + (identity === v ? "#ff6b35" : "#E5E7EB"), background: identity === v ? "#FFF7ED" : "#FAFAFA" }}>
                          <input type="radio" name="identity" value={v} checked={identity === v} onChange={() => setIdentity(v)} style={{ accentColor:"#ff6b35" }} />
                          <span style={{ fontSize:13, fontWeight: identity === v ? 600 : 400, color: identity === v ? "#C2410C" : "#475569" }}>{v}</span>
                        </label>
                      );
                    })}
                  </div>
                </Field>

                {(identity === "工廠直營" || identity === "工貿一體") && (
                  <div>
                    <Field label="生產流程">
                      <textarea value={factoryProcess} onChange={e => setFactoryProcess(e.target.value)} rows={3} placeholder="請描述生產流程..." style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", resize:"vertical" }} />
                    </Field>
                    <Field label="展會經驗" hint="格式：年份-國家-展覽名稱">
                      <textarea value={factoryExhibition} onChange={e => setFactoryExhibition(e.target.value)} rows={2} placeholder="例：2024-德國-法蘭克福展" style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", resize:"vertical" }} />
                    </Field>
                    <Field label="研發與技術">
                      <CheckGroup items={["獨立研發團隊","擁有多項發明專利","10年以上OEM/ODM經驗","內部實驗室","3D快速打樣","每年50+新品"]} checked={factoryRD} onChange={k => tog(factoryRD, setFactoryRD, k)} cols={2} />
                      <OtherField checked={factoryRDOtherChecked} onCheck={() => setFactoryRDOtherChecked(function(p) { return !p; })} value={factoryRDOther} onChange={setFactoryRDOther} placeholder="請說明其它研發能力" />
                    </Field>
                    <Field label="品質認證">
                      <CheckGroup items={["ISO 9001","BSCI","SEDEX","CE / RoHS","UL","100% QC全檢"]} checked={factoryCert} onChange={k => tog(factoryCert, setFactoryCert, k)} cols={3} />
                      <OtherField checked={factoryCertOtherChecked} onCheck={() => setFactoryCertOtherChecked(function(p) { return !p; })} value={factoryCertOther} onChange={setFactoryCertOther} placeholder="請說明其它認證" />
                    </Field>
                    <Field label="訂單彈性">
                      <input value={factoryOrderFlex} onChange={e => setFactoryOrderFlex(e.target.value)} placeholder="MOQ、出貨天數等" style={inp} />
                    </Field>
                    <Field label="目標買家">
                      <CheckGroup items={["全球大品牌商","專業批發商","線下連鎖超市","工業工程承包商"]} checked={factoryBuyer} onChange={k => tog(factoryBuyer, setFactoryBuyer, k)} cols={2} />
                      <OtherField checked={factoryBuyerOtherChecked} onCheck={() => setFactoryBuyerOtherChecked(function(p) { return !p; })} value={factoryBuyerOther} onChange={setFactoryBuyerOther} placeholder="請說明其它買家類型" />
                    </Field>
                  </div>
                )}

                {identity === "貿易商" && (
                  <div>
                    <Field label="服務加值">
                      <CheckGroup items={["海外倉儲","多品類拼櫃","跨境Dropshipping","包裝視覺設計","第三方驗貨"]} checked={tradingService} onChange={k => tog(tradingService, setTradingService, k)} cols={2} />
                      <OtherField checked={tradingServiceOtherChecked} onCheck={() => setTradingServiceOtherChecked(function(p) { return !p; })} value={tradingServiceOther} onChange={setTradingServiceOther} placeholder="請說明其它服務加值" />
                    </Field>
                    <Field label="團隊服務力">
                      <CheckGroup items={["24h在線響應","多語種團隊","海外在地化售後"]} checked={tradingTeam} onChange={k => tog(tradingTeam, setTradingTeam, k)} cols={2} />
                      <OtherField checked={tradingTeamOtherChecked} onCheck={() => setTradingTeamOtherChecked(function(p) { return !p; })} value={tradingTeamOther} onChange={setTradingTeamOther} placeholder="請說明其它服務能力" />
                    </Field>
                    <Field label="目標買家">
                      <CheckGroup items={["海外批發商/分銷商","品牌商/零售商","電商賣家/Amazon","工程項目採購"]} checked={tradingBuyer} onChange={k => tog(tradingBuyer, setTradingBuyer, k)} cols={2} />
                      <OtherField checked={tradingBuyerOtherChecked} onCheck={() => setTradingBuyerOtherChecked(function(p) { return !p; })} value={tradingBuyerOther} onChange={setTradingBuyerOther} placeholder="請說明其它買家類型" />
                    </Field>
                  </div>
                )}

                <Field label="定價策略">
                  <CheckGroup items={["高性價比","高端定制"]} checked={pricing} onChange={k => tog(pricing, setPricing, k)} cols={2} />
                </Field>
              </div>

              {/* Stage 3 */}
              <div style={{ background:"#fff", borderRadius:14, padding:24, boxShadow:"0 1px 6px rgba(0,0,0,0.07)", marginBottom:20 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:18, paddingBottom:12, borderBottom:"2px solid #ff6b35" }}>
                  <div style={{ width:28, height:28, borderRadius:"50%", background:"linear-gradient(135deg,#ff6b35,#f7931e)", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800 }}>3</div>
                  <div style={{ fontWeight:700, color:"#1e293b", fontSize:15 }}>風格設計</div>
                </div>
                <Field label="旺舖風格">
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(3, 1fr)", gap:8, marginTop:8 }}>
                    {STYLES.map(function(s) {
                      return (
                        <label key={s} style={{ display:"flex", alignItems:"center", gap:6, cursor:"pointer", padding:"7px 10px", borderRadius:7, border:"1.5px solid " + (styleChoice === s ? "#ff6b35" : "#E5E7EB"), background: styleChoice === s ? "#FFF7ED" : "#FAFAFA" }}>
                          <input type="radio" name="style" value={s} checked={styleChoice === s} onChange={() => setStyleChoice(s)} style={{ accentColor:"#ff6b35" }} />
                          <span style={{ fontSize:12, fontWeight: styleChoice === s ? 600 : 400, color: styleChoice === s ? "#C2410C" : "#475569" }}>{s}</span>
                        </label>
                      );
                    })}
                  </div>
                </Field>
                <Field label="快速套用配色">
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(4, 1fr)", gap:6, marginTop:8 }}>
                    {PRESETS.map(function(p) {
                      return <button key={p.label} onClick={() => setColors({ bg:p.bg, primary:p.primary, text:p.text, accent:p.accent })} style={{ padding:"6px", borderRadius:7, border:"1.5px solid #E5E7EB", background:"#FAFAFA", fontSize:11, cursor:"pointer", color:"#374151" }}>{p.label}</button>;
                    })}
                  </div>
                </Field>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                  {[["bg","背景色"],["primary","主色調"],["text","文字色"],["accent","強調色"]].map(function(pair) {
                    return (
                      <div key={pair[0]}>
                        <div style={{ fontSize:12, color:"#6B7280", marginBottom:4 }}>{pair[1]}</div>
                        <input type="color" value={colors[pair[0]]} onChange={function(e) { var v = e.target.value; setColors(function(prev) { return Object.assign({}, prev, { [pair[0]]: v }); }); }} style={{ width:"100%", height:38, borderRadius:7, border:"1.5px solid #E5E7EB", cursor:"pointer", padding:2 }} />
                      </div>
                    );
                  })}
                </div>
              </div>

              <button onClick={genWangpu} disabled={!companyName.trim()} style={{ width:"100%", padding:"14px", background: companyName.trim() ? "linear-gradient(135deg,#ff6b35,#f7931e)" : "#E5E7EB", color: companyName.trim() ? "#fff" : "#9CA3AF", border:"none", borderRadius:11, fontSize:16, fontWeight:800, cursor: companyName.trim() ? "pointer" : "not-allowed", boxShadow: companyName.trim() ? "0 4px 14px rgba(255,107,53,0.35)" : "none" }}>
                🚀 生成旺舖 AI Prompt　→　同時解鎖 Listing 擴張器
              </button>
            </div>

            {/* 右側：Prompt 輸出 */}
            <div>
              <div style={{ fontWeight:700, color:"#1e293b", fontSize:15, marginBottom:14 }}>💻 您的 AI 專屬指令</div>
              <div style={{ background:"linear-gradient(135deg,#1e293b,#0f172a)", borderRadius:12, padding:24, minHeight:400, position:"relative" }}>
                <div style={{ position:"absolute", top:0, left:0, right:0, height:32, background:"#334155", borderRadius:"12px 12px 0 0", display:"flex", alignItems:"center", padding:"0 14px", gap:6 }}>
                  <div style={{ width:10, height:10, borderRadius:"50%", background:"#ef4444" }} />
                  <div style={{ width:10, height:10, borderRadius:"50%", background:"#f59e0b" }} />
                  <div style={{ width:10, height:10, borderRadius:"50%", background:"#10b981" }} />
                  <span style={{ fontSize:11, color:"#94a3b8", marginLeft:8 }}>AI Prompt</span>
                  {wangpuPrompt && <div style={{ marginLeft:"auto" }}><CopyBtn text={wangpuPrompt} label="📋 複製指令" /></div>}
                </div>
                <div style={{ marginTop:32, fontFamily:"Courier New, monospace", fontSize:12, lineHeight:1.8, color: wangpuPrompt ? "#e2e8f0" : "#64748b", whiteSpace:"pre-wrap", wordBreak:"break-word" }}>
                  {wangpuPrompt || "填寫左側資料後點擊「生成旺舖 AI Prompt」\n系統將自動生成您的專屬指令...\n\n生成後可複製貼入 ChatGPT / Gemini / Claude 使用。"}
                </div>
              </div>
              {wangpuPrompt && (
                <div style={{ background:"#F0FDF4", border:"1px solid #BBF7D0", borderRadius:10, padding:"12px 16px", marginTop:14 }}>
                  <div style={{ fontSize:13, color:"#16A34A", fontWeight:700, marginBottom:4 }}>✓ 旺舖 Prompt 已生成，Listing 擴張器已解鎖</div>
                  <div style={{ fontSize:12, color:"#166534" }}>切換到上方「⚡ Listing 擴張器」Tab，公司資料將自動帶入每組 Listing 生成。</div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ══ Listing Tab ══ */}
        {activeTab === "listing" && (
          <div>
            {companyUnlocked && (
              <div style={{ background:"#EEF2FF", border:"1px solid #C7D7FE", borderRadius:10, padding:"10px 16px", marginBottom:20, display:"flex", alignItems:"center", gap:10 }}>
                <span style={{ fontSize:16 }}>🏢</span>
                <div>
                  <span style={{ fontSize:13, fontWeight:700, color:"#3B5BDB" }}>公司資料已自動帶入：</span>
                  <span style={{ fontSize:12, color:"#4B5563", marginLeft:6 }}>{companyName}｜{Object.keys(markets).filter(function(k) { return markets[k]; }).join("、") || "未設定市場"}｜{Object.keys(factoryCert).filter(function(k) { return factoryCert[k]; }).join("、") || "未設定認證"}</span>
                </div>
              </div>
            )}

            <StepBar step={lStep} />

            {lStep === 1 && (
              <div style={{ display:"flex", gap:24 }}>
                <div style={{ width:420, flexShrink:0 }}>
                  <div style={{ background:"#fff", borderRadius:14, padding:24, boxShadow:"0 1px 6px rgba(0,0,0,0.07)" }}>
                    <div style={{ fontWeight:700, color:"#1B2A4A", fontSize:15, marginBottom:18 }}>產品資訊</div>
                    <Field label="產品名稱 *">
                      <input value={productName} onChange={e => setProductName(e.target.value)} placeholder="例：肌膚再生精華液" style={inp} />
                    </Field>
                    <Field label="產品圖片（可選）">
                      <div onClick={() => fileRef.current.click()} onDragOver={e => e.preventDefault()} onDrop={function(e) { e.preventDefault(); handleImage(e.dataTransfer.files[0]); }}
                        style={{ border:"2px dashed #D1D5DB", borderRadius:10, padding: imagePreview ? 0 : "20px 12px", textAlign:"center", cursor:"pointer", background:"#FAFAFA", overflow:"hidden" }}>
                        {imagePreview
                          ? (
                            <div style={{ position:"relative" }}>
                              <img src={imagePreview} alt="" style={{ width:"100%", maxHeight:130, objectFit:"cover", display:"block" }} />
                              <div onClick={function(e) { e.stopPropagation(); setImageFile(null); setImagePreview(null); setImageBase64(null); }} style={{ position:"absolute", top:6, right:6, background:"rgba(0,0,0,0.5)", borderRadius:6, padding:"2px 8px", color:"#fff", fontSize:11, cursor:"pointer" }}>移除</div>
                            </div>
                          )
                          : (
                            <div>
                              <div style={{ fontSize:24, marginBottom:4 }}>📷</div>
                              <div style={{ fontSize:12, color:"#6B7280" }}>拖拉或點擊上傳</div>
                            </div>
                          )
                        }
                      </div>
                      <input ref={fileRef} type="file" accept="image/*" style={{ display:"none" }} onChange={function(e) { handleImage(e.target.files[0]); }} />
                    </Field>
                    <Field label="產品描述" hint="請描述材質、規格、尺寸、顏色、用途、特色，愈詳細愈好">
                      <textarea value={description} onChange={e => setDescription(e.target.value)} rows={4} placeholder="例：304不鏽鋼，M8×30mm，六角頭，表面鏡面拋光…" style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", resize:"vertical" }} />
                    </Field>
                    <Field label="產品優勢與特色" hint="專利、認證、核心賣點。若不確定可留空，AI 會在下一步提供建議。">
                      <textarea value={prodAdvantages} onChange={e => setProdAdvantages(e.target.value)} rows={3} placeholder="例：擁有發明專利、通過ISO 9001、支援OEM…" style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", resize:"vertical" }} />
                    </Field>

                    <div style={{ borderTop:"1px dashed #E5E7EB", margin:"16px 0 18px", paddingTop:16 }}>
                      <div style={{ fontSize:12, fontWeight:700, color:"#ff6b35", marginBottom:14 }}>📄 商品詳情頁資料（選填）</div>
                      <Field label="規格尺寸／重量" hint="例：M8×30mm，重量5g">
                        <textarea value={prodSpecs} onChange={e => setProdSpecs(e.target.value)} rows={2} placeholder="尺寸、重量、材質等規格…" style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", resize:"vertical" }} />
                      </Field>
                      <Field label="包裝方式">
                        <input value={prodPackaging} onChange={e => setProdPackaging(e.target.value)} placeholder="例：每盒100pcs，外箱500pcs" style={inp} />
                      </Field>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                        <Field label="MOQ"><input value={prodMoq} onChange={e => setProdMoq(e.target.value)} placeholder="例：100 pcs" style={inp} /></Field>
                        <Field label="交期"><input value={prodLeadtime} onChange={e => setProdLeadtime(e.target.value)} placeholder="樣品7天，大貨30天" style={inp} /></Field>
                      </div>
                      <Field label="客製化服務">
                        <input value={prodCustomize} onChange={e => setProdCustomize(e.target.value)} placeholder="例：支援OEM、客製化Logo" style={inp} />
                      </Field>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                        <Field label="付款方式"><input value={prodPayment} onChange={e => setProdPayment(e.target.value)} placeholder="T/T, Trade Assurance" style={inp} /></Field>
                        <Field label="目標買家類型"><input value={prodBuyerType} onChange={e => setProdBuyerType(e.target.value)} placeholder="品牌商、電商賣家…" style={inp} /></Field>
                      </div>
                    </div>

                    <Field label="目標市場">
                      <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginTop:6 }}>
                        {LISTING_MARKETS.map(function(m) {
                          return (
                            <button key={m} onClick={() => setLMarket(m)} style={{ padding:"5px 12px", borderRadius:20, fontSize:12, cursor:"pointer", border:"1.5px solid " + (lMarket === m ? "#ff6b35" : "#E5E7EB"), background: lMarket === m ? "#FFF7ED" : "#fff", color: lMarket === m ? "#C2410C" : "#374151", fontWeight: lMarket === m ? 600 : 400 }}>{m}</button>
                          );
                        })}
                      </div>
                    </Field>

                    {/* 阿里上架欄位 */}
                    <div style={{ borderTop:"1px dashed #E5E7EB", margin:"16px 0 18px", paddingTop:16 }}>
                      <div style={{ fontSize:12, fontWeight:700, color:"#0052D9", marginBottom:14 }}>🔵 阿里後台批次上架資料（選填）</div>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                        <Field label="Place of Origin" hint="例：China / Taiwan"><input value={placeOfOrigin} onChange={e => setPlaceOfOrigin(e.target.value)} placeholder="China" style={inp} /></Field>
                        <Field label="Brand Name"><input value={brandName} onChange={e => setBrandName(e.target.value)} placeholder="Your Brand" style={inp} /></Field>
                      </div>
                      <Field label="Category" hint="若留空，阿里 AI 自動判斷">
                        <input value={category} onChange={e => setCategory(e.target.value)} placeholder="例：Electronic Components" style={inp} />
                      </Field>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                        <Field label="Unit">
                          <select value={unit} onChange={e => setUnit(e.target.value)} style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit" }}>
                            {["Pieces","Sets","Pairs","Kilograms","Meters","Boxes","Bags","Rolls","Liters","Tons"].map(function(u) { return <option key={u}>{u}</option>; })}
                          </select>
                        </Field>
                        <Field label="Currency">
                          <select value={currency} onChange={e => setCurrency(e.target.value)} style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit" }}>
                            {["USD","EUR","GBP","JPY","AUD","CAD","TWD"].map(function(c) { return <option key={c}>{c}</option>; })}
                          </select>
                        </Field>
                      </div>
                      <Field label="Pricing Type">
                        <div style={{ display:"flex", gap:10, marginTop:6 }}>
                          {["Tiered pricing","SKU pricing","Range pricing"].map(function(pt) {
                            return (
                              <label key={pt} style={{ display:"flex", alignItems:"center", gap:6, cursor:"pointer", fontSize:12, color:"#475569" }}>
                                <input type="radio" name="pricingType" value={pt} checked={pricingType === pt} onChange={() => setPricingType(pt)} style={{ accentColor:"#ff6b35" }} /> {pt}
                              </label>
                            );
                          })}
                        </div>
                      </Field>
                      {pricingType === "Tiered pricing" && (
                        <Field label="Tiered Pricing（階梯價）" hint="最多3階">
                          {tieredPrices.map(function(tp, i) {
                            return (
                              <div key={i} style={{ display:"flex", gap:8, marginBottom:8, alignItems:"center" }}>
                                <span style={{ fontSize:12, color:"#6B7280", minWidth:40 }}>第{i+1}階</span>
                                <input value={tp.price} onChange={function(e) { var n = tieredPrices.slice(); n[i] = Object.assign({}, n[i], { price:e.target.value }); setTieredPrices(n); }} placeholder="單價" style={{ flex:1, padding:"8px 10px", borderRadius:7, border:"1.5px solid #e2e8f0", fontSize:12, outline:"none" }} />
                                <span style={{ fontSize:12, color:"#6B7280" }}>{currency}</span>
                                <input value={tp.moq} onChange={function(e) { var n = tieredPrices.slice(); n[i] = Object.assign({}, n[i], { moq:e.target.value }); setTieredPrices(n); }} placeholder="MOQ" style={{ flex:1, padding:"8px 10px", borderRadius:7, border:"1.5px solid #e2e8f0", fontSize:12, outline:"none" }} />
                                <span style={{ fontSize:12, color:"#6B7280" }}>起</span>
                              </div>
                            );
                          })}
                        </Field>
                      )}
                      {pricingType === "SKU pricing" && (
                        <Field label="SKU Pricing"><input value={skuPrice} onChange={e => setSkuPrice(e.target.value)} placeholder="單價（數字）" style={inp} /></Field>
                      )}
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                        <Field label="Inventory" hint="留空預設 9,999,999"><input value={inventory} onChange={e => setInventory(e.target.value)} placeholder="例：5000" style={inp} /></Field>
                        <Field label="SPU ID" hint="商家自訂型號"><input value={spuId} onChange={e => setSpuId(e.target.value)} placeholder="MY-PROD-001" style={inp} /></Field>
                      </div>
                      <div style={{ marginBottom:14 }}>
                        <div style={{ fontSize:13, fontWeight:600, color:"#1e293b", marginBottom:8 }}>物流資訊 Logistics</div>
                        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8, marginBottom:8 }}>
                          <Field label="長 Length (cm)"><input value={itemLength} onChange={e => setItemLength(e.target.value)} placeholder="cm" style={inp} /></Field>
                          <Field label="寬 Width (cm)"><input value={itemWidth} onChange={e => setItemWidth(e.target.value)} placeholder="cm" style={inp} /></Field>
                          <Field label="高 Height (cm)"><input value={itemHeight} onChange={e => setItemHeight(e.target.value)} placeholder="cm" style={inp} /></Field>
                        </div>
                        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                          <Field label="Gross Weight (KG)"><input value={grossWeight} onChange={e => setGrossWeight(e.target.value)} placeholder="kg" style={inp} /></Field>
                          <Field label="HS Code"><input value={hsCode} onChange={e => setHsCode(e.target.value)} placeholder="例：6109100010" style={inp} /></Field>
                        </div>
                      </div>
                      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                        <Field label="Shipping Template" hint="需在阿里後台建立"><input value={shippingTemplate} onChange={e => setShippingTemplate(e.target.value)} placeholder="物流模板名稱" style={inp} /></Field>
                        <Field label="Lead Time（天）"><input value={shippingLeadtime} onChange={e => setShippingLeadtime(e.target.value)} placeholder="例：15" style={inp} /></Field>
                      </div>
                      <Field label="Logistics Attributes">
                        <select value={logisticsAttr} onChange={e => setLogisticsAttr(e.target.value)} style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit" }}>
                          {["Ordinary goods","Sensitive goods","Contraband"].map(function(a) { return <option key={a}>{a}</option>; })}
                        </select>
                      </Field>
                      <div style={{ background:"#FFF8F0", border:"1px solid #FED7AA", borderRadius:8, padding:"10px 12px", fontSize:11, color:"#92400E", lineHeight:1.7 }}>
                        ⚠️ 以下欄位需在阿里後台操作：Product Images（需上傳圖片取得 URL）、Shipping Template（需後台建立）
                      </div>
                    </div>

                    {lError && <div style={{ background:"#FFF1F2", color:"#E11D48", fontSize:12, padding:"8px 12px", borderRadius:8, marginBottom:12 }}>{lError}</div>}
                    <button onClick={analyzeProduct} disabled={lLoading1} style={{ width:"100%", padding:"12px", background: lLoading1 ? "#93C5FD" : "linear-gradient(135deg,#1B2A4A,#2D4270)", color:"#fff", border:"none", borderRadius:10, fontSize:15, fontWeight:700, cursor: lLoading1 ? "not-allowed" : "pointer" }}>
                      {lLoading1 ? "⏳ AI 分析中…" : "下一步：AI 分析產品優勢 →"}
                    </button>
                  </div>
                </div>

                <div style={{ flex:1 }}>
                  <div style={{ background:"#fff", borderRadius:14, padding:28, boxShadow:"0 1px 6px rgba(0,0,0,0.07)" }}>
                    <div style={{ fontSize:14, fontWeight:700, color:"#1B2A4A", marginBottom:16 }}>🗺 使用流程</div>
                    {[
                      { n:1, t:"填寫產品資訊", d:"輸入產品名稱、描述、優勢，越詳細AI越準確" },
                      { n:2, t:"確認產品優勢清單", d:"AI 根據產業分析，列出建議的認證、賣點供您勾選確認" },
                      { n:3, t:"生成 10 組 Listing", d:"AI 從5個維度產出10組差異化建議，避免重舖" }
                    ].map(function(s) {
                      return (
                        <div key={s.n} style={{ display:"flex", gap:14, marginBottom:18 }}>
                          <div style={{ width:30, height:30, borderRadius:"50%", flexShrink:0, background:"#FFF7ED", color:"#ff6b35", display:"flex", alignItems:"center", justifyContent:"center", fontSize:13, fontWeight:800 }}>{s.n}</div>
                          <div>
                            <div style={{ fontWeight:600, fontSize:14, color:"#111827", marginBottom:3 }}>{s.t}</div>
                            <div style={{ fontSize:12, color:"#6B7280", lineHeight:1.6 }}>{s.d}</div>
                          </div>
                        </div>
                      );
                    })}
                    {companyUnlocked && (
                      <div style={{ background:"#EEF2FF", borderRadius:9, padding:"11px 13px" }}>
                        <div style={{ fontSize:12, color:"#3B5BDB", lineHeight:1.7 }}>💡 公司資料已自動帶入，AI 生成時會參考公司的認證、目標市場、核心優勢。</div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}

            {lStep === 2 && suggestions && (
              <div style={{ maxWidth:800, margin:"0 auto" }}>
                <div style={{ background:"#fff", borderRadius:14, padding:28, boxShadow:"0 1px 6px rgba(0,0,0,0.07)" }}>
                  <div style={{ fontWeight:700, color:"#1B2A4A", fontSize:16, marginBottom:8 }}>AI 產業分析結果</div>
                  <div style={{ background:"#F0F4FF", borderRadius:8, padding:"10px 14px", fontSize:13, color:"#374151", marginBottom:16 }}>{suggestions.industry_summary}</div>
                  <div style={{ fontSize:13, color:"#6B7280", marginBottom:20 }}>以下是 AI 建議的項目。<strong style={{ color:"#111827" }}>請勾選您實際具備的項目</strong>，未勾選的不會出現在 Listing 中。</div>
                  {[
                    { key:"advantages", label:"🏆 產品優勢", color:"#3B5BDB", bg:"#EEF2FF" },
                    { key:"certifications", label:"📋 認證與規範", color:"#16A34A", bg:"#F0FDF4" },
                    { key:"selling_points", label:"💎 核心賣點", color:"#C2410C", bg:"#FFF7ED" }
                  ].map(function(cat) {
                    return (
                      <div key={cat.key} style={{ marginBottom:22 }}>
                        <div style={{ fontSize:13, fontWeight:700, color:cat.color, marginBottom:10, paddingBottom:6, borderBottom:"2px solid " + cat.bg }}>{cat.label}</div>
                        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                          {(suggestions[cat.key] || []).map(function(item) {
                            return (
                              <div key={item.id} onClick={() => setChecked(function(p) { return Object.assign({}, p, { [item.id]: !p[item.id] }); })}
                                style={{ display:"flex", alignItems:"flex-start", gap:12, padding:"10px 14px", borderRadius:8, cursor:"pointer", border:"1.5px solid " + (checked[item.id] ? cat.color : "#E5E7EB"), background: checked[item.id] ? cat.bg : "#FAFAFA" }}>
                                <div style={{ width:20, height:20, borderRadius:5, flexShrink:0, marginTop:1, border:"2px solid " + (checked[item.id] ? cat.color : "#D1D5DB"), background: checked[item.id] ? cat.color : "#fff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, color:"#fff", fontWeight:700 }}>{checked[item.id] ? "✓" : ""}</div>
                                <div>
                                  <div style={{ fontSize:13, fontWeight:600, color:"#111827" }}>{item.label}</div>
                                  <div style={{ fontSize:11, color:"#6B7280", marginTop:2 }}>{item.hint}</div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                  {lError && <div style={{ background:"#FFF1F2", color:"#E11D48", fontSize:12, padding:"8px 12px", borderRadius:8, marginBottom:12 }}>{lError}</div>}
                  <div style={{ display:"flex", gap:10 }}>
                    <button onClick={() => setLStep(1)} style={{ flex:1, padding:"11px", background:"#F3F4F6", color:"#374151", border:"none", borderRadius:10, fontSize:14, fontWeight:600, cursor:"pointer" }}>← 返回修改</button>
                    <button onClick={generateListings} disabled={lLoading2} style={{ flex:2, padding:"11px", background: lLoading2 ? "#93C5FD" : "linear-gradient(135deg,#ff6b35,#f7931e)", color:"#fff", border:"none", borderRadius:10, fontSize:14, fontWeight:700, cursor: lLoading2 ? "not-allowed" : "pointer" }}>
                      {lLoading2 ? "⏳ 生成中…" : "⚡ 確認並生成 10 組 Listing"}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {lStep === 3 && (
              <div>
                <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:16 }}>
                  <div style={{ fontWeight:700, color:"#1B2A4A", fontSize:15 }}>
                    共 {batches.reduce(function(s,b){ return s+b.listings.length; },0)} 組 Listing
                    <span style={{ fontWeight:400, color:"#16A34A", fontSize:13, marginLeft:8 }}>已勾選 {totalSelectedCount()} 組待匯出</span>
                  </div>
                  <div style={{ display:"flex", gap:8 }}>
                    <button onClick={resetProduct} style={{ background:"#F3F4F6", color:"#374151", border:"none", borderRadius:8, padding:"8px 14px", fontSize:13, fontWeight:600, cursor:"pointer" }}>← 全部重新開始</button>
                    <button
                      onClick={function() {
                        var data = getSelectedExportData();
                        if (!data.length) { setExportMsg("⚠ 請先勾選"); setTimeout(function(){ setExportMsg(""); }, 2000); return; }
                        exportToExcel(data, "listing_export.xlsx", getExtraData());
                        setExportMsg("✓ 已下載");
                        setTimeout(function() { setExportMsg(""); }, 2500);
                      }}
                      style={{ background: exportMsg.startsWith("✓") ? "#16A34A" : exportMsg.startsWith("⚠") ? "#D97706" : "#1B2A4A", color:"#fff", border:"none", borderRadius:8, padding:"8px 16px", fontSize:13, fontWeight:600, cursor:"pointer", minWidth:140 }}
                    >
                      {exportMsg || ("⬇ 匯出已勾選 (" + totalSelectedCount() + ")")}
                    </button>
                  </div>
                </div>

                {batches.map(function(batch, bi) {
                  var batchSelectedCount = Object.keys(batch.checkedMap).filter(function(k){ return batch.checkedMap[k]; }).length;
                  return (
                    <div key={batch.id} style={{ marginBottom:20 }}>
                      <div onClick={function(){ toggleBatchCollapse(batch.id); }} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", background:"#1B2A4A", borderRadius: batch.collapsed ? 10 : "10px 10px 0 0", padding:"12px 18px", cursor:"pointer" }}>
                        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                          <span style={{ color:"#fff", fontWeight:700, fontSize:14 }}>{batch.productName}</span>
                          <span style={{ background:"rgba(255,255,255,0.15)", color:"#FDBA74", fontSize:11, fontWeight:700, padding:"2px 9px", borderRadius:20 }}>第 {batch.round} 輪</span>
                          <span style={{ color:"#8BA0C4", fontSize:12 }}>{batch.listings.length} 組｜已勾選 {batchSelectedCount}</span>
                        </div>
                        <span style={{ color:"#8BA0C4", fontSize:12 }}>{batch.collapsed ? "▼ 展開" : "▲ 收起"}</span>
                      </div>
                      {!batch.collapsed && (
                        <div style={{ background:"#F0F2F8", borderRadius:"0 0 10px 10px", padding:"16px" }}>
                          {batch.detailPage && (
                            <div style={{ background:"#fff", border:"1px solid #E5E7EB", borderRadius:12, marginBottom:16, overflow:"hidden", boxShadow:"0 1px 4px rgba(0,0,0,0.05)" }}>
                              <div style={{ background:"linear-gradient(135deg,#1B2A4A,#2D4270)", padding:"14px 20px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                                  <span style={{ fontSize:18 }}>📄</span>
                                  <div>
                                    <div style={{ color:"#fff", fontWeight:700, fontSize:14 }}>商品詳情頁文案</div>
                                    <div style={{ color:"#8BA0C4", fontSize:11, marginTop:2 }}>七區塊完整結構，可直接複製到阿里國際站後台</div>
                                  </div>
                                </div>
                                <div style={{ display:"flex", gap:8 }}>
                                  <button onClick={function(){ toggleBatchDetail(batch.id); }} style={{ background:"rgba(255,255,255,0.15)", color:"#fff", border:"none", borderRadius:7, padding:"6px 14px", fontSize:12, cursor:"pointer", fontWeight:600 }}>
                                    {batch.showDetail ? "▲ 收起" : "▼ 展開"}
                                  </button>
                                  <CopyBtn text={batch.detailPage} label="複製全文" />
                                </div>
                              </div>
                              {batch.showDetail && (
                                <div style={{ padding:"20px 22px" }}>
                                  <pre style={{ fontFamily:"Inter, PingFang TC, sans-serif", fontSize:13, color:"#374151", lineHeight:1.8, whiteSpace:"pre-wrap", wordBreak:"break-word", margin:0 }}>{batch.detailPage}</pre>
                                </div>
                              )}
                            </div>
                          )}
                          {batch.listings.map(function(item, i) {
                            return <ListingCard key={i} item={item} index={i} isChecked={!!batch.checkedMap[i]} onToggle={function(){ toggleListingChecked(batch.id, i); }} />;
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}

                {/* 底部操作區 */}
                <div style={{ background:"#fff", borderRadius:14, padding:24, boxShadow:"0 1px 6px rgba(0,0,0,0.07)", marginTop:8 }}>
                  {!expandMore && (
                    <div style={{ display:"flex", gap:12 }}>
                      <button onClick={function(){ setExpandMore(true); }} style={{ flex:1, padding:"13px", background:"linear-gradient(135deg,#ff6b35,#f7931e)", color:"#fff", border:"none", borderRadius:10, fontSize:14, fontWeight:700, cursor:"pointer" }}>
                        🔁 再擴10組（同一產品：{productName}）
                      </button>
                      <button onClick={resetProductOnly} style={{ flex:1, padding:"13px", background:"linear-gradient(135deg,#059669,#10b981)", color:"#fff", border:"none", borderRadius:10, fontSize:14, fontWeight:700, cursor:"pointer" }}>
                        ＋ 新增其它品項的10組產品
                      </button>
                    </div>
                  )}
                  {expandMore && (
                    <div>
                      <div style={{ fontWeight:700, color:"#1B2A4A", fontSize:14, marginBottom:10 }}>🔁 再擴10組 — {productName}</div>
                      <div style={{ fontSize:12, color:"#6B7280", marginBottom:10 }}>可補充想強調的重點或調整方向，留空則直接以原資訊再生成新的差異化角度。</div>
                      <textarea value={moreNote} onChange={function(e){ setMoreNote(e.target.value); }} rows={3} placeholder="例：這次想多強調歐洲市場的認證、或加強耐用度訴求…" style={{ width:"100%", padding:"9px 12px", borderRadius:8, border:"1.5px solid #e2e8f0", fontSize:13, outline:"none", boxSizing:"border-box", fontFamily:"inherit", resize:"vertical", marginBottom:14 }} />
                      {lError && <div style={{ background:"#FFF1F2", color:"#E11D48", fontSize:12, padding:"8px 12px", borderRadius:8, marginBottom:12 }}>{lError}</div>}
                      <div style={{ display:"flex", gap:10 }}>
                        <button onClick={function(){ setExpandMore(false); setMoreNote(""); }} style={{ padding:"11px 18px", background:"#F3F4F6", color:"#374151", border:"none", borderRadius:10, fontSize:14, fontWeight:600, cursor:"pointer" }}>取消</button>
                        <button onClick={generateMoreForSameProduct} disabled={moreLoading} style={{ flex:1, padding:"11px", background: moreLoading ? "#93C5FD" : "#F3F4F6", color: moreLoading ? "#fff" : "#374151", border:"none", borderRadius:10, fontSize:14, fontWeight:600, cursor: moreLoading ? "not-allowed" : "pointer" }}>
                          {moreLoading ? "⏳ 生成中…" : "略過，直接再產出"}
                        </button>
                        <button onClick={generateMoreForSameProduct} disabled={moreLoading} style={{ flex:2, padding:"11px", background: moreLoading ? "#93C5FD" : "linear-gradient(135deg,#ff6b35,#f7931e)", color:"#fff", border:"none", borderRadius:10, fontSize:14, fontWeight:700, cursor: moreLoading ? "not-allowed" : "pointer" }}>
                          {moreLoading ? "⏳ 生成中…" : "⚡ 帶補充內容再生成10組"}
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  );
}
