/**
 * MHEdu - Công cụ Tích Hợp Tổng Hợp Vào Giáo Án
 * Hỗ trợ 2 chế độ:
 *   1. Tích hợp vào giáo án có sẵn (.docx)
 *   2. Tạo một giáo án mới có tích hợp (từ SGK PDF / ảnh chụp)
 * 8 nội dung tích hợp: NLS (tím), NL AI QĐ 2422 (đỏ), HSKT (xanh lá), GDQPAN (xanh dương),
 * BĐKH-PCTT (cam), GD LTCM ĐĐLS (vàng), An toàn mạng CV 6206 (xanh ngọc), Tư tưởng HCM QĐ 2995 (hồng cánh sen)
 */

import { readDocxFile, readXlsxFile, guessMetadata, generatePatchedDocx, generateNewLessonDocx, CATEGORIES } from './assets/docxModel.js';
import { ATM_CATALOG, AI_CATALOG_BY_GRADE, NLS_TT02_COMPONENTS } from './assets/frameworksData.js';

const { useState, useEffect, useMemo } = React;

const SUBJECTS = [
  "Tin học", "Toán", "Ngữ văn", "Tiếng Anh", "Khoa học tự nhiên", "Lịch sử và Địa lí",
  "Giáo dục công dân", "Công nghệ", "Giáo dục thể chất", "Nghệ thuật", "Âm nhạc", "Mĩ thuật",
  "Hoạt động trải nghiệm", "Vật lí", "Hoá học", "Sinh học", "Lịch sử", "Địa lí",
  "Giáo dục Kinh tế và Pháp luật", "Giáo dục quốc phòng và an ninh", "Tiếng Việt", "Tự nhiên và Xã hội", "Khoa học"
];

const DISABILITY_TYPES = [
  { key: "chung", label: "Khuyết tật chung" },
  { key: "nhin", label: "Khuyết tật về mắt (nhìn)" },
  { key: "vandong", label: "Khuyết tật về xương (vận động)" },
  { key: "tritue", label: "Khuyết tật về trí tuệ" },
  { key: "nghenoi", label: "Khuyết tật nghe, nói" },
  { key: "tuky", label: "Rối loạn phổ tự kỷ" },
  { key: "rlht", label: "Rối loạn học tập đặc thù" },
  { key: "thankinh", label: "Khuyết tật thần kinh" },
  { key: "tamthan", label: "Khuyết tật tâm thần" },
  { key: "khac", label: "Khuyết tật khác" }
];

function getNlsLevelLabel(grade) {
  if (grade <= 3) return "mức CB1 — lớp 1–3";
  if (grade <= 5) return "mức CB2 — lớp 4–5";
  if (grade <= 7) return "mức TC1 — lớp 6–7";
  if (grade <= 9) return "mức TC2 — lớp 8–9";
  return "mức NC1 — lớp 10–12";
}

function getNlsLevelTag(grade) {
  if (grade <= 3) return "CB1";
  if (grade <= 5) return "CB2";
  if (grade <= 7) return "TC1";
  if (grade <= 9) return "TC2";
  return "NC1";
}

function getNlsCatalog(grade) {
  const lvl = getNlsLevelTag(grade);
  return NLS_TT02_COMPONENTS.map(c => ({
    code: `${c.id}.${lvl}a`,
    desc: `${c.domain}: ${c.desc}`
  }));
}

function getAiCatalog(grade) {
  return AI_CATALOG_BY_GRADE[Number(grade)] || AI_CATALOG_BY_GRADE[10];
}

const GDQPAN_SUGGESTIONS = [
  "Liên hệ bài toán số liệu bảo vệ chủ quyền biên giới và biển đảo Tổ quốc",
  "Giáo dục ý thức an toàn thông tin cá nhân và an ninh mạng",
  "Lồng ghép truyền thống vẻ vang, gương anh hùng liệt sĩ QĐND và CAND",
  "Nhắc nhở tinh thần kỷ luật, đoàn kết và trách nhiệm của học sinh với Tổ quốc"
];

const BDKH_SUGGESTIONS = [
  "Liên hệ bài toán số liệu về lượng rác thải, tiết kiệm năng lượng điện nước",
  "Tổ chức hoạt động phân loại rác và giảm thiểu đồ nhựa dùng một lần",
  "Thực hành lối sống xanh, chăm sóc và bảo vệ cây xanh học đường",
  "Ý thức thích ứng và phòng tránh tai nạn thương tích khi thiên tai bão lũ"
];

const INTEGRATION_CARDS = [
  {
    key: "nls",
    title: "Năng lực số",
    desc: "Mục tiêu + hoạt động, chữ tím, mã chỉ báo đầu dòng ở mục tiêu chung, cuối câu trong hoạt động (vd ... (1.2.TC1a))",
    dot: "#7C3AED"
  },
  {
    key: "ai",
    title: "Năng lực trí tuệ nhân tạo",
    desc: "Theo khung QĐ 2422, chữ đỏ",
    dot: "#DC2626"
  },
  {
    key: "kt",
    title: "Học sinh khuyết tật",
    desc: "Mục tiêu riêng + nhiệm vụ điều chỉnh ở từng hoạt động, chữ xanh lá",
    dot: "#15803D"
  },
  {
    key: "qp",
    title: "Giáo dục quốc phòng và an ninh",
    desc: "Chữ xanh dương, có ghi [Tích hợp GDQPAN]",
    dot: "#1D4ED8"
  },
  {
    key: "bd",
    title: "Biến đổi khí hậu - phòng chống thiên tai",
    desc: "Chữ cam, có ghi [Tích hợp BĐKH-PCTT]",
    dot: "#C2410C"
  },
  {
    key: "ltcm",
    title: "Giáo dục lý tưởng cách mạng, đạo đức, lối sống",
    desc: "Lớp 10–12 (QĐ 2546) và Ngữ văn lớp 6–9, chữ vàng, có ghi [Tích hợp GD LTCM, ĐĐLS]",
    dot: "#CA8A04"
  },
  {
    key: "atm",
    title: "Giáo dục an toàn mạng toàn diện",
    desc: "Theo CV 6206, mọi lớp, mọi môn, chữ xanh ngọc, mã ATM1–ATM8 đầu dòng ở mục tiêu chung, cuối câu trong hoạt động (vd ... (ATM4))",
    dot: "#0D9488"
  },
  {
    key: "hcm",
    title: "Giáo dục tư tưởng, đạo đức, phong cách Hồ Chí Minh",
    desc: "Theo QĐ 2995, mọi lớp, mọi môn, chữ hồng cánh sen, có ghi [Tích hợp GD tư tưởng, đạo đức, phong cách HCM]",
    dot: "#DB2777"
  }
];

// IndexedDB helper để lưu SGK đã tải trên máy
const DB_NAME = "mhedu_sgk_cache_v1";
const STORE_NAME = "textbooks";

function openSgkDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME, { keyPath: "id" });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function listSavedTextbooks() {
  try {
    const db = await openSgkDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readonly");
      const req = tx.objectStore(STORE_NAME).getAll();
      req.onsuccess = () => resolve((req.result || []).sort((a, b) => b.updatedAt - a.updatedAt));
      req.onerror = () => reject(req.error);
    });
  } catch { return []; }
}

async function saveTextbookToDb(record) {
  try {
    const db = await openSgkDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readwrite");
      tx.objectStore(STORE_NAME).put(record);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  } catch {}
}

async function deleteTextbookFromDb(id) {
  try {
    const db = await openSgkDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, "readwrite");
      tx.objectStore(STORE_NAME).delete(id);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  } catch {}
}

function guessSubjectAndGradeFromName(fileName) {
  const lower = fileName.toLowerCase();
  let g = null;
  const gm = /(?:^|[^0-9])(1[0-2]|[1-9])(?:[^0-9]|$)/.exec(lower);
  if (gm) g = Number(gm[1]);
  const subjectHints = [
    ["tin", "Tin học"], ["toan", "Toán"], ["toán", "Toán"], ["van", "Ngữ văn"], ["văn", "Ngữ văn"],
    ["anh", "Tiếng Anh"], ["khtn", "Khoa học tự nhiên"], ["gdcd", "Giáo dục công dân"],
    ["cong nghe", "Công nghệ"], ["vat li", "Vật lí"], ["hoa hoc", "Hoá học"], ["sinh hoc", "Sinh học"],
    ["lich su", "Lịch sử"], ["dia li", "Địa lí"], ["tieng viet", "Tiếng Việt"], ["khoa hoc", "Khoa học"]
  ];
  let s = null;
  for (const [kw, label] of subjectHints) {
    if (lower.includes(kw)) { s = label; break; }
  }
  return { grade: g, subject: s };
}

function extractCodesFromText(text, pattern) {
  const matches = String(text || "").match(pattern) || [];
  return [...new Set(matches)];
}

export function App() {
  const [mode, setMode] = useState("existing");
  const [step, setStep] = useState(0);

  // Thông tin chung
  const [subject, setSubject] = useState("Toán");
  const [grade, setGrade] = useState(10);
  const [title, setTitle] = useState("");
  const [ppctName, setPpctName] = useState("");
  const [ppctText, setPpctText] = useState("");

  // Mode 1
  const [docInfo, setDocInfo] = useState(null);
  const [isDragging, setIsDragging] = useState(false);

  // Mode 2
  const [savedTextbooks, setSavedTextbooks] = useState([]);
  const [sgkFile, setSgkFile] = useState(null);
  const [pageFrom, setPageFrom] = useState(5);
  const [pageTo, setPageTo] = useState(8);
  const [pagePreviews, setPagePreviews] = useState([]);
  const [periods, setPeriods] = useState(1);
  const [activityList, setActivityList] = useState([]); // ["Hoạt động 1", ...]
  const [renderingPages, setRenderingPages] = useState(false);

  // 8 nội dung tích hợp
  const [opts, setOpts] = useState({
    nls: true, ai: false, kt: false, qp: false, bd: false, ltcm: false, atm: false, hcm: false
  });

  // NLS & AI & ATM selected codes (with editable YCCĐ)
  const [selectedNlsItems, setSelectedNlsItems] = useState([]); // [{ code, desc }]
  const [selectedAiItems, setSelectedAiItems] = useState([]);   // [{ code, desc }]
  const [selectedAtmItems, setSelectedAtmItems] = useState([]); // [{ code, desc }]

  // HSKT rows (up to 5)
  const [disabilityRows, setDisabilityRows] = useState([{ type: "Khuyết tật chung", className: "" }]);
  const [disabilityNote, setDisabilityNote] = useState("");

  // Notes for others
  const [qpNote, setQpNote] = useState("");
  const [bdNote, setBdNote] = useState("");
  const [ltcmNote, setLtcmNote] = useState("");
  const [hcmNote, setHcmNote] = useState("");

  // Extra options
  const [lessonLang, setLessonLang] = useState("Tiếng Việt");
  const [extraInstructions, setExtraInstructions] = useState("");
  const [strongerAi, setStrongerAi] = useState(false);

  // Processing & Results
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressStatus, setProgressStatus] = useState("");
  const [error, setError] = useState(null);

  const [items, setItems] = useState([]);
  const [newLessonData, setNewLessonData] = useState(null);
  const [colorEnabled, setColorEnabled] = useState(true);
  const [activeFilter, setActiveFilter] = useState("all");
  const [copiedReport, setCopiedReport] = useState(false);
  const [downloading, setDownloading] = useState(false);

  useEffect(() => {
    listSavedTextbooks().then(setSavedTextbooks);
  }, []);

  const nlsCatalog = useMemo(() => getNlsCatalog(grade), [grade]);
  const aiCatalog = useMemo(() => getAiCatalog(grade), [grade]);

  const switchMode = (nextMode) => {
    setMode(nextMode);
    setStep(0);
    setError(null);
  };

  // ==================== HANDLERS ====================
  const handleFileUpload = async (file) => {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setError("Vui lòng chọn file Word định dạng .docx (Word 2007 trở lên). Với file .doc cũ, hãy mở bằng Word rồi Lưu thành .docx.");
      return;
    }
    setError(null);
    setLoading(true);
    setProgress(30);
    setProgressStatus("Đang đọc và phân tích cấu trúc giáo án...");
    try {
      const parsed = await readDocxFile(file);
      setDocInfo(parsed);
      const meta = guessMetadata(parsed);
      if (meta.grade) setGrade(meta.grade);
      if (meta.title) setTitle(meta.title);
      const guessed = guessSubjectAndGradeFromName(file.name);
      if (guessed.subject) setSubject(guessed.subject);
      setProgress(100);
      setItems([]);
    } catch (err) {
      setError(err.message || "Không thể đọc file Word. Vui lòng kiểm tra lại file.");
    } finally {
      setLoading(false);
    }
  };

  const handlePpctUpload = async (file) => {
    if (!file) return;
    try {
      const lower = file.name.toLowerCase();
      if (lower.endsWith(".docx")) {
        const parsed = await readDocxFile(file);
        setPpctName(file.name);
        setPpctText(parsed.text.slice(0, 6000));
      } else if (lower.endsWith(".xlsx")) {
        const text = await readXlsxFile(file);
        setPpctName(file.name);
        setPpctText(text.slice(0, 6000));
      } else {
        setError("PPCT chỉ hỗ trợ file .docx hoặc .xlsx.");
      }
    } catch (err) {
      setError("Không thể đọc file PPCT: " + (err.message || "Lỗi định dạng"));
    }
  };

  const loadPdfBuffer = async (buf, fileName, idForSave = null) => {
    if (!window.pdfjsLib) throw new Error("Thư viện đọc PDF chưa sẵn sàng.");
    const pdfDoc = await window.pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
    const numPages = pdfDoc.numPages;
    const guessed = guessSubjectAndGradeFromName(fileName);
    if (guessed.grade) setGrade(guessed.grade);
    if (guessed.subject) setSubject(guessed.subject);

    const recordId = idForSave || `sgk-${fileName}-${buf.byteLength}`;
    setSgkFile({ id: recordId, name: fileName, type: "pdf", numPages, pdfDoc });
    const defFrom = Math.min(5, numPages);
    const defTo = Math.min(defFrom + 3, numPages);
    setPageFrom(defFrom);
    setPageTo(defTo);
    setPagePreviews([]);

    if (!idForSave) {
      await saveTextbookToDb({
        id: recordId, name: fileName, numPages,
        grade: guessed.grade || grade, subject: guessed.subject || subject,
        buf, updatedAt: Date.now()
      });
      listSavedTextbooks().then(setSavedTextbooks);
    }
    await renderPdfRange(pdfDoc, defFrom, defTo);
  };

  const renderPdfRange = async (pdfDoc, fromP, toP) => {
    if (!pdfDoc) return;
    const start = Math.max(1, Number(fromP) || 1);
    const end = Math.min(pdfDoc.numPages, Math.max(start, Number(toP) || start));
    if (end - start + 1 > 20) {
      setError("Chỉ chọn tối đa 20 trang SGK cho mỗi bài dạy.");
      return;
    }
    setError(null);
    setRenderingPages(true);
    try {
      const previews = [];
      for (let p = start; p <= end; p++) {
        const page = await pdfDoc.getPage(p);
        const viewport = page.getViewport({ scale: 1.2 });
        const canvas = document.createElement("canvas");
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
        const dataUrl = canvas.toDataURL("image/jpeg", 0.78);
        let text = "";
        try {
          const tc = await page.getTextContent();
          text = tc.items.map(it => it.str).join(" ").replace(/\s+/g, " ").trim();
        } catch {}
        previews.push({ pageNum: p, dataUrl, text });
      }
      setPagePreviews(previews);
    } catch (e) {
      setError("Không thể hiển thị trang SGK: " + (e.message || ""));
    } finally {
      setRenderingPages(false);
    }
  };

  const handleSgkUpload = async (files) => {
    if (!files || files.length === 0) return;
    setError(null);
    const first = files[0];
    if (first.name.toLowerCase().endsWith(".pdf")) {
      setRenderingPages(true);
      try {
        await loadPdfBuffer(await first.arrayBuffer(), first.name);
      } catch (e) {
        setError("Lỗi đọc file PDF SGK: " + (e.message || ""));
      } finally {
        setRenderingPages(false);
      }
    } else {
      const imgs = Array.from(files).slice(0, 20);
      const previews = [];
      for (let i = 0; i < imgs.length; i++) {
        const dataUrl = await new Promise((res) => {
          const reader = new FileReader();
          reader.onload = () => res(reader.result);
          reader.readAsDataURL(imgs[i]);
        });
        previews.push({ pageNum: i + 1, dataUrl, text: "" });
      }
      setSgkFile({ id: `imgs-${Date.now()}`, name: `${imgs.length} ảnh chụp SGK`, type: "images", numPages: imgs.length });
      setPageFrom(1);
      setPageTo(imgs.length);
      setPagePreviews(previews);
    }
  };

  const handleUseSavedSgk = async (tb) => {
    setError(null);
    setRenderingPages(true);
    try {
      if (tb.grade) setGrade(tb.grade);
      if (tb.subject) setSubject(tb.subject);
      await loadPdfBuffer(tb.buf, tb.name, tb.id);
    } catch (e) {
      setError("Không thể mở SGK đã lưu: " + (e.message || ""));
    } finally {
      setRenderingPages(false);
    }
  };

  // ==================== GEMINI AI CALLER ====================
  const callGeminiJson = async (parts) => {
    let userKey = "";
    try { userKey = (localStorage.getItem("smartedu_gemini_api_key") || "").trim(); } catch {}
    const BACKUP_KEYS = [];
    const keysToTry = (userKey ? [userKey, ...BACKUP_KEYS] : BACKUP_KEYS).filter(Boolean);
    if (keysToTry.length === 0) {
      if (typeof window.openKeyModal === "function") window.openKeyModal();
      throw new Error("Chưa có Gemini API Key. Vui lòng lấy key miễn phí tại https://aistudio.google.com/apikey và dán vào hộp thoại '🔑 Đổi API Key'.");
    }

    const MODELS = strongerAi
      ? ["gemini-3.8-flash", "gemini-3.5-flash", "gemini-3-flash-preview", "gemini-3.5-flash-lite"]
      : ["gemini-3.5-flash", "gemini-3-flash-preview", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-3.8-flash"];

    let lastErr = null;
    for (const key of keysToTry) {
      for (const model of MODELS) {
        try {
          setProgressStatus(`AI (${model}) đang phân tích và xây dựng nội dung tích hợp...`);
          const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              contents: [{ role: "user", parts }],
              generationConfig: { temperature: 0.25, responseMimeType: "application/json" }
            })
          });
          if (!res.ok) {
            const errData = await res.json().catch(() => ({}));
            throw new Error(errData.error?.message || `HTTP ${res.status}`);
          }
          const data = await res.json();
          const rawText = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
          const cleaned = rawText.replace(/```json/gi, "").replace(/```/g, "").trim();
          return JSON.parse(cleaned);
        } catch (e) {
          lastErr = e;
          console.warn(`Thử model ${model} thất bại:`, e.message);
        }
      }
    }
    throw lastErr || new Error("AI không phản hồi.");
  };

  const buildIntegrationRequirementsText = () => {
    const hsktList = disabilityRows
      .filter(r => r.type)
      .map(r => r.className ? `${r.type} (Lớp ${r.className})` : r.type);

    return `
NGÔN NGỮ HIỂN THỊ CỦA GIÁO ÁN: ${lessonLang}.
${ppctText ? `CĂN CỨ PHÂN PHỐI CHƯƠNG TRÌNH (PPCT) ĐÃ TẢI (dùng chung NLS, NL AI, An toàn mạng):\n"""\n${ppctText.slice(0, 3500)}\n"""\n` : ""}
CÁC NỘI DUNG YÊU CẦU TÍCH HỢP:
${opts.nls ? `- NĂNG LỰC SỐ (NLS, mã "nls", theo TT 02/2025/TT-BGDĐT gồm 6 miền: I. Khai thác dữ liệu & thông tin [1.1–1.3]; II. Giao tiếp & hợp tác [2.1–2.6]; III. Sáng tạo nội dung số [3.1–3.4]; IV. An toàn [4.1–4.4]; V. Giải quyết vấn đề [5.1–5.4]; VI. Ứng dụng AI [6.1–6.3]): Lồng ghép mục tiêu + hoạt động; mã chỉ báo đặt đầu dòng ở mục tiêu chung (vd: "1.2.${getNlsLevelTag(grade)}a – ..."), hoặc cuối câu trong hoạt động (vd: "... (1.2.${getNlsLevelTag(grade)}a)"). ${selectedNlsItems.length ? `Bắt buộc ưu tiên các mã đã chọn: ${selectedNlsItems.map(x => `${x.code} (${x.desc})`).join("; ")}.` : `Tự chọn mã phù hợp trong khung lớp ${grade} (${getNlsLevelLabel(grade)}).`}` : ""}
${opts.ai ? `- NĂNG LỰC TRÍ TUỆ NHÂN TẠO (NL AI, mã "ai"): Theo Khung nội dung giáo dục AI QĐ 2422/QĐ-BGDĐT (18/8/2026) gồm 4 mạch A (Tư duy lấy con người làm trung tâm), B (Đạo đức AI), C (Kĩ thuật & ứng dụng AI), D (Thiết kế hệ thống AI); mã hóa chuẩn [Lớp].[Mã chủ đề].[STT] (mở rộng có MR). Lồng ghép nhẹ nhàng, không gây quá tải môn học, chú trọng kiểm chứng thông tin và khai báo minh bạch khi dùng AI. ${selectedAiItems.length ? `Bắt buộc ưu tiên các mã: ${selectedAiItems.map(x => `${x.code} (${x.desc})`).join("; ")}.` : `Chọn đúng mã YCCĐ phù hợp trong khung lớp ${grade} dưới đây:\n${aiCatalog.map(x => `  + ${x.code}: ${x.desc}`).join("\n")}`}` : ""}
${opts.kt ? `- HỌC SINH KHUYẾT TẬT HÒA NHẬP (mã "kt"): Lồng ghép mục tiêu riêng + nhiệm vụ điều chỉnh giảm tải ở từng hoạt động cho: ${hsktList.join("; ") || "Khuyết tật chung"}. Chỉ ghi tên lớp (nếu có), KHÔNG ghi tên học sinh. ${disabilityNote ? `[Ghi chú: ${disabilityNote}]` : ""}` : ""}
${opts.qp ? `- GIÁO DỤC QUỐC PHÒNG VÀ AN NINH (GDQPAN, mã "qp"): Lồng ghép theo TT 08/2024/TT-BGDĐT. ${qpNote ? `[Ghi chú: ${qpNote}]` : ""}` : ""}
${opts.bd ? `- BIẾN ĐỔI KHÍ HẬU - PHÒNG CHỐNG THIÊN TAI (BĐKH-PCTT, mã "bd"): ${bdNote ? `[Ghi chú: ${bdNote}]` : "Lồng ghép giáo dục môi trường, lối sống xanh, tiết kiệm năng lượng hoặc phòng tránh thiên tai."}` : ""}
${opts.ltcm ? `- GIÁO DỤC LÝ TƯỞNG CÁCH MẠNG, ĐẠO ĐỨC, LỐI SỐNG (mã "ltcm"): Theo QĐ 2546 (lớp 10–12) hoặc Ngữ văn 6–9. ${ltcmNote ? `[Ghi chú: ${ltcmNote}]` : ""}` : ""}
${opts.atm ? `- GIÁO DỤC AN TOÀN MẠNG TOÀN DIỆN (mã "atm"): Theo Công văn 6206/BGDĐT-GDPT (14/9/2026) và Sơ đồ tra cứu tích hợp An toàn mạng – Năng lực số TT 02/2025:
${ATM_CATALOG.map(a => `  + ${a.code} (${a.title}): NL bắt buộc [${a.requiredNls.join(", ")}]; NL khuyến khích [${a.recommendedNls.join(", ")}] — ${a.desc}`).join("\n")}
  Đặt mã ATM1–ATM8 đầu dòng ở mục tiêu chung, cuối câu trong hoạt động (vd "... (ATM4)"). Khi tích hợp cùng NLS, đối chiếu đúng các mã NL TT02 liên kết ở trên. ${selectedAtmItems.length ? `Ưu tiên các mã: ${selectedAtmItems.map(x => `${x.code} (${x.title}: ${x.desc})`).join("; ")}.` : "Tự chọn 1–2 mã ATM1–ATM8 phù hợp bài học."}` : ""}
${opts.hcm ? `- GIÁO DỤC TƯ TƯỞNG, ĐẠO ĐỨC, PHONG CÁCH HỒ CHÍ MINH (mã "hcm"): Theo QĐ 2995. ${hcmNote ? `[Ghi chú: ${hcmNote}]` : ""}` : ""}
${extraInstructions.trim() ? `- ĐIỀU CHỈNH THÊM TỪ GIÁO VIÊN:\n${extraInstructions.trim()}` : ""}
`.trim();
  };

  const buildAutoIntegrationRequirementsText = () => {
    return `
CHẾ ĐỘ TỰ ĐỘNG NHẬN DIỆN & TÍCH HỢP THÔNG MINH:
- Ngôn ngữ hiển thị của giáo án: ${lessonLang}.
${ppctText ? `- Căn cứ PPCT đã tải:\n"""\n${ppctText.slice(0, 3500)}\n"""\n` : ""}
- Hãy tự phân tích nội dung bài học, môn học (${subject}), khối lớp (Lớp ${grade}) để TỰ ĐỘNG CHỌN từ 2 đến 4 nội dung tích hợp PHÙ HỢP NHẤT, TỰ NHIÊN NHẤT (không nhồi nhét khiên cưỡng) trong 8 loại sau:
  1. "nls" — Năng lực số (TT 02/2025/TT-BGDĐT, mức ${getNlsLevelLabel(grade)}, mã ${getNlsLevelTag(grade)}, vd 1.2.${getNlsLevelTag(grade)}a đặt đầu dòng mục tiêu chung hoặc cuối câu hoạt động).
  2. "ai" — Năng lực Trí tuệ nhân tạo (QĐ 2422/QĐ-BGDĐT), dùng đúng mã lớp ${grade} trong danh sách:
${aiCatalog.map(x => `     + ${x.code}: ${x.desc}`).join("\n")}
  3. "atm" — Giáo dục an toàn mạng toàn diện (CV 6206/BGDĐT-GDPT), đặt mã ATM1–ATM8 đầu dòng mục tiêu hoặc cuối câu hoạt động, liên kết đúng mã TT 02/2025:
${ATM_CATALOG.map(a => `     + ${a.code} (${a.title}) -> NL bắt buộc [${a.requiredNls.join(", ")}], khuyến khích [${a.recommendedNls.join(", ")}]`).join("\n")}
  4. "kt" — Học sinh khuyết tật hòa nhập (chỉ chọn nếu giáo án có nhắc tới HSKT hoặc hoạt động cần điều chỉnh giảm tải).
  5. "qp" — Giáo dục Quốc phòng & An ninh (TT 08/2024/TT-BGDĐT).
  6. "bd" — Biến đổi khí hậu & Phòng chống thiên tai (môi trường, tiết kiệm năng lượng, lối sống xanh).
  7. "ltcm" — Giáo dục lý tưởng cách mạng, đạo đức, lối sống (QĐ 2546 lớp 10–12 hoặc Ngữ văn 6–9).
  8. "hcm" — Giáo dục tư tưởng, đạo đức, phong cách Hồ Chí Minh (QĐ 2995).
${extraInstructions.trim() ? `- Điều chỉnh thêm từ giáo viên:\n${extraInstructions.trim()}` : ""}
`.trim();
  };

  // ==================== RUN AI ====================
  const handleRunAi = async (autoMode = false) => {
    const activeOpts = Object.entries(opts).filter(([_, v]) => v).map(([k]) => k);
    if (!autoMode && activeOpts.length === 0) {
      setError("Vui lòng bật ít nhất 1 nội dung cần tích hợp hoặc dùng nút Tự động nhận diện.");
      return;
    }
    if (mode === "existing" && !docInfo) {
      setError("Chưa có giáo án để tích hợp. Hãy chọn file Word trước.");
      setStep(0);
      return;
    }
    if (mode === "new" && pagePreviews.length === 0) {
      setError("Chưa có trang SGK nào được chọn. Vui lòng tải SGK và bấm OK ở Bước 1.");
      setStep(0);
      return;
    }

    setError(null);
    setLoading(true);
    setProgress(15);
    setProgressStatus(autoMode ? "AI đang tự động phân tích bài dạy & chọn nội dung tích hợp phù hợp..." : "Đang chuẩn bị dữ liệu và kết nối Gemini AI...");

    const timer = setInterval(() => {
      setProgress(p => (p < 88 ? p + Math.floor(Math.random() * 7) + 3 : p));
    }, 450);

    try {
      if (mode === "existing") {
        const parasSlice = docInfo.text.split("\n").slice(0, 180).join("\n");
        const promptText = `Bạn là Chuyên gia Phương pháp Dạy học và Trợ lý AI Giáo dục của MHEdu.
Nhiệm vụ: Phân tích bài giáo án dưới đây và đề xuất các đoạn chèn tích hợp vào đúng các hoạt động (Mục tiêu, Khởi động, Hình thành kiến thức, Luyện tập, Vận dụng).

THÔNG TIN BÀI HỌC:
- Môn học: ${subject}
- Khối lớp: Lớp ${grade}
- Tên bài: ${title || docInfo.fileName}

${autoMode ? buildAutoIntegrationRequirementsText() : buildIntegrationRequirementsText()}

DANH SÁCH CÁC ĐOẠN VĂN CỦA GIÁO ÁN GỐC (có gắn ID đoạn [p1], [p2], ...):
"""
${parasSlice}
"""

QUY TẮC BẮT BUỘC:
1. Đọc kỹ các đoạn văn, xác định đúng các phần Mục tiêu, Khởi động, Hình thành kiến thức, Luyện tập, Vận dụng.
2. Chọn đúng mã đoạn "a" (ví dụ "p4", "p12", "p25") cần chèn thêm nội dung.
3. Trả về đúng JSON thuần túy:
{
  "items": [
    { "a": "p4", "w": "below", "k": "nls", "t": "1.1.TC2b – Nội dung tích hợp cụ thể..." }
  ]
}
Trong đó "w" là "below" hoặc "end", "k" thuộc ["nls","ai","kt","qp","bd","ltcm","atm","hcm"]. Tạo 6 đến 16 đoạn chèn phân bổ hợp lý cho tất cả các nội dung đã bật.`;

        const parsed = await callGeminiJson([{ text: promptText }]);
        if (!parsed || !Array.isArray(parsed.items) || parsed.items.length === 0) {
          throw new Error("AI trả về dữ liệu trống.");
        }
        const mappedItems = parsed.items.map((it, idx) => ({
          id: "item-" + idx,
          a: it.a || "p1",
          w: it.w || "below",
          k: it.k || "nls",
          t: it.t || "",
          on: true
        }));
        if (autoMode) {
          const used = new Set(mappedItems.map(x => x.k));
          setOpts({
            nls: used.has("nls"), ai: used.has("ai"), kt: used.has("kt"), qp: used.has("qp"),
            bd: used.has("bd"), ltcm: used.has("ltcm"), atm: used.has("atm"), hcm: used.has("hcm")
          });
        }
        setItems(mappedItems);
        setProgress(100);
        setStep(2);
      } else {
        const extractedTexts = pagePreviews.map(p => `--- TRANG ${p.pageNum} ---\n${p.text || "(Trang ảnh chụp - xem hình đính kèm)"}`).join("\n\n");
        const promptText = `Bạn là Chuyên gia Phương pháp Dạy học (Công văn 5512/BGDĐT) và Trợ lý AI Giáo dục MHEdu.
Nhiệm vụ: Dựa vào nội dung các trang Sách giáo khoa (SGK) đính kèm, hãy soạn một Kế hoạch bài dạy (Giáo án) hoàn chỉnh, chuẩn sư phạm và tích hợp đầy đủ các nội dung được yêu cầu.

THÔNG TIN BÀI DẠY:
- Môn học: ${subject}
- Khối lớp: Lớp ${grade}
- Tên bài: ${title || "Bài học từ SGK"}
- Thời lượng: ${periods} tiết
${activityList.filter(a => a.trim()).length ? `- Kế hoạch hoạt động theo thứ tự:\n${activityList.filter(a => a.trim()).map((a, i) => `  ${i + 1}. ${a.trim()}`).join("\n")}` : ""}

${autoMode ? buildAutoIntegrationRequirementsText() : buildIntegrationRequirementsText()}

NỘI DUNG CHỮ TRÍCH XUẤT TỪ SGK:
"""
${extractedTexts.slice(0, 8000)}
"""

Trả về đúng JSON thuần túy:
{
  "title": "Tên bài học chuẩn",
  "sections": [
    {
      "heading": "I. MỤC TIÊU",
      "blocks": [
        { "type": "normal", "text": "1. Kiến thức: ..." },
        { "type": "nls", "text": "1.1.TC2b – ..." }
      ]
    },
    { "heading": "II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU", "blocks": [...] },
    { "heading": "III. TIẾN TRÌNH DẠY HỌC", "blocks": [...] }
  ]
}
Trong đó "type" chỉ nhận: "normal", "nls", "ai", "kt", "qp", "bd", "ltcm", "atm", "hcm".`;

        const parts = [{ text: promptText }];
        for (const p of pagePreviews.slice(0, 6)) {
          if (p.dataUrl && p.dataUrl.includes(",")) {
            parts.push({ inlineData: { mimeType: "image/jpeg", data: p.dataUrl.split(",")[1] } });
          }
        }
        const parsed = await callGeminiJson(parts);
        if (!parsed || !Array.isArray(parsed.sections) || parsed.sections.length === 0) {
          throw new Error("AI trả về cấu trúc giáo án chưa hợp lệ.");
        }
        const builtLesson = {
          title: parsed.title || title || "Kế hoạch bài dạy tích hợp",
          subject, grade, periods,
          sections: parsed.sections.map((sec, sIdx) => ({
            id: "sec-" + sIdx,
            heading: sec.heading || "",
            blocks: (sec.blocks || []).map((b, bIdx) => ({
              id: `blk-${sIdx}-${bIdx}`,
              type: b.type || "normal",
              text: b.text || ""
            }))
          }))
        };
        if (autoMode) {
          const used = new Set(builtLesson.sections.flatMap(s => s.blocks.map(b => b.type)).filter(t => t !== "normal"));
          setOpts({
            nls: used.has("nls"), ai: used.has("ai"), kt: used.has("kt"), qp: used.has("qp"),
            bd: used.has("bd"), ltcm: used.has("ltcm"), atm: used.has("atm"), hcm: used.has("hcm")
          });
        }
        setNewLessonData(builtLesson);
        setProgress(100);
        setStep(2);
      }
    } catch (e) {
      setError("AI không thể phản hồi: " + (e?.message || "Kiểm tra lại kết nối mạng hoặc Gemini API Key."));
    } finally {
      clearInterval(timer);
      setLoading(false);
    }
  };

  // ==================== BÁO CÁO TÍCH HỢP ====================
  const integrationSummary = useMemo(() => {
    const list = mode === "existing"
      ? items.filter(i => i.on).map(i => ({ k: i.k, t: i.t }))
      : (newLessonData?.sections || []).flatMap(s => s.blocks.filter(b => b.type !== "normal").map(b => ({ k: b.type, t: b.text })));

    const byCat = {};
    for (const k of Object.keys(CATEGORIES)) {
      byCat[k] = list.filter(x => x.k === k);
    }
    const nlsCodes = extractCodesFromText((byCat.nls || []).map(x => x.t).join(" "), /\b\d+\.\d+\.[A-Z]{2}\d+[a-z]?\b/g);
    const aiCodes = extractCodesFromText((byCat.ai || []).map(x => x.t).join(" "), /\b(?:\d+\.[A-D]\d+\.(?:MR)?\d+|[A-D]\d+)\b/g);
    const atmCodes = extractCodesFromText((byCat.atm || []).map(x => x.t).join(" "), /\bATM[1-8]\b/g);

    return { total: list.length, byCat, nlsCodes, aiCodes, atmCodes };
  }, [mode, items, newLessonData]);

  const handleCopyReport = () => {
    const s = integrationSummary;
    const rpt = `📋 BÁO CÁO TÍCH HỢP GIÁO ÁN
• Nền tảng: MHEdu AI
• Bài học: ${title || docInfo?.fileName || newLessonData?.title || "Kế hoạch bài dạy"}
• Môn học: ${subject} - Khối lớp: Lớp ${grade}
${ppctName ? `• Nguồn PPCT: ${ppctName}\n` : ""}• Tổng số vị trí tích hợp: ${s.total} vị trí
  + Năng lực số (NLS): ${s.byCat.nls.length} vị trí${s.nlsCodes.length ? ` (${s.nlsCodes.join(", ")})` : ""}
  + Năng lực AI (QĐ 2422): ${s.byCat.ai.length} vị trí${s.aiCodes.length ? ` (${s.aiCodes.join(", ")})` : ""}
  + Học sinh khuyết tật: ${s.byCat.kt.length} vị trí
  + GDQP&AN (TT 08/2024): ${s.byCat.qp.length} vị trí
  + BĐKH-PCTT: ${s.byCat.bd.length} vị trí
  + GD LTCM, ĐĐLS: ${s.byCat.ltcm.length} vị trí
  + An toàn mạng (CV 6206): ${s.byCat.atm.length} vị trí${s.atmCodes.length ? ` (${s.atmCodes.join(", ")})` : ""}
  + Tư tưởng HCM (QĐ 2995): ${s.byCat.hcm.length} vị trí
• Thời gian: ${new Date().toLocaleString("vi-VN")}`;

    navigator.clipboard.writeText(rpt).then(() => {
      setCopiedReport(true);
      setTimeout(() => setCopiedReport(false), 3000);
    }).catch(() => alert("Không thể sao chép tự động."));
  };

  // ==================== TẢI WORD ====================
  const handleDownloadWord = async () => {
    setDownloading(true);
    try {
      let blob;
      let outName;
      if (mode === "existing") {
        if (!docInfo || items.length === 0) return;
        blob = await generatePatchedDocx(docInfo, items.filter(it => it.on), colorEnabled);
        outName = `${docInfo.fileName.replace(/\.docx$/i, "")}_Tich_Hop.docx`;
      } else {
        if (!newLessonData) return;
        blob = await generateNewLessonDocx(newLessonData, colorEnabled);
        outName = `Giao_an_${subject}_Lop_${grade}_Tich_Hop.docx`.replace(/\s+/g, "_");
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = outName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      alert("Lỗi khi tạo file Word: " + (err.message || "Không xác định"));
    } finally {
      setDownloading(false);
    }
  };

  const filteredItems = useMemo(() => {
    if (activeFilter === "all") return items;
    return items.filter(it => it.k === activeFilter);
  }, [items, activeFilter]);

  const activeCount = useMemo(() => items.filter(it => it.on).length, [items]);
  const step1Ready = mode === "existing" ? !!docInfo : pagePreviews.length > 0;
  const step3Ready = mode === "existing" ? items.length > 0 : !!newLessonData;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans pb-24">
      <div className="max-w-6xl mx-auto px-4 pt-6">

        {/* Header & 2-Tab Switcher */}
        <div className="bg-white rounded-3xl p-6 shadow-sm border border-slate-200/80 mb-5">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 rounded-2xl bg-gradient-to-tr from-purple-600 to-indigo-500 flex items-center justify-center text-white text-2xl shadow-md shrink-0">🧩</div>
            <div>
              <div className="flex items-center gap-2.5 flex-wrap">
                <h1 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">TÍCH HỢP TỔNG HỢP VÀO GIÁO ÁN</h1>
                <span className="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 text-xs font-bold uppercase">✨ BẢN ĐẦY ĐỦ</span>
              </div>
              <p className="text-xs sm:text-sm text-slate-500 mt-1 font-medium">
                Hỗ trợ 8 nội dung tích hợp chuẩn Bộ GD&ĐT (NLS, NL AI, HSKT, GDQPAN, BĐKH, LTCM-ĐĐLS, An toàn mạng CV 6206, Tư tưởng HCM QĐ 2995).
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-5 pt-4 border-t border-slate-100">
            <button
              onClick={() => switchMode("existing")}
              className={`p-3 rounded-2xl border-2 text-center font-extrabold text-sm transition-all ${
                mode === "existing" ? "border-purple-600 bg-purple-600 text-white shadow-sm" : "border-slate-200 bg-white text-slate-600 hover:border-purple-300"
              }`}
            >
              📄 Tích hợp vào giáo án có sẵn
            </button>
            <button
              onClick={() => switchMode("new")}
              className={`p-3 rounded-2xl border-2 text-center font-extrabold text-sm transition-all ${
                mode === "new" ? "border-purple-600 bg-purple-600 text-white shadow-sm" : "border-slate-200 bg-white text-slate-600 hover:border-purple-300"
              }`}
            >
              📖 Tạo một giáo án mới có tích hợp
            </button>
          </div>
        </div>

        {/* Stepper */}
        <div className="grid grid-cols-3 gap-2 sm:gap-4 mb-6">
          {[
            { id: 0, num: "1", title: mode === "existing" ? "Giáo án" : "Bài dạy & SGK" },
            { id: 1, num: "2", title: "Nội dung tích hợp" },
            { id: 2, num: "3", title: mode === "existing" ? "Tải Word" : "Xem · Tải Word" }
          ].map(s => {
            const isActive = step === s.id;
            const isDone = (s.id === 0 && step1Ready) || (s.id === 1 && step3Ready) || (s.id === 2 && step3Ready);
            const canClick = s.id === 0 || (s.id === 1 && step1Ready) || (s.id === 2 && step3Ready);
            return (
              <button
                key={s.id}
                onClick={() => canClick && setStep(s.id)}
                className={`p-3 rounded-2xl border text-center transition-all flex items-center justify-center gap-2 ${
                  isActive ? "bg-purple-600 text-white border-purple-600 shadow-md" : isDone ? "bg-white text-emerald-700 border-emerald-300" : "bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed"
                }`}
              >
                <span className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-black ${
                  isActive ? "bg-white text-purple-700" : isDone ? "bg-emerald-600 text-white" : "bg-slate-200 text-slate-500"
                }`}>{isDone && !isActive ? "✓" : s.num}</span>
                <span className="font-bold text-sm">{s.title}</span>
              </button>
            );
          })}
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-2xl text-sm mb-6 flex items-start gap-3">
            <span className="text-xl">⚠️</span>
            <div className="flex-1 font-medium">{error}</div>
            <button onClick={() => setError(null)} className="text-red-400 hover:text-red-700 font-bold">✕</button>
          </div>
        )}

        {/* ==================== BƯỚC 1: MODE 1 ==================== */}
        {step === 0 && mode === "existing" && (
          <div className="bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6">
            <div>
              <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2"><span>📄</span> Giáo án cần tích hợp</h2>
              <p className="text-xs text-slate-500">Tải file Word và kiểm tra thông tin bài dạy</p>
            </div>

            {!docInfo ? (
              <div
                onDragOver={e => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={e => { e.preventDefault(); setIsDragging(false); if (e.dataTransfer.files?.[0]) handleFileUpload(e.dataTransfer.files[0]); }}
                onClick={() => document.getElementById("file-lesson-input").click()}
                className={`border-2 border-dashed rounded-3xl p-8 sm:p-12 text-center transition-all cursor-pointer ${
                  isDragging ? "border-purple-500 bg-purple-50" : "border-purple-300 bg-purple-50/40 hover:bg-purple-50"
                }`}
              >
                <input id="file-lesson-input" type="file" accept=".docx" className="hidden" onChange={e => e.target.files?.[0] && handleFileUpload(e.target.files[0])} />
                <div className="w-12 h-12 rounded-full bg-white text-purple-600 flex items-center justify-center text-2xl mx-auto mb-3 shadow-sm">📤</div>
                <div className="inline-block px-5 py-2.5 bg-purple-600 text-white font-bold text-sm rounded-xl shadow-sm mb-2">Chọn file giáo án Word (.docx)</div>
                <p className="text-xs font-semibold text-slate-700">hoặc kéo thả file vào khung này</p>
                <p className="text-xs text-slate-500 mt-1">1 bài hoặc cả tuần (tối đa khoảng 100 trang). File .doc cũ: mở bằng Word rồi Lưu thành .docx.</p>
              </div>
            ) : (
              <div className="bg-emerald-50/60 rounded-2xl p-5 border-2 border-dashed border-emerald-400 text-center">
                <h4 className="font-bold text-slate-900 text-base">{docInfo.fileName}</h4>
                <div className="flex items-center justify-center gap-3 flex-wrap mt-2 text-xs font-semibold text-slate-600">
                  <span>{docInfo.paraCount} đoạn</span> • <span>{docInfo.tables} bảng</span> • <span>{docInfo.images} hình</span> • <span>{docInfo.maths} công thức</span>
                </div>
                <button onClick={() => { setDocInfo(null); setItems([]); }} className="mt-3 text-xs font-bold text-blue-700 bg-white border border-blue-200 hover:bg-blue-50 px-4 py-1.5 rounded-xl">
                  📤 CHỌN FILE KHÁC
                </button>
              </div>
            )}

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1 uppercase">Lớp</label>
                <select value={grade} onChange={e => setGrade(Number(e.target.value))} className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold">
                  {[1,2,3,4,5,6,7,8,9,10,11,12].map(g => <option key={g} value={g}>Lớp {g}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1 uppercase">Môn học (Không bắt buộc)</label>
                <select value={subject} onChange={e => setSubject(e.target.value)} className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold">
                  {SUBJECTS.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1 uppercase">Tên bài (Không bắt buộc)</label>
                <input type="text" value={title} onChange={e => setTitle(e.target.value)} placeholder="VD: BÀI 2: TẬP HỢP" className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold" />
              </div>
            </div>

            {!docInfo && (
              <div className="bg-red-50 border border-red-200 text-red-700 rounded-2xl p-4 text-xs font-medium">
                <strong>Cần bổ sung trước khi sang bước "Nội dung tích hợp":</strong>
                <ul className="list-disc ml-5 mt-1"><li>Chọn file giáo án Word (.docx).</li></ul>
              </div>
            )}

            <div className="flex flex-wrap justify-end gap-3">
              <button
                disabled={!docInfo}
                onClick={() => handleRunAi(true)}
                className="px-6 py-3 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 disabled:from-slate-300 disabled:to-slate-300 text-white font-extrabold rounded-xl text-sm shadow-sm"
              >
                ✨ TỰ ĐỘNG NHẬN DIỆN & TÍCH HỢP NGAY
              </button>
              <button
                disabled={!docInfo}
                onClick={() => setStep(1)}
                className="px-6 py-3 bg-purple-600 hover:bg-purple-700 disabled:bg-slate-300 text-white font-bold rounded-xl text-sm shadow-sm"
              >
                TIẾP: NỘI DUNG TÍCH HỢP →
              </button>
            </div>
          </div>
        )}

        {/* ==================== BƯỚC 1: MODE 2 (TẠO MỚI TỪ SGK) ==================== */}
        {step === 0 && mode === "new" && (
          <div className="bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6">
            <div>
              <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2"><span>📖</span> Bước 1: Chọn Sách giáo khoa & Trang bài dạy</h2>
              <p className="text-xs text-slate-500">Tải file SGK (.pdf) hoặc ảnh chụp trang sách. SGK được lưu trên máy để dùng lại nhanh.</p>
            </div>

            {savedTextbooks.length > 0 && (
              <div className="bg-slate-50 rounded-2xl p-4 border border-slate-200">
                <div className="text-xs font-bold text-slate-600 uppercase mb-2.5">📚 SGK đã có trên máy</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                  {savedTextbooks.map(tb => (
                    <div key={tb.id} className="bg-white p-3 rounded-xl border border-slate-200 flex items-center justify-between gap-2">
                      <div className="min-w-0">
                        <div className="text-xs font-bold text-slate-800 truncate">{tb.name}</div>
                        <div className="text-[11px] text-slate-500">{tb.numPages} trang • Lớp {tb.grade} • {tb.subject}</div>
                      </div>
                      <div className="flex items-center gap-1.5 shrink-0">
                        <button onClick={() => handleUseSavedSgk(tb)} className="px-2.5 py-1 bg-purple-600 text-white text-xs font-bold rounded-lg">DÙNG</button>
                        <button onClick={() => deleteTextbookFromDb(tb.id).then(() => listSavedTextbooks().then(setSavedTextbooks))} className="px-2 py-1 text-slate-400 hover:text-red-600 text-xs font-bold">✕</button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="border-2 border-dashed border-purple-300 bg-purple-50/40 hover:bg-purple-50 rounded-2xl p-6 text-center cursor-pointer" onClick={() => document.getElementById("sgk-file-input").click()}>
              <input id="sgk-file-input" type="file" accept=".pdf,image/*" multiple className="hidden" onChange={e => handleSgkUpload(e.target.files)} />
              <div className="text-2xl mb-1">📤</div>
              <div className="text-sm font-bold text-slate-800">
                {sgkFile ? `Đang chọn: ${sgkFile.name} (${sgkFile.numPages} trang)` : "Tải SGK khác (file PDF hoặc ảnh chụp)"}
              </div>
            </div>

            {sgkFile && sgkFile.type === "pdf" && (
              <div className="bg-purple-50/60 border border-purple-200 rounded-2xl p-4 flex flex-wrap items-end gap-3">
                <div>
                  <label className="block text-xs font-bold text-purple-900 mb-1">Từ trang</label>
                  <input type="number" min={1} max={sgkFile.numPages} value={pageFrom} onChange={e => setPageFrom(Number(e.target.value))} className="w-24 p-2 bg-white border border-purple-300 rounded-xl text-sm font-bold text-center" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-purple-900 mb-1">Đến trang (tối đa 20)</label>
                  <input type="number" min={1} max={sgkFile.numPages} value={pageTo} onChange={e => setPageTo(Number(e.target.value))} className="w-24 p-2 bg-white border border-purple-300 rounded-xl text-sm font-bold text-center" />
                </div>
                <button onClick={() => renderPdfRange(sgkFile.pdfDoc, pageFrom, pageTo)} disabled={renderingPages} className="px-5 py-2.5 bg-purple-600 text-white font-bold text-sm rounded-xl">
                  {renderingPages ? "Đang tải..." : "OK • Xem trang"}
                </button>
              </div>
            )}

            {pagePreviews.length > 0 && (
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                {pagePreviews.map(p => (
                  <div key={p.pageNum} className="border border-slate-200 rounded-xl overflow-hidden bg-white">
                    <img src={p.dataUrl} alt={`Trang ${p.pageNum}`} className="w-full h-40 object-cover object-top" />
                    <div className="text-center py-1 text-xs font-bold text-slate-600 bg-slate-50 border-t">Trang {p.pageNum}</div>
                  </div>
                ))}
              </div>
            )}

            <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 pt-4 border-t">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1 uppercase">Lớp</label>
                <select value={grade} onChange={e => setGrade(Number(e.target.value))} className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold">
                  {[1,2,3,4,5,6,7,8,9,10,11,12].map(g => <option key={g} value={g}>Lớp {g}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1 uppercase">Môn học</label>
                <select value={subject} onChange={e => setSubject(e.target.value)} className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold">
                  {SUBJECTS.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1 uppercase">Số tiết</label>
                <select value={periods} onChange={e => setPeriods(Number(e.target.value))} className="w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold">
                  {[1,2,3,4,5,6,7,8,9,10].map(p => <option key={p} value={p}>{p} tiết</option>)}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1 uppercase">Tên bài <span className="text-red-500">*</span></label>
                <input type="text" value={title} onChange={e => setTitle(e.target.value)} placeholder="VD: Bài 1..." className={`w-full p-2.5 bg-white border rounded-xl text-sm font-semibold ${!title.trim() ? "border-red-300" : "border-slate-300"}`} />
              </div>
            </div>

            <div className="space-y-2">
              <div className="text-xs font-bold text-slate-700">Kế hoạch hoạt động (không bắt buộc)</div>
              <div className="text-[11px] text-slate-500">Để trống = AI tự thiết kế theo nội dung bài. Ghi hoạt động thì AI soạn đúng theo thứ tự đó.</div>
              {activityList.map((act, idx) => (
                <div key={idx} className="flex items-center gap-2">
                  <input
                    type="text"
                    value={act}
                    onChange={e => setActivityList(activityList.map((x, i) => i === idx ? e.target.value : x))}
                    placeholder={`Hoạt động ${idx + 1}`}
                    className="flex-1 p-2 bg-white border border-slate-300 rounded-xl text-xs"
                  />
                  <button type="button" onClick={() => setActivityList(activityList.filter((_, i) => i !== idx))} className="text-slate-400 hover:text-red-600 px-2 font-bold">🗑</button>
                </div>
              ))}
              <button
                type="button"
                onClick={() => setActivityList([...activityList, ""])}
                className="text-xs font-extrabold text-blue-700 hover:underline"
              >
                + THÊM HOẠT ĐỘNG
              </button>
            </div>

            {(pagePreviews.length === 0 || !title.trim()) && (
              <div className="bg-red-50 border border-red-200 text-red-700 rounded-2xl p-4 text-xs font-medium">
                <strong>Cần bổ sung trước khi sang bước "Nội dung tích hợp":</strong>
                <ul className="list-disc ml-5 mt-1">
                  {pagePreviews.length === 0 && <li>Chọn SGK và các trang của bài (file PDF hoặc ảnh chụp).</li>}
                  {!title.trim() && <li>Nhập tên bài.</li>}
                </ul>
              </div>
            )}

            <div className="flex flex-wrap justify-end gap-3">
              <button
                disabled={pagePreviews.length === 0 || !title.trim()}
                onClick={() => handleRunAi(true)}
                className="px-6 py-3 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 disabled:from-slate-300 disabled:to-slate-300 text-white font-extrabold rounded-xl text-sm shadow-sm"
              >
                ✨ TỰ ĐỘNG NHẬN DIỆN & SOẠN GIÁO ÁN NGAY
              </button>
              <button
                disabled={pagePreviews.length === 0 || !title.trim()}
                onClick={() => setStep(1)}
                className="px-6 py-3 bg-purple-600 hover:bg-purple-700 disabled:bg-slate-300 text-white font-bold rounded-xl text-sm"
              >
                TIẾP: NỘI DUNG TÍCH HỢP →
              </button>
            </div>
          </div>
        )}

        {/* ==================== BƯỚC 2: NỘI DUNG TÍCH HỢP (2-COLUMN LAYOUT) ==================== */}
        {step === 1 && (
          <div className="bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-purple-600 text-white flex items-center justify-center text-lg font-bold">📚</div>
              <div>
                <h2 className="text-lg font-bold text-slate-900">Nội dung tích hợp</h2>
                <p className="text-xs text-slate-500">Chọn các nội dung muốn lồng ghép vào giáo án</p>
              </div>
            </div>

            {/* 2-Column Grid: Left = 8 Selectable Cards, Right = Active Configurations */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
              {/* Left Column: 8 Cards */}
              <div className="lg:col-span-5 space-y-3">
                {INTEGRATION_CARDS.map(card => {
                  const checked = !!opts[card.key];
                  return (
                    <div
                      key={card.key}
                      onClick={() => setOpts(prev => ({ ...prev, [card.key]: !prev[card.key] }))}
                      style={{ borderColor: checked ? card.dot : "#e2e8f0", backgroundColor: checked ? `${card.dot}0A` : "#fff" }}
                      className="p-3.5 rounded-2xl border-2 cursor-pointer transition-all hover:shadow-xs"
                    >
                      <div className="flex items-start gap-2.5">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => {}}
                          style={{ accentColor: card.dot }}
                          className="w-4 h-4 mt-0.5 rounded cursor-pointer shrink-0"
                        />
                        <div className="min-w-0">
                          <div style={{ color: card.dot }} className="font-extrabold text-sm leading-snug">{card.title}</div>
                          <div className="text-[11px] text-slate-500 mt-1 leading-relaxed">{card.desc}</div>
                          {checked && (
                            <div className="flex items-center gap-1.5 mt-2 text-xs font-bold text-slate-700">
                              <span>Màu chữ:</span>
                              <span style={{ backgroundColor: card.dot }} className="w-4 h-4 rounded-full inline-block border border-white shadow-xs ring-1 ring-slate-300" />
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Right Column: Config Panels for Checked Items */}
              <div className="lg:col-span-7 space-y-4">
                {!Object.values(opts).some(Boolean) && (
                  <div className="p-8 rounded-2xl border-2 border-dashed border-slate-200 text-center text-xs text-slate-400">
                    Tích chọn ít nhất một nội dung ở cột bên trái để cấu hình chi tiết.
                  </div>
                )}

                {/* 1. Config NLS + PPCT */}
                {opts.nls && (
                  <div className="p-4 rounded-2xl border border-violet-300 bg-violet-50/40 space-y-3">
                    <div>
                      <div className="text-sm font-extrabold text-slate-800">Phân phối chương trình (không bắt buộc)</div>
                      <div className="text-xs text-slate-500 mt-0.5">Chỉ tải khi PPCT của trường CÓ cột Năng lực số / Năng lực AI / An toàn mạng. File Excel (.xlsx) hoặc Word (.docx).</div>
                      <div className="mt-2 flex items-center gap-2">
                        <label className="px-3.5 py-2 bg-white border border-violet-300 text-violet-700 text-xs font-bold rounded-xl cursor-pointer hover:bg-violet-50">
                          📤 {ppctName ? `Đã tải: ${ppctName}` : "CHỌN FILE PPCT"}
                          <input type="file" accept=".docx,.xlsx" className="hidden" onChange={e => e.target.files?.[0] && handlePpctUpload(e.target.files[0])} />
                        </label>
                        {ppctName && <button onClick={() => { setPpctName(""); setPpctText(""); }} className="text-xs text-red-600 font-bold">Xóa</button>}
                      </div>
                    </div>
                    <div className="pt-3 border-t border-violet-200/70">
                      <div className="text-sm font-extrabold text-slate-800">Chọn mã Năng lực số cụ thể (không bắt buộc)</div>
                      <div className="text-xs text-slate-500 mb-2">Để trống = AI tự chọn mã phù hợp trong khung của lớp {grade}. Chọn mã = AI chỉ dùng các mã này (tối đa 10).</div>
                      <select
                        value=""
                        onChange={e => {
                          const code = e.target.value;
                          if (!code || selectedNlsItems.some(x => x.code === code) || selectedNlsItems.length >= 10) return;
                          const found = nlsCatalog.find(x => x.code === code);
                          if (found) setSelectedNlsItems([...selectedNlsItems, { ...found }]);
                        }}
                        className="w-full p-2.5 bg-white border border-violet-300 rounded-xl text-xs font-semibold"
                      >
                        <option value="">Mã Năng lực số ({getNlsLevelLabel(grade)})</option>
                        {nlsCatalog.map(c => <option key={c.code} value={c.code}>{c.code} — {c.desc}</option>)}
                      </select>
                      {selectedNlsItems.length > 0 && (
                        <div className="mt-2.5 space-y-2">
                          {selectedNlsItems.map((it, idx) => (
                            <div key={it.code} className="bg-white p-2.5 rounded-xl border border-violet-200">
                              <div className="flex items-center justify-between mb-1">
                                <span className="text-xs font-extrabold text-violet-700">{it.code} — Yêu cầu cần đạt (Biểu hiện)</span>
                                <button onClick={() => setSelectedNlsItems(selectedNlsItems.filter((_, i) => i !== idx))} className="text-xs text-slate-400 hover:text-red-600 font-bold">✕</button>
                              </div>
                              <input
                                type="text"
                                value={it.desc}
                                onChange={e => setSelectedNlsItems(selectedNlsItems.map((x, i) => i === idx ? { ...x, desc: e.target.value } : x))}
                                className="w-full p-1.5 border border-slate-200 rounded-lg text-xs"
                              />
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* 2. Config NL AI */}
                {opts.ai && (
                  <div className="p-4 rounded-2xl border border-red-300 bg-red-50/40 space-y-2.5">
                    <div className="text-sm font-extrabold text-red-700">Chọn yêu cầu Năng lực AI cụ thể (không bắt buộc)</div>
                    <div className="text-xs text-slate-500">Để trống = AI tự chọn yêu cầu phù hợp trong khung lớp {grade}. Chọn mã = AI chỉ dùng các mã này (tối đa 10).</div>
                    <select
                      value=""
                      onChange={e => {
                        const code = e.target.value;
                        if (!code || selectedAiItems.some(x => x.code === code) || selectedAiItems.length >= 10) return;
                        const found = aiCatalog.find(x => x.code === code);
                        if (found) setSelectedAiItems([...selectedAiItems, { ...found }]);
                      }}
                      className="w-full p-2.5 bg-white border border-red-300 rounded-xl text-xs font-semibold"
                    >
                      <option value="">Yêu cầu cần đạt Năng lực AI (lớp {grade})</option>
                      {aiCatalog.map(c => <option key={c.code} value={c.code}>{c.code} — {c.desc}</option>)}
                    </select>
                    {selectedAiItems.length > 0 && (
                      <div className="space-y-2">
                        {selectedAiItems.map((it, idx) => (
                          <div key={it.code} className="bg-white p-2.5 rounded-xl border border-red-200">
                            <div className="flex items-center justify-between mb-1">
                              <span className="text-xs font-extrabold text-red-700">{it.code} — Yêu cầu cần đạt</span>
                              <button onClick={() => setSelectedAiItems(selectedAiItems.filter((_, i) => i !== idx))} className="text-xs text-slate-400 hover:text-red-600 font-bold">✕</button>
                            </div>
                            <input
                              type="text"
                              value={it.desc}
                              onChange={e => setSelectedAiItems(selectedAiItems.map((x, i) => i === idx ? { ...x, desc: e.target.value } : x))}
                              className="w-full p-1.5 border border-slate-200 rounded-lg text-xs"
                            />
                          </div>
                        ))}
                      </div>
                    )}
                    <div className="text-[11px] text-slate-500">Phân phối chương trình (nếu có) tải ở khung Năng lực số phía trên — dùng chung.</div>
                  </div>
                )}

                {/* 3. Config HSKT */}
                {opts.kt && (
                  <div className="p-4 rounded-2xl border border-emerald-300 bg-emerald-50/40 space-y-2.5">
                    <div className="text-sm font-extrabold text-emerald-800">Học sinh khuyết tật trong lớp</div>
                    <div className="text-xs text-slate-500">Chọn dạng khuyết tật (tối đa 5). Chỉ ghi tên lớp, KHÔNG ghi tên học sinh.</div>
                    <div className="space-y-2">
                      {disabilityRows.map((row, idx) => (
                        <div key={idx} className="flex flex-wrap sm:flex-nowrap items-center gap-2">
                          <select
                            value={row.type}
                            onChange={e => setDisabilityRows(disabilityRows.map((r, i) => i === idx ? { ...r, type: e.target.value } : r))}
                            className="flex-1 p-2 bg-white border border-emerald-300 rounded-xl text-xs font-semibold"
                          >
                            {DISABILITY_TYPES.map(d => <option key={d.key} value={d.label}>{d.label}</option>)}
                          </select>
                          <input
                            type="text"
                            value={row.className}
                            onChange={e => setDisabilityRows(disabilityRows.map((r, i) => i === idx ? { ...r, className: e.target.value } : r))}
                            placeholder="Lớp (không bắt buộc)"
                            className="w-36 p-2 bg-white border border-slate-300 rounded-xl text-xs"
                          />
                          {disabilityRows.length > 1 && (
                            <button onClick={() => setDisabilityRows(disabilityRows.filter((_, i) => i !== idx))} className="text-slate-400 hover:text-red-600 px-2 font-bold">🗑</button>
                          )}
                        </div>
                      ))}
                    </div>
                    {disabilityRows.length < 5 && (
                      <button
                        onClick={() => setDisabilityRows([...disabilityRows, { type: "Khuyết tật chung", className: "" }])}
                        className="text-xs font-extrabold text-emerald-700 hover:underline"
                      >
                        + THÊM DẠNG KHUYẾT TẬT
                      </button>
                    )}
                    <input
                      type="text"
                      value={disabilityNote}
                      onChange={e => setDisabilityNote(e.target.value)}
                      placeholder="Ghi chú biện pháp hỗ trợ thêm (nếu có)..."
                      className="w-full p-2 bg-white border border-emerald-200 rounded-xl text-xs"
                    />
                  </div>
                )}

                {/* 4. Config GDQPAN */}
                {opts.qp && (
                  <div className="p-4 rounded-2xl border border-blue-300 bg-blue-50/40 space-y-2">
                    <div className="text-sm font-extrabold text-blue-800">Giáo dục quốc phòng & an ninh (TT 08/2024)</div>
                    <div className="flex flex-wrap gap-1.5">
                      {GDQPAN_SUGGESTIONS.map((s, i) => (
                        <button key={i} type="button" onClick={() => setQpNote(p => p ? p + "; " + s : s)} className="text-[11px] bg-white border border-blue-200 text-blue-700 px-2 py-1 rounded-lg">+ {s}</button>
                      ))}
                    </div>
                    <textarea value={qpNote} onChange={e => setQpNote(e.target.value)} rows={2} placeholder="Ghi chú lồng ghép GDQPAN (không bắt buộc)..." className="w-full p-2 bg-white border border-blue-200 rounded-xl text-xs" />
                  </div>
                )}

                {/* 5. Config BĐKH */}
                {opts.bd && (
                  <div className="p-4 rounded-2xl border border-orange-300 bg-orange-50/40 space-y-2">
                    <div className="text-sm font-extrabold text-orange-800">Biến đổi khí hậu - Phòng chống thiên tai</div>
                    <div className="flex flex-wrap gap-1.5">
                      {BDKH_SUGGESTIONS.map((s, i) => (
                        <button key={i} type="button" onClick={() => setBdNote(p => p ? p + "; " + s : s)} className="text-[11px] bg-white border border-orange-200 text-orange-700 px-2 py-1 rounded-lg">+ {s}</button>
                      ))}
                    </div>
                    <textarea value={bdNote} onChange={e => setBdNote(e.target.value)} rows={2} placeholder="Ghi chú lồng ghép BĐKH-PCTT (không bắt buộc)..." className="w-full p-2 bg-white border border-orange-200 rounded-xl text-xs" />
                  </div>
                )}

                {/* 6. Config GD LTCM, ĐĐLS */}
                {opts.ltcm && (
                  <div className="p-4 rounded-2xl border border-yellow-400 bg-yellow-50/50 space-y-2">
                    <div className="text-sm font-extrabold text-yellow-800">Giáo dục lý tưởng cách mạng, đạo đức, lối sống (QĐ 2546)</div>
                    <textarea value={ltcmNote} onChange={e => setLtcmNote(e.target.value)} rows={2} placeholder="Yêu cầu lồng ghép lý tưởng cách mạng, đạo đức, lối sống (không bắt buộc)..." className="w-full p-2 bg-white border border-yellow-300 rounded-xl text-xs" />
                  </div>
                )}

                {/* 7. Config An toàn mạng (CV 6206) */}
                {opts.atm && (
                  <div className="p-4 rounded-2xl border border-teal-300 bg-teal-50/40 space-y-2.5">
                    <div className="text-sm font-extrabold text-teal-800">Chọn mã An toàn mạng cụ thể (không bắt buộc)</div>
                    <div className="text-xs text-slate-500">Để trống = AI tự chọn 1–2 mã phù hợp với bài. Chọn mã = AI chỉ dùng các mã này (tối đa 10).</div>
                    <select
                      value=""
                      onChange={e => {
                        const code = e.target.value;
                        if (!code || selectedAtmItems.some(x => x.code === code) || selectedAtmItems.length >= 10) return;
                        const found = ATM_CATALOG.find(x => x.code === code);
                        if (found) setSelectedAtmItems([...selectedAtmItems, { ...found }]);
                      }}
                      className="w-full p-2.5 bg-white border border-teal-300 rounded-xl text-xs font-semibold"
                    >
                      <option value="">Mã An toàn mạng (CV 6206)</option>
                      {ATM_CATALOG.map(c => <option key={c.code} value={c.code}>{c.code} — {c.title}</option>)}
                    </select>
                    {selectedAtmItems.length > 0 && (
                      <div className="space-y-1.5">
                        {selectedAtmItems.map((it, idx) => (
                          <div key={it.code} className="bg-white p-2 rounded-xl border border-teal-200 flex items-start justify-between gap-2 text-xs">
                            <div>
                              <span className="font-extrabold text-teal-700">{it.code} — {it.title}</span>
                              <div className="text-[11px] text-slate-500 mt-0.5">
                                Liên kết TT 02/2025: <strong>Bắt buộc</strong> ({(it.requiredNls || []).join(", ")}) • <strong>Khuyến khích</strong> ({(it.recommendedNls || []).join(", ")})
                              </div>
                            </div>
                            <button onClick={() => setSelectedAtmItems(selectedAtmItems.filter((_, i) => i !== idx))} className="text-slate-400 hover:text-red-600 font-bold">✕</button>
                          </div>
                        ))}
                      </div>
                    )}
                    <div className="text-[11px] text-slate-500">Phân phối chương trình (nếu có) tải ở khung Năng lực số phía trên — dùng chung.</div>
                  </div>
                )}

                {/* 8. Config Tư tưởng HCM */}
                {opts.hcm && (
                  <div className="p-4 rounded-2xl border border-pink-300 bg-pink-50/40 space-y-2">
                    <div className="text-sm font-extrabold text-pink-800">Giáo dục tư tưởng, đạo đức, phong cách Hồ Chí Minh (QĐ 2995)</div>
                    <textarea value={hcmNote} onChange={e => setHcmNote(e.target.value)} rows={2} placeholder="Ghi chú nội dung học tập và làm theo tư tưởng, đạo đức, phong cách HCM (không bắt buộc)..." className="w-full p-2 bg-white border border-pink-200 rounded-xl text-xs" />
                  </div>
                )}
              </div>
            </div>

            {/* Ngôn ngữ & Điều chỉnh thêm */}
            <div className="pt-4 border-t border-slate-200 space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-4">
                <div>
                  <div className="text-xs font-bold text-slate-700 mb-1.5">Ngôn ngữ hiển thị của giáo án</div>
                  <div className="flex gap-2">
                    {["Tiếng Việt", "Tiếng Anh"].map(lang => (
                      <button
                        key={lang}
                        type="button"
                        onClick={() => setLessonLang(lang)}
                        className={`px-4 py-1.5 rounded-full text-xs font-bold border transition-all ${
                          lessonLang === lang ? "bg-purple-600 text-white border-purple-600" : "bg-white text-slate-600 border-slate-300"
                        }`}
                      >
                        {lang}
                      </button>
                    ))}
                  </div>
                </div>
                <label className="flex items-center gap-2.5 px-4 py-2.5 rounded-full bg-purple-50 border border-purple-200 cursor-pointer">
                  <input type="checkbox" checked={strongerAi} onChange={e => setStrongerAi(e.target.checked)} className="w-4 h-4 accent-purple-600" />
                  <span className="text-xs font-extrabold text-purple-900">🧠 AI mạnh hơn (Model Pro)</span>
                </label>
              </div>

              <div>
                <div className="text-xs font-bold text-slate-700 mb-1">Điều chỉnh thêm (không bắt buộc)</div>
                <textarea
                  value={extraInstructions}
                  maxLength={3000}
                  onChange={e => setExtraInstructions(e.target.value)}
                  rows={2}
                  placeholder={`Nội dung điều chỉnh (${extraInstructions.length}/3.000 ký tự) — Mô tả điều kiện thực tế của lớp hoặc mong muốn riêng`}
                  className="w-full p-3 bg-white border border-slate-300 rounded-xl text-xs outline-none focus:ring-2 focus:ring-purple-400"
                />
              </div>
            </div>

            <div className="flex flex-wrap justify-between items-center gap-3 pt-4 border-t border-slate-200">
              <button onClick={() => setStep(0)} className="px-4 py-2.5 text-slate-600 hover:text-slate-900 font-bold text-sm">← GIÁO ÁN</button>
              <div className="flex flex-wrap items-center gap-3">
                <button onClick={() => handleRunAi(true)} className="px-5 py-3.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-extrabold rounded-xl text-sm shadow-md">
                  ✨ AI TỰ CHỌN & TÍCH HỢP
                </button>
                <button onClick={() => handleRunAi(false)} className="px-7 py-3.5 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-extrabold rounded-xl text-sm shadow-md">
                  {mode === "existing" ? "⚡ TÍCH HỢP VÀO GIÁO ÁN →" : "⚡ TẠO GIÁO ÁN MỚI →"}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ==================== BƯỚC 3: KẾT QUẢ & TẢI WORD ==================== */}
        {step === 2 && (
          <div className="bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200">
              <div>
                <h2 className="text-lg font-bold text-slate-900">📥 {mode === "existing" ? "Tải Word đã tích hợp" : "Xem · Tải Word"}</h2>
                <p className="text-xs text-slate-500 mt-0.5">Kiểm tra báo cáo tích hợp, sửa nội dung nếu muốn và tải file .docx.</p>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={handleCopyReport} className="px-3.5 py-2 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-300 font-bold text-xs rounded-xl">
                  {copiedReport ? "✓ Đã sao chép!" : "📋 Sao chép báo cáo"}
                </button>
                <button onClick={() => setStep(1)} className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl">
                  🔄 Tích hợp lại
                </button>
              </div>
            </div>

            {/* Báo cáo tích hợp */}
            <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4">
              <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
                <div className="text-xs font-extrabold text-slate-800 uppercase">📊 Báo cáo tích hợp ({integrationSummary.total} vị trí)</div>
                {ppctName && <span className="text-[11px] font-semibold text-violet-700 bg-violet-100 px-2.5 py-0.5 rounded-full">PPCT: {ppctName}</span>}
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs">
                {Object.entries(CATEGORIES).map(([k, cat]) => {
                  const cnt = (integrationSummary.byCat[k] || []).length;
                  if (cnt === 0 && !opts[k]) return null;
                  return (
                    <div key={k} style={{ borderColor: cat.border }} className="bg-white p-2.5 rounded-xl border">
                      <div style={{ color: cat.textCol }} className="font-bold">{cat.label}: {cnt} vị trí</div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Toolbar */}
            <div className="flex flex-wrap items-center justify-between gap-3 py-2 border-b border-slate-100">
              {mode === "existing" ? (
                <div className="flex flex-wrap items-center gap-1.5">
                  <button onClick={() => setActiveFilter("all")} className={`px-3 py-1.5 rounded-xl text-xs font-bold ${activeFilter === "all" ? "bg-purple-600 text-white" : "bg-slate-100 text-slate-600"}`}>
                    Tất cả ({items.length})
                  </button>
                  {Object.entries(CATEGORIES).map(([k, cat]) => {
                    const cnt = items.filter(i => i.k === k).length;
                    if (cnt === 0) return null;
                    return (
                      <button key={k} onClick={() => setActiveFilter(k)} className={`px-3 py-1.5 rounded-xl text-xs font-bold ${activeFilter === k ? "bg-purple-600 text-white" : "bg-slate-100 text-slate-600"}`}>
                        {cat.label} ({cnt})
                      </button>
                    );
                  })}
                </div>
              ) : (
                <div className="text-sm font-extrabold text-slate-800">{newLessonData?.title}</div>
              )}

              <div className="flex items-center gap-3">
                {mode === "existing" && (
                  <>
                    <button onClick={() => setItems(prev => prev.map(it => ({ ...it, on: true })))} className="text-xs font-bold text-purple-700 hover:underline">Chọn hết</button>
                    <button onClick={() => setItems(prev => prev.map(it => ({ ...it, on: false })))} className="text-xs font-bold text-slate-500 hover:underline">Bỏ hết</button>
                  </>
                )}
                <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-700">
                  <input type="checkbox" checked={colorEnabled} onChange={e => setColorEnabled(e.target.checked)} className="w-4 h-4 accent-purple-600 rounded" />
                  <span>Tô màu phần tích hợp</span>
                </label>
              </div>
            </div>

            {mode === "existing" && (
              <div className="space-y-3">
                {filteredItems.map(item => {
                  const cat = CATEGORIES[item.k] || CATEGORIES.nls;
                  return (
                    <div key={item.id} className={`p-4 rounded-2xl border ${item.on ? "bg-white border-slate-300" : "bg-slate-50/60 border-slate-200 opacity-60"}`}>
                      <div className="flex items-center gap-2 mb-2">
                        <input type="checkbox" checked={item.on} onChange={e => setItems(prev => prev.map(it => it.id === item.id ? { ...it, on: e.target.checked } : it))} className="w-4 h-4 accent-purple-600 rounded" />
                        <span style={{ backgroundColor: cat.bg, color: cat.textCol, borderColor: cat.border }} className="px-2.5 py-0.5 rounded-lg border text-xs font-extrabold uppercase">{cat.label}</span>
                        <span className="text-xs text-slate-400 font-medium">Vị trí: {item.a} • {item.w === "end" ? "nối cuối đoạn" : "dòng mới bên dưới"}</span>
                      </div>
                      <textarea
                        value={item.t}
                        onChange={e => setItems(prev => prev.map(it => it.id === item.id ? { ...it, t: e.target.value } : it))}
                        rows={2}
                        style={{ color: colorEnabled ? cat.textCol : "#1e293b" }}
                        className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium outline-none focus:bg-white focus:ring-2 focus:ring-purple-400"
                      />
                    </div>
                  );
                })}
              </div>
            )}

            {mode === "new" && newLessonData && (
              <div className="space-y-4">
                {newLessonData.sections.map((sec, sIdx) => (
                  <div key={sec.id} className="border border-slate-200 rounded-2xl p-4 bg-slate-50/40">
                    <h3 className="font-extrabold text-sm text-slate-900 mb-3">{sec.heading}</h3>
                    <div className="space-y-2">
                      {sec.blocks.map((blk, bIdx) => {
                        const cat = blk.type !== "normal" ? CATEGORIES[blk.type] : null;
                        return (
                          <div key={blk.id} className="flex items-start gap-2">
                            {cat && (
                              <span style={{ backgroundColor: cat.bg, color: cat.textCol, borderColor: cat.border }} className="px-2 py-0.5 rounded border text-[10px] font-extrabold uppercase shrink-0 mt-1">{cat.label}</span>
                            )}
                            <textarea
                              value={blk.text}
                              rows={Math.max(1, Math.ceil(blk.text.length / 100))}
                              onChange={e => {
                                const val = e.target.value;
                                setNewLessonData(prev => ({
                                  ...prev,
                                  sections: prev.sections.map((s, si) => si !== sIdx ? s : {
                                    ...s, blocks: s.blocks.map((b, bi) => bi !== bIdx ? b : { ...b, text: val })
                                  })
                                }));
                              }}
                              style={{ color: (cat && colorEnabled) ? cat.textCol : "#1e293b" }}
                              className="w-full p-2 bg-white border border-slate-200 rounded-xl text-xs font-medium outline-none focus:ring-2 focus:ring-purple-400"
                            />
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
            )}

            <div className="sticky bottom-4 bg-slate-900 text-white rounded-2xl p-4 shadow-2xl flex flex-col sm:flex-row items-center justify-between gap-4 z-40">
              <div>
                <div className="text-sm font-bold">📥 {mode === "existing" ? `Đã duyệt ${activeCount} / ${items.length} đoạn chèn` : `Giáo án ${subject} Lớp ${grade} (${periods} tiết)`}</div>
                <div className="text-xs text-slate-400 mt-0.5">Xuất file Word (.docx) chuẩn định dạng, giữ nguyên màu phân biệt.</div>
              </div>
              <div className="flex items-center gap-3 w-full sm:w-auto">
                <button onClick={() => { setStep(0); setDocInfo(null); setItems([]); setNewLessonData(null); }} className="px-4 py-3 bg-white/10 hover:bg-white/20 text-white font-bold text-xs rounded-xl">
                  🔄 {mode === "existing" ? "Giáo án khác" : "Soạn bài khác"}
                </button>
                <button onClick={handleDownloadWord} disabled={downloading || (mode === "existing" && activeCount === 0)} className="flex-1 sm:flex-none px-6 py-3 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-700 text-white font-black text-sm rounded-xl shadow-lg">
                  {downloading ? "Đang tạo file..." : "📥 Tải Word (.docx)"}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {loading && (
        <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl text-center">
            <div className="w-16 h-16 rounded-2xl bg-purple-100 text-purple-600 flex items-center justify-center text-3xl mx-auto mb-4 animate-bounce">✨</div>
            <h3 className="text-lg font-black text-slate-900 mb-1">AI Đang Xử Lý Tích Hợp</h3>
            <p className="text-xs text-slate-500 mb-6 font-medium">{progressStatus}</p>
            <div className="w-full bg-slate-100 h-3 rounded-full overflow-hidden mb-3">
              <div className="bg-gradient-to-r from-purple-600 to-indigo-600 h-full rounded-full transition-all duration-300" style={{ width: `${progress}%` }} />
            </div>
            <div className="flex justify-between text-xs text-slate-400 font-bold"><span>Đang xử lý...</span><span>{progress}%</span></div>
          </div>
        </div>
      )}
    </div>
  );
}

const rootEl = document.getElementById("root");
if (rootEl) {
  const root = ReactDOM.createRoot(rootEl);
  root.render(React.createElement(App));
}
