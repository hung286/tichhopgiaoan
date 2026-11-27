// tich-hop-giao-an-standalone/assets/docxModel.js
var W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
var MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
var XML_NS = "http://www.w3.org/XML/1998/namespace";
var IGNORE_TAGS = /* @__PURE__ */ new Set(["txbxContent", "drawing", "pict", "object", "Fallback", "footnote", "endnote", "comment"]);
function getBodyParagraphs(doc) {
  const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
  if (!body) return [];
  const paras = body.getElementsByTagNameNS(W_NS, "p");
  const result = [];
  for (let s = 0; s < paras.length; s++) {
    let ignore = false;
    for (let r = paras[s].parentNode; r && r !== body; r = r.parentNode) {
      if (IGNORE_TAGS.has(r.localName)) {
        ignore = true;
        break;
      }
    }
    if (!ignore) result.push(paras[s]);
  }
  return result;
}
function getTableCoord(pNode, tableMap) {
  let tc = null;
  for (let r = pNode.parentNode; r && r.nodeType === 1; r = r.parentNode) {
    if (r.namespaceURI === W_NS && r.localName === "tc") {
      tc = r;
      break;
    }
    if (r.localName === "body") break;
  }
  if (!tc) return "";
  const tr = tc.parentNode;
  const tbl = tr?.parentNode;
  if (!tr || !tbl || tbl.localName !== "tbl") return "";
  const getIndex = (parent, tag) => {
    let idx = 0;
    for (let f = parent.parentNode.firstChild; f; f = f.nextSibling) {
      if (f.nodeType === 1 && f.localName === tag) {
        idx++;
        if (f === parent) return idx;
      }
    }
    return idx;
  };
  return `B${tableMap.get(tbl) ?? 0}.H${getIndex(tr, "tr")}.C${getIndex(tc, "tc")}`;
}
function extractParagraphText(pNode, stats) {
  let text = "";
  const walk = (node) => {
    for (let l = node.firstChild; l; l = l.nextSibling) {
      if (l.nodeType !== 1) continue;
      const tag = l.localName;
      if (l.namespaceURI === MATH_NS && (tag === "oMath" || tag === "oMathPara")) {
        stats.maths++;
        text += " [C\xF4ng th\u1EE9c] ";
        continue;
      }
      if (tag === "t" && l.namespaceURI === W_NS) {
        text += l.textContent || "";
      } else if (tag === "tab" || tag === "br" || tag === "cr") {
        text += " ";
      } else if (tag === "drawing" || tag === "pict") {
        text += " [H\xECnh] ";
        stats.images++;
      } else if (tag === "object") {
        text += " [C\xF4ng th\u1EE9c] ";
        stats.maths++;
      } else {
        if (tag === "delText" || tag === "instrText" || tag === "rPr" || tag === "pPr") continue;
        walk(l);
      }
    }
  };
  walk(pNode);
  return text.replace(/\s+/g, " ").trim();
}
function serializeXml(xmlDoc, origXml) {
  let serialized = new XMLSerializer().serializeToString(xmlDoc);
  if (!serialized.startsWith("<?xml")) {
    const declMatch = /^<\?xml[^>]*\?>/.exec(origXml);
    const decl = declMatch ? declMatch[0] : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    serialized = decl + "\r\n" + serialized;
  }
  return serialized;
}
function parseXml(xmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error("File Word b\u1ECB l\u1ED7i c\u1EA5u tr\xFAc XML, kh\xF4ng \u0111\u1ECDc \u0111\u01B0\u1EE3c.");
  }
  return doc;
}
async function readDocxFile(file) {
  const buf = await file.arrayBuffer();
  if (!window.JSZip) throw new Error("Ch\u01B0a t\u1EA3i \u0111\u01B0\u1EE3c th\u01B0 vi\u1EC7n JSZip.");
  const zip = await window.JSZip.loadAsync(buf);
  const docFile = zip.file("word/document.xml");
  if (!docFile) throw new Error("File kh\xF4ng ph\u1EA3i v\u0103n b\u1EA3n Word .docx h\u1EE3p l\u1EC7.");
  const xmlString = await docFile.async("string");
  const xmlDoc = parseXml(xmlString);
  const tables = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, "tbl"));
  const tableMap = new Map(tables.map((tbl, idx) => [tbl, idx + 1]));
  const stats = { images: 0, maths: 0 };
  const lines = [];
  const paraMap = /* @__PURE__ */ new Map();
  const bodyParas = getBodyParagraphs(xmlDoc);
  bodyParas.forEach((pNode, idx) => {
    const text = extractParagraphText(pNode, stats);
    if (!text) return;
    const pId = `p${idx + 1}`;
    const coord = getTableCoord(pNode, tableMap);
    paraMap.set(pId, text);
    lines.push(`[${pId}${coord ? "|" + coord : ""}] ${text}`);
  });
  return {
    fileName: file.name,
    buf,
    xml: xmlString,
    text: lines.join("\n"),
    paras: paraMap,
    paraCount: bodyParas.length,
    tables: tables.length,
    images: stats.images,
    maths: stats.maths
  };
}
function guessMetadata(docInfo) {
  const firstParas = [...docInfo.paras.values()].slice(0, 40);
  let grade = null;
  for (const text of firstParas) {
    const match = /(?:lớp|khối)\s*:?\s*(1[0-2]|[1-9])(?!\d)/i.exec(text);
    if (match) {
      grade = Number(match[1]);
      break;
    }
  }
  if (!grade) {
    const fnMatch = /(?:lớp|lop|khối|khoi|toán|toan|văn|van|tiếng việt|tieng viet|anh|lý|ly|hoá|hóa|hoa|sinh|sử|su|địa|dia|tin|khtn|gdcd|công nghệ|cong nghe)[\s_-]*(1[0-2]|[1-9])(?!\d)/i.exec(docInfo.fileName);
    if (fnMatch) grade = Number(fnMatch[1]);
  }
  const titlePatterns = [
    /^(?:bài|tiết)\s*\d+\s*[:.\-–]/i,
    /^chủ đề\s*[\dIVX]+\s*[:.\-–]/i,
    /^(?:tên bài|bài học)\s*:/i
  ];
  let foundTitle = "";
  for (const pat of titlePatterns) {
    const found = firstParas.find((p) => pat.test(p) && p.length < 200);
    if (found) {
      foundTitle = found.replace(/^(?:tên bài|bài học)\s*:\s*/i, "").slice(0, 200).trim();
      break;
    }
  }
  return { grade: grade || 6, title: foundTitle };
}
var CATEGORIES = {
  nls: { key: "nls", label: "N\u0103ng l\u1EF1c s\u1ED1", color: "7C3AED", bg: "#f5f3ff", border: "#c4b5fd", textCol: "#5b21b6" },
  ai: { key: "ai", label: "N\u0103ng l\u1EF1c AI", color: "DC2626", bg: "#fef2f2", border: "#fca5a5", textCol: "#991b1b" },
  kt: { key: "kt", label: "HS khuy\u1EBFt t\u1EADt", color: "15803D", bg: "#f0fdf4", border: "#86efac", textCol: "#166534" },
  qp: { key: "qp", label: "GDQPAN", color: "1D4ED8", bg: "#eef2ff", border: "#a5b4fc", textCol: "#1e40af", prefix: "[T\xEDch h\u1EE3p GDQPAN]: " },
  bd: { key: "bd", label: "B\u0110KH - PCTT", color: "C2410C", bg: "#fff7ed", border: "#fdba74", textCol: "#9a3412", prefix: "[T\xEDch h\u1EE3p B\u0110KH-PCTT]: " },
  ltcm: { key: "ltcm", label: "GD LTCM, \u0110\u0110LS", color: "CA8A04", bg: "#fefce8", border: "#fde047", textCol: "#854d0e", prefix: "[T\xEDch h\u1EE3p GD LTCM, \u0110\u0110LS]: " },
  atm: { key: "atm", label: "An to\xE0n m\u1EA1ng", color: "0D9488", bg: "#f0fdfa", border: "#5eead4", textCol: "#115e59" },
  hcm: { key: "hcm", label: "T\u01B0 t\u01B0\u1EDFng HCM", color: "DB2777", bg: "#fdf2f8", border: "#f9a8d4", textCol: "#9d174d", prefix: "[T\xEDch h\u1EE3p GD t\u01B0 t\u01B0\u1EDFng, \u0111\u1EA1o \u0111\u1EE9c, phong c\xE1ch HCM]: " }
};
function createEl(doc, tag, attrs = {}) {
  const el = doc.createElementNS(W_NS, "w:" + tag);
  for (const [k, v] of Object.entries(attrs)) {
    el.setAttributeNS(W_NS, "w:" + k, v);
  }
  return el;
}
function getRunFormatting(pNode) {
  if (!pNode) return [];
  const runs = Array.from(pNode.getElementsByTagNameNS(W_NS, "r")).filter((r) => r.getElementsByTagNameNS(W_NS, "t").length);
  const rPr = runs.length ? Array.from(runs[runs.length - 1].children).find((c) => c.localName === "rPr") : null;
  return rPr ? Array.from(rPr.children).filter((c) => ["rFonts", "sz", "szCs", "lang"].includes(c.localName)) : [];
}
function createTextRun(doc, text, style, templateRuns) {
  const r = createEl(doc, "r");
  const rPr = createEl(doc, "rPr");
  for (const t2 of templateRuns) {
    if (t2.localName === "rFonts" || t2.localName === "sz" || t2.localName === "szCs" || t2.localName === "lang") {
      rPr.appendChild(t2.cloneNode(true));
    }
  }
  if (style.bold) {
    rPr.appendChild(createEl(doc, "b"));
    rPr.appendChild(createEl(doc, "bCs"));
  }
  if (style.italic) {
    rPr.appendChild(createEl(doc, "i"));
    rPr.appendChild(createEl(doc, "iCs"));
  }
  if (style.underline) {
    rPr.appendChild(createEl(doc, "u", { val: "single" }));
  }
  if (style.color) {
    rPr.appendChild(createEl(doc, "color", { val: style.color }));
  }
  if (rPr.childNodes.length) r.appendChild(rPr);
  const t = createEl(doc, "t");
  t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.textContent = text;
  r.appendChild(t);
  return r;
}
function parseBold(text) {
  const tokens = [];
  const regex = /\*\*(.+?)\*\*/g;
  let lastIndex = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      tokens.push({ text: text.slice(lastIndex, match.index), bold: false });
    }
    tokens.push({ text: match[1], bold: true });
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < text.length) {
    tokens.push({ text: text.slice(lastIndex), bold: false });
  }
  return tokens.filter((t) => t.text.length > 0);
}
function buildRunsForItem(doc, item, colorEnabled, templateRuns) {
  const cat = CATEGORIES[item.k] || CATEGORIES.nls;
  const color = colorEnabled ? cat.color : void 0;
  const italic = item.k === "kt";
  const runs = [];
  let text = item.t.trim();
  if (cat.prefix && !text.startsWith(cat.prefix)) {
    runs.push(createTextRun(doc, cat.prefix, { bold: true, color, italic }, templateRuns));
  }
  const parts = parseBold(text);
  for (const p of parts) {
    runs.push(createTextRun(doc, p.text, { bold: p.bold, italic, color }, templateRuns));
  }
  return runs;
}
async function generatePatchedDocx(docInfo, items, colorEnabled = true) {
  if (!window.JSZip) throw new Error("Th\u01B0 vi\u1EC7n JSZip ch\u01B0a s\u1EB5n s\xE0ng.");
  const xmlDoc = parseXml(docInfo.xml);
  const bodyParas = getBodyParagraphs(xmlDoc);
  const findPara = (pId) => {
    const num = Number(pId.replace(/^p/i, "")) - 1;
    return bodyParas[num] || null;
  };
  const insertedMap = /* @__PURE__ */ new Map();
  for (const item of items) {
    if (item.on === false || !item.t.trim()) continue;
    const targetPara = findPara(item.a);
    if (!targetPara) continue;
    const templateRuns = getRunFormatting(targetPara);
    const newRuns = buildRunsForItem(xmlDoc, item, colorEnabled, templateRuns);
    if (item.w === "end") {
      newRuns.unshift(createTextRun(xmlDoc, " ", {}, templateRuns));
      for (const r of newRuns) targetPara.appendChild(r);
    } else {
      const newP = createEl(xmlDoc, "p");
      const pPr = targetPara.getElementsByTagNameNS(W_NS, "pPr")[0];
      if (pPr) {
        const clonedPr = pPr.cloneNode(true);
        for (const child of Array.from(clonedPr.children)) {
          if (["pageBreakBefore", "pPrChange"].includes(child.localName)) {
            clonedPr.removeChild(child);
          }
        }
        newP.appendChild(clonedPr);
      }
      for (const r of newRuns) newP.appendChild(r);
      const anchorNode = insertedMap.get(item.a) || targetPara;
      anchorNode.parentNode.insertBefore(newP, anchorNode.nextSibling);
      insertedMap.set(item.a, newP);
    }
  }
  const zip = await window.JSZip.loadAsync(docInfo.buf);
  const newXml = serializeXml(xmlDoc, docInfo.xml);
  zip.file("word/document.xml", newXml, { createFolders: false });
  return await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}
async function readXlsxFile(file) {
  if (!window.JSZip) throw new Error("Ch\u01B0a t\u1EA3i \u0111\u01B0\u1EE3c th\u01B0 vi\u1EC7n JSZip.");
  const buf = await file.arrayBuffer();
  const zip = await window.JSZip.loadAsync(buf);
  const sharedStrings = [];
  const ssFile = zip.file("xl/sharedStrings.xml");
  if (ssFile) {
    const ssXml = await ssFile.async("string");
    const ssDoc = parseXml(ssXml);
    const siList = ssDoc.getElementsByTagName("si");
    for (let i = 0; i < siList.length; i++) {
      const tNodes = siList[i].getElementsByTagName("t");
      let val = "";
      for (let j = 0; j < tNodes.length; j++) {
        val += tNodes[j].textContent || "";
      }
      sharedStrings.push(val.trim());
    }
  }
  const sheetPaths = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(p)).sort();
  const lines = [];
  for (const path of sheetPaths) {
    const sXml = await zip.file(path).async("string");
    const sDoc = parseXml(sXml);
    const rows = sDoc.getElementsByTagName("row");
    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r].getElementsByTagName("c");
      const rowVals = [];
      for (let c = 0; c < cells.length; c++) {
        const cell = cells[c];
        const tAttr = cell.getAttribute("t");
        const vNode = cell.getElementsByTagName("v")[0];
        const isNode = cell.getElementsByTagName("is")[0];
        let val = "";
        if (tAttr === "s" && vNode) {
          const idx = Number(vNode.textContent);
          val = sharedStrings[idx] ?? "";
        } else if (tAttr === "inlineStr" && isNode) {
          val = isNode.textContent || "";
        } else if (vNode) {
          val = vNode.textContent || "";
        }
        if (val.trim()) rowVals.push(val.trim());
      }
      if (rowVals.length > 0) {
        lines.push(rowVals.join(" | "));
      }
    }
  }
  return lines.join("\n");
}
async function generateNewLessonDocx(lessonData, colorEnabled = true) {
  if (!window.JSZip) throw new Error("Th\u01B0 vi\u1EC7n JSZip ch\u01B0a s\u1EB5n s\xE0ng.");
  const esc = (str) => String(str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const makeRunXml = (text, { bold = false, italic = false, color = null, size = 26 } = {}) => {
    let rPr = `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>`;
    rPr += `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`;
    if (bold) rPr += `<w:b/><w:bCs/>`;
    if (italic) rPr += `<w:i/><w:iCs/>`;
    if (color) rPr += `<w:color w:val="${color}"/>`;
    return `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
  };
  const makeParaXml = (runsXml, { align = "left", spaceAfter = 100, spaceBefore = 40 } = {}) => {
    const jc = align !== "left" ? `<w:jc w:val="${align}"/>` : "";
    return `<w:p><w:pPr><w:spacing w:before="${spaceBefore}" w:after="${spaceAfter}" w:line="276" w:lineRule="auto"/>${jc}</w:pPr>${runsXml}</w:p>`;
  };
  const bodyParts = [];
  bodyParts.push(makeParaXml(
    makeRunXml(`K\u1EBE HO\u1EA0CH B\xC0I D\u1EA0Y`, { bold: true, size: 30, color: "1E1B4B" }),
    { align: "center", spaceAfter: 60 }
  ));
  bodyParts.push(makeParaXml(
    makeRunXml(`M\xF4n: ${lessonData.subject || ""} \u2013 L\u1EDBp ${lessonData.grade || ""} | Th\u1EDDi l\u01B0\u1EE3ng: ${lessonData.periods || 1} ti\u1EBFt`, { bold: true, italic: true, size: 24 }),
    { align: "center", spaceAfter: 80 }
  ));
  bodyParts.push(makeParaXml(
    makeRunXml(`B\xC0I: ${String(lessonData.title || "").toUpperCase()}`, { bold: true, size: 28, color: "1D4ED8" }),
    { align: "center", spaceAfter: 180 }
  ));
  for (const sec of lessonData.sections || []) {
    if (sec.heading) {
      bodyParts.push(makeParaXml(
        makeRunXml(sec.heading, { bold: true, size: 26, color: "1E293B" }),
        { spaceBefore: 140, spaceAfter: 80 }
      ));
    }
    for (const blk of sec.blocks || []) {
      const cat = blk.type && blk.type !== "normal" ? CATEGORIES[blk.type] : null;
      const color = cat && colorEnabled ? cat.color : null;
      const italic = blk.type === "kt";
      let runsXml = "";
      let text = String(blk.text || "").trim();
      if (cat?.prefix && !text.startsWith(cat.prefix)) {
        runsXml += makeRunXml(cat.prefix, { bold: true, italic, color });
      }
      const parts = parseBold(text);
      for (const p of parts) {
        runsXml += makeRunXml(p.text, { bold: p.bold, italic, color });
      }
      if (runsXml) {
        bodyParts.push(makeParaXml(runsXml, { spaceAfter: 80 }));
      }
    }
  }
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyParts.join("\n    ")}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const zip = new window.JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);
  return await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

// tich-hop-giao-an-standalone/assets/frameworksData.js
var AI_CATALOG_BY_GRADE = {
  "1": [
    {
      "code": "1.A1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t s\u1ED1 c\u1EA3m x\xFAc quen thu\u1ED9c c\u1EE7a con ng\u01B0\u1EDDi (vui, bu\u1ED3n, gi\u1EADn d\u1EEF, s\u1EE3 h\xE3i, ng\u1EA1c nhi\xEAn,\u2026) qua c\xE1c t\xECnh hu\u1ED1ng g\u1EA7n g\u0169i trong cu\u1ED9c s\u1ED1ng v\xE0 b\u01B0\u1EDBc \u0111\u1EA7u nh\u1EADn ra r\u1EB1ng con ng\u01B0\u1EDDi c\xF3 c\u1EA3m x\xFAc c\xF2n AI th\xEC kh\xF4ng."
    },
    {
      "code": "1.A1.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng AI c\xF3 th\u1EC3 m\xF4 ph\u1ECFng ho\u1EB7c nh\u1EADn di\u1EC7n c\u1EA3m x\xFAc c\u1EE7a con ng\u01B0\u1EDDi, ch\u1EE9 kh\xF4ng tr\u1EA3i nghi\u1EC7m c\u1EA3m x\xFAc nh\u01B0 con ng\u01B0\u1EDDi."
    },
    {
      "code": "1.A1.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng vi\u1EC7c AI th\u1EC3 hi\u1EC7n c\u1EA3m x\xFAc (v\xED d\u1EE5: c\u01B0\u1EDDi, bu\u1ED3n, ng\u1EA1c nhi\xEAn, t\u1EE9c gi\u1EADn tr\xEAn m\xE0n h\xECnh ho\u1EB7c qua gi\u1ECDng n\xF3i) l\xE0 do con ng\u01B0\u1EDDi l\u1EADp tr\xECnh ho\u1EB7c thi\u1EBFt k\u1EBF tr\u01B0\u1EDBc."
    },
    {
      "code": "1.A1.4",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 v\xED d\u1EE5 minh h\u1ECDa (nh\u01B0 robot c\u01B0\u1EDDi khi ng\u01B0\u1EDDi d\xF9ng k\u1EC3 chuy\u1EC7n vui, ho\u1EB7c tr\u1EE3 l\xED \u1EA3o n\xF3i \u201CT\xF4i r\u1EA5t vui \u0111\u01B0\u1EE3c gi\xFAp b\u1EA1n\u201D), v\xE0 gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c r\u1EB1ng nh\u1EEFng bi\u1EC3u hi\u1EC7n n\xE0y ch\u1EC9 l\xE0 ph\u1EA3n \u1EE9ng \u0111\u01B0\u1EE3c l\u1EADp tr\xECnh s\u1EB5n."
    },
    {
      "code": "1.A2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1c c\xE1ch m\xE0 AI th\u1EC3 hi\u1EC7n c\u1EA3m x\xFAc qua h\xECnh \u1EA3nh, gi\u1ECDng n\xF3i ho\u1EB7c h\xE0nh vi (v\xED d\u1EE5: bi\u1EC3u t\u01B0\u1EE3ng khu\xF4n m\u1EB7t vui, gi\u1ECDng n\xF3i th\xE2n thi\u1EC7n, ho\u1EB7c l\u1EDDi ch\xE0o d\u1EC5 th\u01B0\u01A1ng)."
    },
    {
      "code": "1.A2.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng nh\u1EEFng bi\u1EC3u hi\u1EC7n n\xE0y gi\xFAp AI giao ti\u1EBFp t\u1EF1 nhi\xEAn h\u01A1n v\u1EDBi con ng\u01B0\u1EDDi, khi\u1EBFn ng\u01B0\u1EDDi d\xF9ng c\u1EA3m th\u1EA5y tho\u1EA3i m\xE1i, g\u1EA7n g\u0169i v\xE0 d\u1EC5 h\u1EE3p t\xE1c h\u01A1n."
    },
    {
      "code": "1.A2.2",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt v\xE0 k\u1EC3 t\xEAn \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 s\u1EA3n ph\u1EA9m ho\u1EB7c thi\u1EBFt b\u1ECB c\xF3 s\u1EED d\u1EE5ng tr\xED tu\u1EC7 nh\xE2n t\u1EA1o (v\xED d\u1EE5: loa th\xF4ng minh, \u0111i\u1EC7n tho\u1EA1i c\xF3 tr\u1EE3 l\xED \u1EA3o, robot h\xFAt b\u1EE5i, xe t\u1EF1 l\xE1i, camera nh\u1EADn di\u1EC7n khu\xF4n m\u1EB7t, \u1EE9ng d\u1EE5ng d\u1ECBch ng\xF4n ng\u1EEF\u2026)."
    },
    {
      "code": "1.A2.3",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c c\xF4ng d\u1EE5ng ch\xEDnh c\u1EE7a t\u1EEBng s\u1EA3n ph\u1EA9m v\xE0 c\xE1ch AI gi\xFAp s\u1EA3n ph\u1EA9m ho\u1EA1t \u0111\u1ED9ng th\xF4ng minh h\u01A1n (v\xED d\u1EE5: loa th\xF4ng minh hi\u1EC3u l\u1EC7nh b\u1EB1ng gi\u1ECDng n\xF3i; robot h\xFAt b\u1EE5i t\u1EF1 t\xECm \u0111\u01B0\u1EDDng v\xE0 tr\xE1nh ch\u01B0\u1EDBng ng\u1EA1i v\u1EADt). B."
    },
    {
      "code": "1.B1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 h\xE0nh vi s\u1EED d\u1EE5ng AI c\xF3 th\u1EC3 g\xE2y h\u1EA1i cho ng\u01B0\u1EDDi kh\xE1c."
    },
    {
      "code": "1.B3.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng kh\xF4ng \u0111\u01B0\u1EE3c ph\xE9p s\u1EED d\u1EE5ng AI v\u1EDBi m\u1EE5c \u0111\xEDch l\xE0m h\u1EA1i ng\u01B0\u1EDDi kh\xE1c."
    },
    {
      "code": "1.B3.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 vi\u1EC7c con ng\u01B0\u1EDDi s\u1EED d\u1EE5ng AI \u0111\xFAng c\xE1ch, v\xEC m\u1EE5c \u0111\xEDch t\u1ED1t \u0111\u1EB9p. C."
    },
    {
      "code": "1.C1.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c AI trong m\u1ED9t s\u1ED1 v\xED d\u1EE5 c\u1EE5 th\u1EC3 v\xE0 \u0111\u01A1n gi\u1EA3n v\u1EC1 AI."
    },
    {
      "code": "1.C1.2",
      "mr": false,
      "desc": "Nh\u1EADn di\u1EC7n \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 AI quen thu\u1ED9c v\xE0 ph\u1ED5 bi\u1EBFn tr\xEAn \u0111i\u1EC7n tho\u1EA1i ho\u1EB7c m\xE1y t\xEDnh b\u1EA3ng (v\xED d\u1EE5: tr\u1EE3 l\xED \u1EA3o, nh\u1EADn di\u1EC7n khu\xF4n m\u1EB7t, \u2026)."
    },
    {
      "code": "1.C1.3",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c c\xE1c thi\u1EBFt b\u1ECB th\xF4ng minh c\xF3 c\xE1c b\u1ED9 ph\u1EADn gi\u1ED1ng v\u1EDBi b\u1ED9 ph\u1EADn c\u1EE7a con ng\u01B0\u1EDDi (v\xED d\u1EE5: camera l\xE0 \u201Cm\u1EAFt\u201D v\xE0 micro l\xE0 \u201Ctai\u201D c\u1EE7a c\xE1c thi\u1EBFt b\u1ECB \u0111\xF3)."
    },
    {
      "code": "1.C1.4",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c AI c\xF3 kh\u1EA3 n\u0103ng hi\u1EC3u c\xE1c m\u1EC7nh l\u1EC7nh \u0111\u01A1n gi\u1EA3n c\u1EE7a con ng\u01B0\u1EDDi, tr\xF2 chuy\u1EC7n theo k\u1ECBch b\u1EA3n \u0111\u1ECBnh s\u1EB5n."
    },
    {
      "code": "1.C1.MR1",
      "mr": true,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c AI c\xF3 kh\u1EA3 n\u0103ng x\u1EED l\xED h\xECnh \u1EA3nh nh\u1EADn \u0111\u01B0\u1EE3c \u0111\u1EC3 nh\u1EADn di\u1EC7n c\xE1c \u0111\u1ED3 v\u1EADt."
    },
    {
      "code": "1.C1.MR2",
      "mr": true,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c AI c\xF3 kh\u1EA3 n\u0103ng x\u1EED l\xED \xE2m thanh \u0111\u1EC3 ph\xE2n bi\u1EC7t c\xE1c lo\u1EA1i \xE2m thanh kh\xE1c nhau. D."
    },
    {
      "code": "1.D1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t t\xECnh hu\u1ED1ng m\xE0 AI \u201Ch\u1ECDc\u201D t\u1EEB h\xECnh \u1EA3nh ho\u1EB7c th\xF4ng tin do con ng\u01B0\u1EDDi cung c\u1EA5p (v\xED d\u1EE5: AI h\u1ECDc nh\u1EADn bi\u1EBFt con m\xE8o, qu\u1EA3 t\xE1o\u2026)."
    },
    {
      "code": "1.D1.MR1",
      "mr": true,
      "desc": "Bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng \u0111\u1EC3 AI tr\u1EA3 l\u1EDDi \u0111\xFAng c\u1EA7n c\xF3 nhi\u1EC1u v\xED d\u1EE5 \u0111\xFAng v\xE0 kh\xE1c nhau \u0111\u1EC3 AI h\u1ECDc."
    },
    {
      "code": "1.D2.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng c\xF3 lo\u1EA1i m\xE1y th\xF4ng minh ch\u1EC9 l\xE0m \u0111\u01B0\u1EE3c m\u1ED9t vi\u1EC7c (nh\u01B0 nh\u1EADn bi\u1EBFt h\xECnh \u1EA3nh), c\xF3 lo\u1EA1i c\xF3 th\u1EC3 l\xE0m nhi\u1EC1u vi\u1EC7c kh\xE1c nhau (nh\u01B0 l\u1EAFng nghe - tr\u1EA3 l\u1EDDi, th\u1EF1c hi\u1EC7n theo l\u1EC7nh\u2026)."
    },
    {
      "code": "1.D2.2",
      "mr": false,
      "desc": "X\u1EBFp \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 m\xE1y th\xF4ng minh v\xE0o hai nh\xF3m: nh\xF3m ch\u1EC9 l\xE0m \u0111\u01B0\u1EE3c m\u1ED9t vi\u1EC7c v\xE0 nh\xF3m l\xE0m \u0111\u01B0\u1EE3c nhi\u1EC1u vi\u1EC7c kh\xE1c nhau."
    }
  ],
  "2": [
    {
      "code": "2.A1.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng trong cu\u1ED9c s\u1ED1ng m\xE0 AI c\xF3 th\u1EC3 h\u1ED7 tr\u1EE3 con ng\u01B0\u1EDDi hi\u1EC7u qu\u1EA3, ch\u1EB3ng h\u1EA1n nh\u01B0: s\u1EED d\u1EE5ng AI \u0111\u1EC3 d\u1ECBch ng\xF4n ng\u1EEF, t\xECm ki\u1EBFm th\xF4ng tin nhanh, ph\xE1t hi\u1EC7n l\u1ED7i ch\xEDnh t\u1EA3, ho\u1EB7c \u0111i\u1EC1u khi\u1EC3n thi\u1EBFt b\u1ECB th\xF4ng minh trong nh\xE0, \u2026"
    },
    {
      "code": "2.A1.2",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng trong cu\u1ED9c s\u1ED1ng m\xE0 kh\xF4ng n\xEAn ho\u1EB7c c\u1EA7n th\u1EADn tr\u1ECDng khi s\u1EED d\u1EE5ng AI, v\xED d\u1EE5: khi AI c\xF3 th\u1EC3 l\xE0m l\u1ED9 th\xF4ng tin c\xE1 nh\xE2n (nh\u01B0 khu\xF4n m\u1EB7t, \u0111\u1ECBa ch\u1EC9, d\u1EEF li\u1EC7u gi\u1ECDng n\xF3i); khi AI thay th\u1EBF ho\xE0n to\xE0n con ng\u01B0\u1EDDi, d\u1EABn \u0111\u1EBFn s\u1EF1 thi\u1EBFu tr\xE1ch nhi\u1EC7m."
    },
    {
      "code": "2.A1.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 c\u1EE5 th\u1EC3 v\u1EC1 t\xECnh hu\u1ED1ng c\u1EA7n con ng\u01B0\u1EDDi gi\xE1m s\xE1t AI (v\xED d\u1EE5: ng\u01B0\u1EDDi l\xE1i xe v\u1EABn c\u1EA7n quan s\xE1t khi xe t\u1EF1 l\xE1i ho\u1EA1t \u0111\u1ED9ng; b\xE1c s\u0129 c\u1EA7n ki\u1EC3m tra l\u1EA1i k\u1EBFt qu\u1EA3 do AI g\u1EE3i \xFD)."
    },
    {
      "code": "2.A1.4",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 c\u1EE5 th\u1EC3 v\u1EC1 th\xE1i \u0111\u1ED9 \u0111\xFAng \u0111\u1EAFn khi s\u1EED d\u1EE5ng AI: bi\u1EBFt tin t\u01B0\u1EDFng v\xE0o c\xF4ng ngh\u1EC7 \u1EDF m\u1EE9c h\u1EE3p l\xFD, nh\u01B0ng v\u1EABn c\xF3 tr\xE1ch nhi\u1EC7m v\xE0 s\u1EF1 ki\u1EC3m so\xE1t c\u1EE7a con ng\u01B0\u1EDDi."
    },
    {
      "code": "2.A2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 thi\u1EBFt b\u1ECB ho\u1EB7c \u1EE9ng d\u1EE5ng trong gia \u0111\xECnh c\xF3 s\u1EED d\u1EE5ng AI, ch\u1EB3ng h\u1EA1n nh\u01B0 loa th\xF4ng minh, robot h\xFAt b\u1EE5i, m\xE1y \u0111i\u1EC1u ho\xE0 t\u1EF1 \u0111\u1ED9ng \u0111i\u1EC1u ch\u1EC9nh nhi\u1EC7t \u0111\u1ED9, camera nh\u1EADn di\u1EC7n ng\u01B0\u1EDDi trong nh\xE0, ho\u1EB7c tivi th\xF4ng minh g\u1EE3i \xFD ch\u01B0\u01A1ng tr\xECnh y\xEAu th\xEDch."
    },
    {
      "code": "2.A2.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh s\u1EED d\u1EE5ng m\u1ED9t s\u1ED1 thi\u1EBFt b\u1ECB ho\u1EB7c \u1EE9ng d\u1EE5ng c\xF3 s\u1EED d\u1EE5ng AI v\u1EDBi s\u1EF1 gi\xE1m s\xE1t, h\u1ED7 tr\u1EE3 c\u1EE7a gi\xE1o vi\xEAn."
    },
    {
      "code": "2.A2.MR2",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng kh\xF4ng \u0111\u01B0\u1EE3c s\u1EED d\u1EE5ng thi\u1EBFt b\u1ECB ho\u1EB7c \u1EE9ng d\u1EE5ng c\xF3 s\u1EED d\u1EE5ng AI m\xE0 kh\xF4ng c\xF3 s\u1EF1 gi\xE1m s\xE1t c\u1EE7a ng\u01B0\u1EDDi l\u1EDBn."
    },
    {
      "code": "2.A2.2",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt v\xE0 k\u1EC3 t\xEAn \u0111\u01B0\u1EE3c c\xE1c th\xE0nh vi\xEAn trong gia \u0111\xECnh m\xE0 AI c\xF3 th\u1EC3 h\u1ED7 tr\u1EE3."
    },
    {
      "code": "2.A2.MR3",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng m\u1EE5c ti\xEAu c\u1EE7a AI trong b\u1ED1i c\u1EA3nh gia \u0111\xECnh l\xE0 gi\xFAp \xEDch cho cu\u1ED9c s\u1ED1ng c\u1EE7a c\xE1c th\xE0nh vi\xEAn ti\u1EC7n nghi h\u01A1n, an to\xE0n h\u01A1n v\xE0 c\xF3 nhi\u1EC1u th\u1EDDi gian cho nhau h\u01A1n."
    },
    {
      "code": "2.A3.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng m\u1ED7i l\u1EA7n con ng\u01B0\u1EDDi n\xF3i chuy\u1EC7n, \u0111\u1EB7t c\xE2u h\u1ECFi, t\xECm ki\u1EBFm, ho\u1EB7c \u0111\u01B0a ra l\u1EF1a ch\u1ECDn v\u1EDBi c\xE1c h\u1EC7 th\u1ED1ng AI (nh\u01B0 tr\u1EE3 l\xED \u1EA3o, chatbot, \u1EE9ng d\u1EE5ng g\u1EE3i \xFD video ho\u1EB7c nh\u1EA1c) th\xEC AI c\xF3 th\u1EC3 ghi nh\u1EADn d\u1EEF li\u1EC7u \u0111\u1EC3 h\u1ECDc h\u1ECFi c\xE1ch con ng\u01B0\u1EDDi suy ngh\u0129 v\xE0 ph\u1EA3n h\u1ED3i."
    },
    {
      "code": "2.A3.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng con ng\u01B0\u1EDDi c\xF3 vai tr\xF2 quan tr\u1ECDng trong vi\u1EC7c \u201Cd\u1EA1y\u201D AI th\xF4ng qua c\xE1ch t\u01B0\u01A1ng t\xE1c: khi con ng\u01B0\u1EDDi \u0111\u01B0a ra ph\u1EA3n h\u1ED3i \u0111\xFAng, l\u1ECBch s\u1EF1, v\xE0 c\xF3 tr\xE1ch nhi\u1EC7m, AI c\xF3 th\u1EC3 h\u1ECDc \u0111\u01B0\u1EE3c c\xE1c m\u1EABu th\xF4ng tin t\xEDch c\u1EF1c; ng\u01B0\u1EE3c l\u1EA1i, n\u1EBFu d\u1EEF li\u1EC7u sai l\u1EC7ch ho\u1EB7c thi\u1EBFu ch\xEDnh x\xE1c, AI c\xF3 th\u1EC3 h\u1ECDc sai. B."
    },
    {
      "code": "2.B1.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng AI \u0111\xF4i khi c\xF3 th\u1EC3 thi\xEAn ki\u1EBFn, t\u1EE9c l\xE0 \u0111\u1ED1i x\u1EED kh\xF4ng c\xF4ng b\u1EB1ng v\u1EDBi m\u1ED9t s\u1ED1 nh\xF3m ng\u01B0\u1EDDi."
    },
    {
      "code": "2.B1.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 c\xE1c t\xECnh hu\u1ED1ng th\u1EC3 hi\u1EC7n s\u1EF1 thi\xEAn ki\u1EBFn c\u1EE7a AI."
    },
    {
      "code": "2.B1.MR2",
      "mr": true,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c r\u1EB1ng n\u1EBFu d\u1EEF li\u1EC7u ch\u01B0a \u0111a d\u1EA1ng ho\u1EB7c ch\u01B0a c\xF4ng b\u1EB1ng th\xEC AI c\xF3 th\u1EC3 h\u1ECDc theo c\xE1ch thi\xEAn v\u1ECB."
    },
    {
      "code": "2.B3.1",
      "mr": false,
      "desc": "K\u1EC3 \u0111\u01B0\u1EE3c t\xEAn m\u1ED9t s\u1ED1 th\u1EE9 l\xE0 c\u1EE7a ri\xEAng em v\xE0 m\u1ED9t s\u1ED1 th\u1EE9 l\xE0 c\u1EE7a ng\u01B0\u1EDDi kh\xE1c."
    },
    {
      "code": "2.B3.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 quy\u1EC1n s\u1EDF h\u1EEFu \u0111\u1ED1i v\u1EDBi s\u1EA3n ph\u1EA9m do con ng\u01B0\u1EDDi ho\u1EB7c AI t\u1EA1o ra."
    },
    {
      "code": "2.B3.MR1",
      "mr": true,
      "desc": "Ph\xE2n bi\u1EC7t \u0111\u01B0\u1EE3c th\u1EE9 c\u1EE7a m\xECnh v\u1EDBi th\u1EE9 c\u1EE7a ng\u01B0\u1EDDi kh\xE1c trong m\u1ED9t v\xE0i t\xECnh hu\u1ED1ng \u0111\u01A1n gi\u1EA3n."
    },
    {
      "code": "2.B3.MR2",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng mu\u1ED1n d\xF9ng th\u1EE9 c\u1EE7a ng\u01B0\u1EDDi kh\xE1c th\xEC ph\u1EA3i h\u1ECFi v\xE0 \u0111\u01B0\u1EE3c ng\u01B0\u1EDDi \u0111\xF3 \u0111\u1ED3ng \xFD. C."
    },
    {
      "code": "2.C1.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c \u201Cd\u1EEF li\u1EC7u\u201D l\xE0 nh\u1EEFng v\xED d\u1EE5 (h\xECnh \u1EA3nh, \xE2m thanh) m\xE0 con ng\u01B0\u1EDDi d\xF9ng \u0111\u1EC3 d\u1EA1y cho AI."
    },
    {
      "code": "2.C1.MR1",
      "mr": true,
      "desc": "So s\xE1nh \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c \u0111\u1ED9 c\u01A1 b\u1EA3n gi\u1EEFa c\xE1ch h\u1ECDc c\u1EE7a con ng\u01B0\u1EDDi v\xE0 AI."
    },
    {
      "code": "2.C3.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng AI c\xF3 th\u1EC3 ph\xE2n lo\u1EA1i."
    },
    {
      "code": "2.C3.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng AI c\xF3 th\u1EC3 ph\xE2n lo\u1EA1i sai."
    },
    {
      "code": "2.C3.MR1",
      "mr": true,
      "desc": "Quan s\xE1t \u0111\u01B0\u1EE3c thao t\xE1c ph\xE2n lo\u1EA1i \u0111\u1ED3 v\u1EADt do gi\xE1o vi\xEAn th\u1EF1c hi\u1EC7n tr\xEAn thi\u1EBFt b\u1ECB \u0111i\u1EC7n t\u1EED ho\u1EB7c m\xF4 ph\u1ECFng; t\u1EEB k\u1EBFt qu\u1EA3 quan s\xE1t \u0111\xF3, so s\xE1nh \u0111\u01B0\u1EE3c c\xE1ch AI ph\xE2n lo\u1EA1i v\u1EDBi c\xE1ch con ng\u01B0\u1EDDi ph\xE2n lo\u1EA1i. D."
    },
    {
      "code": "2.D1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 v\u1EA5n \u0111\u1EC1 \u0111\u01A1n gi\u1EA3n, g\u1EA7n g\u0169i trong \u0111\u1EDDi s\u1ED1ng c\xF3 th\u1EC3 \xE1p d\u1EE5ng AI \u0111\u1EC3 gi\u1EA3i quy\u1EBFt."
    },
    {
      "code": "2.D1.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \xFD t\u01B0\u1EDFng m\xE1y th\xF4ng minh."
    },
    {
      "code": "2.D1.MR1",
      "mr": true,
      "desc": "N\xEAu m\u1ED9t s\u1ED1 v\xED d\u1EE5 ph\xF9 h\u1EE3p \u0111\u1EC3 \u201Cd\u1EA1y\u201D AI trong m\u1ED9t t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3."
    },
    {
      "code": "2.D2.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c \u0111\u1ED9 c\u01A1 b\u1EA3n v\u1EC1 vai tr\xF2 c\u1EE7a d\u1EEF li\u1EC7u trong vi\u1EC7c \u201Cd\u1EA1y\u201D AI trong m\u1ED9t t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3."
    },
    {
      "code": "2.D2.2",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c \u0111\u1ED9 c\u01A1 b\u1EA3n r\u1EB1ng c\u1EA7n cung c\u1EA5p cho AI d\u1EEF li\u1EC7u ch\xEDnh x\xE1c, r\xF5 r\xE0ng \u0111\u1EC3 AI \u0111\u01B0a ra k\u1EBFt qu\u1EA3 \u0111\xFAng."
    },
    {
      "code": "2.D2.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh thu th\u1EADp m\u1ED9t s\u1ED1 d\u1EEF li\u1EC7u \u0111\u1EC3 \u201Cd\u1EA1y\u201D AI trong m\u1ED9t t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3."
    }
  ],
  "3": [
    {
      "code": "3.A1.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng s\u1EED d\u1EE5ng AI h\u1ED7 tr\u1EE3 h\u1ECDc sinh trong qu\xE1 tr\xECnh h\u1ECDc t\u1EADp."
    },
    {
      "code": "3.A1.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1ch s\u1EED d\u1EE5ng tr\u1EE3 l\xED h\u1ECDc t\u1EADp th\xF4ng minh, \u1EE9ng d\u1EE5ng h\u1ECDc ng\xF4n ng\u1EEF \u0111\u1EC3 h\u1ED7 tr\u1EE3 qu\xE1 tr\xECnh h\u1ECDc t\u1EADp."
    },
    {
      "code": "3.A1.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh s\u1EED d\u1EE5ng tr\u1EE3 l\xED h\u1ECDc t\u1EADp th\xF4ng minh, \u1EE9ng d\u1EE5ng h\u1ECDc ng\xF4n ng\u1EEF."
    },
    {
      "code": "3.A1.3",
      "mr": false,
      "desc": "Ph\xE2n bi\u1EC7t \u0111\u01B0\u1EE3c gi\u1EEFa vi\u1EC7c s\u1EED d\u1EE5ng AI nh\u01B0 c\xF4ng c\u1EE5 h\u1ED7 tr\u1EE3 (\u0111\xFAng \u0111\u1EAFn, ch\u1EE7 \u0111\u1ED9ng, c\xF3 ki\u1EC3m so\xE1t) v\xE0 vi\u1EC7c \u0111\u1EC3 AI l\xE0m thay to\xE0n b\u1ED9 (th\u1EE5 \u0111\u1ED9ng, thi\u1EBFu t\u1EF1 l\u1EADp)."
    },
    {
      "code": "3.A1.MR2",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c h\u1EADu qu\u1EA3 c\u1EE7a vi\u1EC7c ph\u1EE5 thu\u1ED9c qu\xE1 m\u1EE9c v\xE0o AI, nh\u01B0: gi\u1EA3m kh\u1EA3 n\u0103ng t\u01B0 duy \u0111\u1ED9c l\u1EADp v\xE0 s\xE1ng t\u1EA1o; m\u1EA5t d\u1EA7n k\u0129 n\u0103ng gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1; \u1EA3nh h\u01B0\u1EDFng \u0111\u1EBFn s\u1EF1 t\u1EF1 tin v\xE0 b\u1EA3n l\u0129nh c\xE1 nh\xE2n."
    },
    {
      "code": "3.A1.4",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng c\xF3 th\u1EC3 g\xE2y r\u1EE7i ro ho\u1EB7c h\u1EADu qu\u1EA3 kh\xF4ng mong mu\u1ED1n khi s\u1EED d\u1EE5ng AI."
    },
    {
      "code": "3.A1.5",
      "mr": false,
      "desc": "T\u1EF1 \u0111\u1EB7t v\xE0 tr\u1EA3 l\u1EDDi \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xE2u h\u1ECFi tr\u01B0\u1EDBc khi quy\u1EBFt \u0111\u1ECBnh s\u1EED d\u1EE5ng AI trong m\u1ED9t t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3, v\xED d\u1EE5: Vi\u1EC7c n\xE0y c\xF3 c\u1EA7n d\xF9ng AI kh\xF4ng? N\u1EBFu AI \u0111\u01B0a ra k\u1EBFt qu\u1EA3 sai th\xEC \u0111i\u1EC1u g\xEC c\xF3 th\u1EC3 x\u1EA3y ra? \u2026"
    },
    {
      "code": "3.A1.MR3",
      "mr": true,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c r\u1EB1ng vi\u1EC7c suy ngh\u0129 v\xE0 ki\u1EC3m tra tr\u01B0\u1EDBc khi d\xF9ng AI gi\xFAp con ng\u01B0\u1EDDi ph\xF2ng tr\xE1nh r\u1EE7i ro, b\u1EA3o v\u1EC7 an to\xE0n v\xE0 quy\u1EC1n ri\xEAng t\u01B0, \u0111\u1ED3ng th\u1EDDi \u0111\u1EA3m b\u1EA3o AI \u0111\u01B0\u1EE3c s\u1EED d\u1EE5ng v\xEC m\u1EE5c \u0111\xEDch t\u1ED1t."
    },
    {
      "code": "3.A2.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng ho\u1EB7c thi\u1EBFt b\u1ECB c\xF3 s\u1EED d\u1EE5ng AI trong tr\u01B0\u1EDDng h\u1ECDc, ch\u1EB3ng h\u1EA1n nh\u01B0: robot ho\u1EB7c b\u1EA3ng th\xF4ng minh, \u1EE9ng d\u1EE5ng ch\u1EA5m b\xE0i ho\u1EB7c nh\u1EADn d\u1EA1ng ch\u1EEF vi\u1EBFt."
    },
    {
      "code": "3.A2.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1EE5c \u0111\xEDch c\u1EE7a vi\u1EC7c s\u1EED d\u1EE5ng AI trong tr\u01B0\u1EDDng h\u1ECDc l\xE0 h\u1ED7 tr\u1EE3 gi\xE1o vi\xEAn v\xE0 h\u1ECDc sinh h\u1ECDc t\u1EADp hi\u1EC7u qu\u1EA3 h\u01A1n, gi\xFAp: c\xE1 nh\xE2n ho\xE1 vi\u1EC7c h\u1ECDc theo n\u0103ng l\u1EF1c t\u1EEBng h\u1ECDc sinh; t\u0103ng h\u1EE9ng th\xFA h\u1ECDc t\u1EADp th\xF4ng qua c\xE1c ho\u1EA1t \u0111\u1ED9ng t\u01B0\u01A1ng t\xE1c."
    },
    {
      "code": "3.A2.MR1",
      "mr": true,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c tr\u01B0\u1EDDng h\u1ECDc mong mu\u1ED1n trong t\u01B0\u01A1ng lai \u2013 n\u01A1i c\xF3 s\u1EF1 xu\u1EA5t hi\u1EC7n c\u1EE7a m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng, thi\u1EBFt b\u1ECB AI."
    },
    {
      "code": "3.A2.MR2",
      "mr": true,
      "desc": "\u0110\u1EC1 xu\u1EA5t \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 nguy\xEAn t\u1EAFc khi s\u1EED d\u1EE5ng AI trong h\u1ECDc t\u1EADp."
    },
    {
      "code": "3.A2.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1c \u0111\u1ED1i t\u01B0\u1EE3ng trong tr\u01B0\u1EDDng h\u1ECDc m\xE0 AI c\xF3 th\u1EC3 h\u1ED7 tr\u1EE3, bao g\u1ED3m: h\u1ECDc sinh, gi\xE1o vi\xEAn, \u2026"
    },
    {
      "code": "3.A2.MR3",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1EE5c ti\xEAu c\u1EE7a AI trong b\u1ED1i c\u1EA3nh tr\u01B0\u1EDDng h\u1ECDc l\xE0 h\u1ED7 tr\u1EE3 t\u1EA5t c\u1EA3 c\xE1c th\xE0nh vi\xEAn c\u1EE7a nh\xE0 tr\u01B0\u1EDDng, gi\xFAp vi\u1EC7c d\u1EA1y v\xE0 h\u1ECDc tr\u1EDF n\xEAn d\u1EC5 d\xE0ng, hi\u1EC7u qu\u1EA3 v\xE0 c\xF4ng b\u1EB1ng h\u01A1n."
    },
    {
      "code": "3.A3.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 v\xED d\u1EE5 c\u1EE5 th\u1EC3 v\u1EC1 vi\u1EC7c AI c\xF3 th\u1EC3 sai, nh\u01B0: \u1EE9ng d\u1EE5ng d\u1ECBch sai ngh\u0129a t\u1EEB ho\u1EB7c c\xE2u; AI nh\u1EADn di\u1EC7n nh\u1EA7m h\xECnh \u1EA3nh ho\u1EB7c khu\xF4n m\u1EB7t; chatbot \u0111\u01B0a ra th\xF4ng tin kh\xF4ng \u0111\xFAng s\u1EF1 th\u1EADt."
    },
    {
      "code": "3.A3.2",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng khi s\u1EED d\u1EE5ng AI, h\u1ECDc sinh c\u1EA7n ki\u1EC3m tra l\u1EA1i k\u1EBFt qu\u1EA3, so s\xE1nh v\u1EDBi nhi\u1EC1u ngu\u1ED3n th\xF4ng tin kh\xE1c nhau v\xE0 trao \u0111\u1ED5i v\u1EDBi gi\xE1o vi\xEAn ho\u1EB7c ng\u01B0\u1EDDi l\u1EDBn \u0111\u1EC3 x\xE1c minh. B."
    },
    {
      "code": "3.B2.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt v\xE0 n\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 vi\u1EC7c th\xF4ng tin ho\u1EB7c s\u1EA3n ph\u1EA9m do AI t\u1EA1o ra c\xF3 th\u1EC3 kh\xF4ng \u0111\xFAng v\u1EDBi s\u1EF1 th\u1EADt."
    },
    {
      "code": "3.B3.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng kh\xF4ng n\xEAn t\u1EA1o ra ho\u1EB7c s\u1EED d\u1EE5ng AI cho c\xE1c m\u1EE5c \u0111\xEDch x\u1EA5u nh\u01B0 l\u1EEBa \u0111\u1EA3o, b\u1EAFt n\u1EA1t hay g\xE2y h\u1EA1i cho ng\u01B0\u1EDDi kh\xE1c."
    },
    {
      "code": "3.B3.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng con ng\u01B0\u1EDDi c\u1EA7n cung c\u1EA5p d\u1EEF li\u1EC7u \u0111\xFAng v\xE0 \u0111a d\u1EA1ng \u0111\u1EC3 AI ho\u1EA1t \u0111\u1ED9ng ch\xEDnh x\xE1c, c\xF4ng b\u1EB1ng. C."
    },
    {
      "code": "3.C4.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c kh\xE1i ni\u1EC7m d\u1EEF li\u1EC7u th\xF4ng qua m\u1ED9t s\u1ED1 v\xED d\u1EE5."
    },
    {
      "code": "3.C4.MR1",
      "mr": true,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c \u0111\u1EB7c tr\u01B0ng c\u1EE7a m\u1ED9t s\u1ED1 b\u1ED9 d\u1EEF li\u1EC7u \u0111\u01A1n gi\u1EA3n."
    },
    {
      "code": "3.C5.1",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c c\u1EA5u tr\xFAc n\u1EBFu \u2026 th\xEC \u2026 trong vi\u1EC7c gi\u1EA3i quy\u1EBFt t\xECnh hu\u1ED1ng ho\u1EB7c ph\xE2n lo\u1EA1i."
    },
    {
      "code": "3.C5.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 AI \u0111\u01A1n gi\u1EA3n s\u1EED d\u1EE5ng c\u1EA5u tr\xFAc \u201Cn\u1EBFu \u2026 th\xEC \u2026\u201D."
    },
    {
      "code": "3.C5.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c \u0111\u1EB7c \u0111i\u1EC3m ch\xEDnh c\u1EE7a h\u1ECDc m\xE1y l\xE0 h\u1ECDc tr\xEAn d\u1EEF li\u1EC7u."
    },
    {
      "code": "3.C5.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng c\xF3 th\u1EC3 \xE1p d\u1EE5ng h\u1ECDc m\xE1y (v\xED d\u1EE5: ph\xE2n lo\u1EA1i rau c\u1EE7 trong n\xF4ng tr\u1EA1i th\xF4ng minh, d\u1EF1 b\xE1o m\u01B0a l\u0169, \u2026)."
    },
    {
      "code": "3.C5.4",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c b\xE0i to\xE1n h\u1ECDc m\xE1y ph\xE2n lo\u1EA1i v\xE0 b\xE0i to\xE1n h\u1ECDc m\xE1y d\u1EF1 \u0111o\xE1n th\xF4ng qua t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3."
    },
    {
      "code": "3.C5.MR2",
      "mr": true,
      "desc": "Minh h\u1ECDa \u0111\u01B0\u1EE3c b\xE0i to\xE1n h\u1ECDc m\xE1y ph\xE2n lo\u1EA1i v\xE0 b\xE0i to\xE1n h\u1ECDc m\xE1y d\u1EF1 \u0111o\xE1n b\u1EB1ng c\xE1c c\xF4ng c\u1EE5 m\xE3 ngu\u1ED3n m\u1EDF ho\u1EB7c mi\u1EC5n ph\xED nh\u01B0 Teachable Machine, ML for Kids t\xEDch h\u1EE3p trong Scratch, ho\u1EB7c m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 ph\xF9 h\u1EE3p kh\xE1c. D."
    },
    {
      "code": "3.D1.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c qu\xE1 tr\xECnh \u0111\u01A1n gi\u1EA3n \u0111\u1EC3 hu\u1EA5n luy\u1EC7n AI g\u1ED3m c\xE1c b\u01B0\u1EDBc: thu th\u1EADp v\xED d\u1EE5 v\xE0 cho AI h\u1ECDc t\u1EEB c\xE1c v\xED d\u1EE5 \u0111\xF3."
    },
    {
      "code": "3.D2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng d\u1EEF li\u1EC7u d\xF9ng \u0111\u1EC3 d\u1EA1y AI c\xF3 th\u1EC3 b\u1ECB thi\u1EBFu, nh\u1EA7m l\u1EABn ho\u1EB7c kh\xF4ng \u0111\xFAng s\u1EF1 th\u1EADt."
    },
    {
      "code": "3.D2.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 y\xEAu c\u1EA7u c\u01A1 b\u1EA3n \u0111\u1ED1i v\u1EDBi d\u1EEF li\u1EC7u d\xF9ng \u0111\u1EC3 hu\u1EA5n luy\u1EC7n AI."
    },
    {
      "code": "3.D2.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh ki\u1EC3m tra m\u1ED9t b\u1ED9 th\u1EBB d\u1EEF li\u1EC7u m\u1EABu v\xE0 lo\u1EA1i b\u1ECF \u0111\u01B0\u1EE3c c\xE1c th\u1EBB g\xE1n nh\xE3n nh\u1EA7m, th\u1EBB kh\xF4ng \u0111\xFAng s\u1EF1 th\u1EADt."
    },
    {
      "code": "3.D2.3",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng n\u1EBFu d\u1EEF li\u1EC7u sai ho\u1EB7c kh\xF4ng t\u1ED1t th\xEC AI s\u1EBD kh\xF4ng hi\u1EC7u qu\u1EA3, th\u1EADm ch\xED ho\u1EA1t \u0111\u1ED9ng sai."
    }
  ],
  "4": [
    {
      "code": "4.A1.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 l\u0129nh v\u1EF1c ho\u1EB7c c\xF4ng vi\u1EC7c c\u1EE5 th\u1EC3 m\xE0 tr\xED tu\u1EC7 nh\xE2n t\u1EA1o (AI) c\xF3 th\u1EC3 h\u1ED7 tr\u1EE3 con ng\u01B0\u1EDDi, ch\u1EB3ng h\u1EA1n nh\u01B0 n\xF4ng nghi\u1EC7p, y t\u1EBF, giao th\xF4ng, \u2026"
    },
    {
      "code": "4.A1.2",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c r\u1EB1ng tr\xED tu\u1EC7 nh\xE2n t\u1EA1o (AI) \u0111\u01B0\u1EE3c t\u1EA1o ra \u0111\u1EC3 h\u1ED7 tr\u1EE3 con ng\u01B0\u1EDDi h\u1ECDc t\u1EADp v\xE0 l\xE0m vi\u1EC7c hi\u1EC7u qu\u1EA3 h\u01A1n nh\u01B0ng kh\xF4ng th\u1EC3 thay th\u1EBF t\u01B0 duy, c\u1EA3m x\xFAc v\xE0 s\u1EF1 s\xE1ng t\u1EA1o c\u1EE7a con ng\u01B0\u1EDDi."
    },
    {
      "code": "4.A1.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 c\u1EE5 th\u1EC3 v\u1EC1 vi\u1EC7c AI h\u1ED7 tr\u1EE3 trong h\u1ECDc t\u1EADp v\xE0 gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c v\xEC sao ng\u01B0\u1EDDi h\u1ECDc v\u1EABn c\u1EA7n t\u1EF1 suy ngh\u0129, ki\u1EC3m tra v\xE0 hi\u1EC3u b\xE0i thay v\xEC ch\u1EC9 ch\xE9p k\u1EBFt qu\u1EA3 do AI \u0111\u01B0a ra."
    },
    {
      "code": "4.A2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c r\u1EB1ng AI \u0111\u01B0\u1EE3c con ng\u01B0\u1EDDi t\u1EA1o ra \u0111\u1EC3 h\u1ED7 tr\u1EE3 c\xE1c ho\u1EA1t \u0111\u1ED9ng trong \u0111\u1EDDi s\u1ED1ng x\xE3 h\u1ED9i, gi\xFAp gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1, ti\u1EBFt ki\u1EC7m th\u1EDDi gian v\xE0 n\xE2ng cao ch\u1EA5t l\u01B0\u1EE3ng cu\u1ED9c s\u1ED1ng."
    },
    {
      "code": "4.A2.2",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1c \u0111\u1ED1i t\u01B0\u1EE3ng trong x\xE3 h\u1ED9i AI c\xF3 th\u1EC3 h\u1ED7 tr\u1EE3, bao g\u1ED3m: ng\u01B0\u1EDDi lao \u0111\u1ED9ng, ng\u01B0\u1EDDi y\u1EBFu th\u1EBF, \u2026"
    },
    {
      "code": "4.A3.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng vi\u1EC7c quy\u1EBFt \u0111\u1ECBnh c\xF3 s\u1EED d\u1EE5ng AI hay kh\xF4ng ph\u1EE5 thu\u1ED9c v\xE0o m\u1EE5c \u0111\xEDch, nhu c\u1EA7u v\xE0 s\u1EF1 an to\xE0n c\u1EE7a con ng\u01B0\u1EDDi."
    },
    {
      "code": "4.A3.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh \u0111\u01B0a quy\u1EBFt \u0111\u1ECBnh c\xF3 s\u1EED d\u1EE5ng AI hay kh\xF4ng trong m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3: N\u1EBFu AI mang l\u1EA1i l\u1EE3i \xEDch, gi\xFAp ti\u1EBFt ki\u1EC7m th\u1EDDi gian, t\u0103ng hi\u1EC7u qu\u1EA3 ho\u1EB7c \u0111\u1EA3m b\u1EA3o an to\xE0n, con ng\u01B0\u1EDDi c\xF3 th\u1EC3 l\u1EF1a ch\u1ECDn s\u1EED d\u1EE5ng. N\u1EBFu AI g\xE2y r\u1EE7i ro, x\xE2m ph\u1EA1m quy\u1EC1n ri\xEAng t\u01B0, ho\u1EB7c \u1EA3nh h\u01B0\u1EDFng \u0111\u1EBFn gi\xE1 tr\u1ECB \u0111\u1EA1o \u0111\u1EE9c, con ng\u01B0\u1EDDi c\xF3 quy\u1EC1n t\u1EEB ch\u1ED1i ho\u1EB7c d\u1EEBng s\u1EED d\u1EE5ng. B."
    },
    {
      "code": "4.B2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 lo\u1EA1i th\xF4ng tin c\xE1 nh\xE2n c\u1EA7n \u0111\u01B0\u1EE3c gi\u1EEF b\xED m\u1EADt v\xE0 kh\xF4ng n\xEAn chia s\u1EBB cho AI."
    },
    {
      "code": "4.B2.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 h\u1EADu qu\u1EA3 khi th\xF4ng tin c\xE1 nh\xE2n b\u1ECB l\u1ED9 ho\u1EB7c b\u1ECB l\u1EE3i d\u1EE5ng."
    },
    {
      "code": "4.B2.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh x\u1EED l\xED theo h\u01B0\u1EDBng d\u1EABn m\u1ED9t t\xECnh hu\u1ED1ng gi\u1EA3 \u0111\u1ECBnh \u0111\u01A1n gi\u1EA3n v\u1EC1 vi\u1EC7c b\u1ECB l\u1ED9 th\xF4ng tin c\xE1 nh\xE2n khi d\xF9ng AI. C."
    },
    {
      "code": "4.C2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1c \u1EE9ng d\u1EE5ng c\u1EE7a AI trong h\u1ECDc t\u1EADp v\xE0 \u0111\u1EDDi s\u1ED1ng, \u0111\u1EB7c bi\u1EC7t l\xE0 c\xE1c \u1EE9ng d\u1EE5ng g\u1EA7n g\u0169i v\u1EDBi b\u1ED1i c\u1EA3nh Vi\u1EC7t Nam."
    },
    {
      "code": "4.C2.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh tr\u1EA3i nghi\u1EC7m m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng AI trong h\u1ECDc t\u1EADp v\xE0 \u0111\u1EDDi s\u1ED1ng."
    },
    {
      "code": "4.C5.MR1",
      "mr": true,
      "desc": "Th\u1EF1c hi\u1EC7n l\u1EA1i \u0111\u01B0\u1EE3c theo t\u1EEBng b\u01B0\u1EDBc c\xE1c thao t\xE1c c\u01A1 b\u1EA3n v\u1EDBi c\xF4ng c\u1EE5 tr\u1EA3i nghi\u1EC7m AI tr\u1EF1c quan d\u1EF1a tr\xEAn h\u1ECDc m\xE1y c\xF3 gi\xE1m s\xE1t nh\u01B0 Teachable Machine, ML for Kids t\xEDch h\u1EE3p trong Scratch, ho\u1EB7c m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 ph\xF9 h\u1EE3p kh\xE1c theo h\u01B0\u1EDBng d\u1EABn m\u1EABu c\u1EE7a gi\xE1o vi\xEAn: (1) m\u1EDF c\xF4ng c\u1EE5 v\xE0 \u0111\u1EB7t t\xEAn cho c\xE1c nh\xF3m c\u1EA7n ph\xE2n lo\u1EA1i; (2) \u0111\u01B0a v\xE0o m\u1ED7i nh\xF3m m\u1ED9t s\u1ED1 h\xECnh \u1EA3nh m\u1EABu; (3) b\u1EA5m l\u1EC7nh cho m\xE1y h\u1ECDc t\u1EEB c\xE1c m\u1EABu \u0111\xF3; (4) \u0111\u01B0a m\u1ED9t h\xECnh \u1EA3nh m\u1EDBi v\xE0o v\xE0 quan s\xE1t k\u1EBFt qu\u1EA3 m\xE1y \u0111\u01B0a ra."
    },
    {
      "code": "4.C5.MR2",
      "mr": true,
      "desc": "L\u1EB7p l\u1EA1i \u0111\u01B0\u1EE3c \u0111\u1EA7y \u0111\u1EE7 b\u1ED1n b\u01B0\u1EDBc tr\xEAn v\u1EDBi m\u1ED9t nh\xF3m d\u1EEF li\u1EC7u t\u1EF1 ch\u1ECDn kh\xE1c v\xE0 m\xF4 t\u1EA3 \u0111\u01B0\u1EE3c thao t\xE1c \u1EDF m\u1ED7i b\u01B0\u1EDBc \u0111\xE3 l\xE0m. D."
    },
    {
      "code": "4.D1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \xFD t\u01B0\u1EDFng ban \u0111\u1EA7u v\u1EC1 c\xE1ch AI c\xF3 th\u1EC3 h\u1ECDc ho\u1EB7c gi\xFAp gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1, \u01B0u ti\xEAn c\xE1c v\u1EA5n \u0111\u1EC1 g\u1EA7n g\u0169i v\u1EDBi b\u1ED1i c\u1EA3nh Vi\u1EC7t Nam, v\xED d\u1EE5: AI h\u1ECDc t\u1EEB \u1EA3nh c\xE1c lo\u1EA1i r\xE1c \u0111\u1EC3 ph\xE2n lo\u1EA1i r\xE1c t\xE1i ch\u1EBF."
    },
    {
      "code": "4.D1.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1ch s\u1EED d\u1EE5ng AI \u0111\u1EC3 gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1 trong t\xECnh hu\u1ED1ng \u0111\u01A1n gi\u1EA3n, \u01B0u ti\xEAn c\xE1c v\u1EA5n \u0111\u1EC1 g\u1EA7n g\u0169i v\u1EDBi b\u1ED1i c\u1EA3nh Vi\u1EC7t Nam, v\xED d\u1EE5: AI h\u1ECDc t\u1EEB gi\u1ECDng n\xF3i v\xE0 ch\u1EEF vi\u1EBFt c\u1EE7a \u0111\u1ED3ng b\xE0o d\xE2n t\u1ED9c thi\u1EC3u s\u1ED1 \u0111\u1EC3 d\u1ECBch sang ti\u1EBFng ph\u1ED5 th\xF4ng, gi\xFAp h\u1ECDc sinh v\xF9ng cao h\u1ECDc t\u1EADp thu\u1EADn l\u1EE3i h\u01A1n."
    },
    {
      "code": "4.D2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng con ng\u01B0\u1EDDi c\u1EA7n li\xEAn t\u1EE5c \u0111\xE1nh gi\xE1 v\xE0 n\xE2ng c\u1EA5p s\u1EA3n ph\u1EA9m \u0111\u1EC3 h\u1EC7 th\u1ED1ng AI cho ra k\u1EBFt qu\u1EA3 t\u1ED1t h\u01A1n."
    }
  ],
  "5": [
    {
      "code": "5.A1.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng AI c\xF3 th\u1EC3 th\u1EF1c hi\u1EC7n m\u1ED9t s\u1ED1 c\xF4ng vi\u1EC7c thay con ng\u01B0\u1EDDi, \u0111\u1EB7c bi\u1EC7t l\xE0 nh\u1EEFng vi\u1EC7c l\u1EB7p \u0111i l\u1EB7p l\u1EA1i, nguy hi\u1EC3m ho\u1EB7c \u0111\xF2i h\u1ECFi t\xEDnh ch\xEDnh x\xE1c cao, nh\u01B0: robot l\u1EAFp r\xE1p s\u1EA3n ph\u1EA9m trong nh\xE0 m\xE1y; AI ki\u1EC3m tra l\u1ED7i trong s\u1EA3n xu\u1EA5t; xe t\u1EF1 l\xE1i ch\u1EDF h\xE0ng ho\xE1."
    },
    {
      "code": "5.A1.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 c\u1EE5 th\u1EC3 v\u1EC1 tr\xE1ch nhi\u1EC7m c\u1EE7a ng\u01B0\u1EDDi t\u1EA1o v\xE0 s\u1EED d\u1EE5ng AI, v\xED d\u1EE5: n\u1EBFu AI ch\u1EA9n \u0111o\xE1n nh\u1EA7m b\u1EC7nh, b\xE1c s\u0129 c\u1EA7n ki\u1EC3m tra l\u1EA1i v\xE0 ch\u1ECBu tr\xE1ch nhi\u1EC7m \u0111i\u1EC1u ch\u1EC9nh."
    },
    {
      "code": "5.A1.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c con ng\u01B0\u1EDDi ch\u1ECBu tr\xE1ch nhi\u1EC7m cu\u1ED1i c\xF9ng v\u1EC1 m\u1ECDi quy\u1EBFt \u0111\u1ECBnh ho\u1EB7c k\u1EBFt qu\u1EA3 do AI t\u1EA1o ra."
    },
    {
      "code": "5.A2.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c r\u1EB1ng m\u1EE5c \u0111\xEDch ch\xEDnh c\u1EE7a AI l\xE0 h\u1ED7 tr\u1EE3 con ng\u01B0\u1EDDi, ch\u1EE9 kh\xF4ng thay th\u1EBF vai tr\xF2, t\u01B0 duy, c\u1EA3m x\xFAc v\xE0 tr\xE1ch nhi\u1EC7m c\u1EE7a con ng\u01B0\u1EDDi trong x\xE3 h\u1ED9i."
    },
    {
      "code": "5.A2.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng d\xF9 AI l\xE0m \u0111\u01B0\u1EE3c nhi\u1EC1u vi\u1EC7c, nh\u01B0ng ch\u1EC9 con ng\u01B0\u1EDDi m\u1EDBi c\xF3 kh\u1EA3 n\u0103ng suy ngh\u0129 s\xE1ng t\u1EA1o, hi\u1EC3u c\u1EA3m x\xFAc v\xE0 \u0111\u01B0a ra quy\u1EBFt \u0111\u1ECBnh \u0111\u1EA1o \u0111\u1EE9c \u2013 v\xEC v\u1EADy, con ng\u01B0\u1EDDi ph\u1EA3i lu\xF4n ki\u1EC3m so\xE1t, \u0111\u1ECBnh h\u01B0\u1EDBng v\xE0 ch\u1ECBu tr\xE1ch nhi\u1EC7m v\u1EC1 c\xE1ch AI \u0111\u01B0\u1EE3c s\u1EED d\u1EE5ng."
    },
    {
      "code": "5.A2.2",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng AI \u0111\u01B0\u1EE3c con ng\u01B0\u1EDDi t\u1EA1o ra nh\u1EB1m ph\u1EE5c v\u1EE5 l\u1EE3i \xEDch chung c\u1EE7a x\xE3 h\u1ED9i, gi\xFAp c\u1EA3i thi\u1EC7n ch\u1EA5t l\u01B0\u1EE3ng cu\u1ED9c s\u1ED1ng, gi\u1EA3i quy\u1EBFt c\xE1c v\u1EA5n \u0111\u1EC1 ph\u1EE9c t\u1EA1p v\xE0 h\u1ED7 tr\u1EE3 con ng\u01B0\u1EDDi trong nhi\u1EC1u l\u0129nh v\u1EF1c kh\xE1c nhau."
    },
    {
      "code": "5.A2.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 c\u1EE5 th\u1EC3 v\u1EC1 vi\u1EC7c AI mang l\u1EA1i l\u1EE3i \xEDch cho c\u1ED9ng \u0111\u1ED3ng: y t\u1EBF, gi\xE1o d\u1EE5c, m\xF4i tr\u01B0\u1EDDng."
    },
    {
      "code": "5.A3.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng AI \u0111ang ng\xE0y c\xE0ng tr\u1EDF n\xEAn ph\u1ED5 bi\u1EBFn trong cu\u1ED9c s\u1ED1ng, v\xEC v\u1EADy m\u1ECDi ng\u01B0\u1EDDi \u2013 d\xF9 l\xE0 tr\u1EBB em, ng\u01B0\u1EDDi l\u1EDBn hay ng\u01B0\u1EDDi cao tu\u1ED5i \u2013 \u0111\u1EC1u c\u1EA7n hi\u1EC3u v\xE0 bi\u1EBFt c\xE1ch s\u1EED d\u1EE5ng AI \u0111\u1EC3 ph\u1EE5c v\u1EE5 cu\u1ED9c s\u1ED1ng, h\u1ECDc t\u1EADp v\xE0 c\xF4ng vi\u1EC7c m\u1ED9t c\xE1ch an to\xE0n, hi\u1EC7u qu\u1EA3."
    },
    {
      "code": "5.A3.2",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c r\u1EB1ng vi\u1EC7c s\u1EED d\u1EE5ng AI ph\u1EA3i ph\xF9 h\u1EE3p v\u1EDBi nhu c\u1EA7u, kh\u1EA3 n\u0103ng v\xE0 m\u1EE5c \u0111\xEDch c\u1EE7a t\u1EEBng ng\u01B0\u1EDDi, tr\xE1nh l\u1EA1m d\u1EE5ng ho\u1EB7c s\u1EED d\u1EE5ng kh\xF4ng an to\xE0n (v\xED d\u1EE5: chia s\u1EBB th\xF4ng tin c\xE1 nh\xE2n, tin t\u01B0\u1EDFng tuy\u1EC7t \u0111\u1ED1i v\xE0o AI)."
    },
    {
      "code": "5.A3.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c \xFD t\u01B0\u1EDFng v\u1EC1 m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng trong t\u01B0\u01A1ng lai m\xE0 AI h\u1ED7 tr\u1EE3 tr\u1EBB em, ng\u01B0\u1EDDi l\u1EDBn ho\u1EB7c ng\u01B0\u1EDDi cao tu\u1ED5i. B."
    },
    {
      "code": "5.B1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 v\xED d\u1EE5 v\u1EC1 s\u1EF1 c\xF4ng b\u1EB1ng ho\u1EB7c kh\xF4ng c\xF4ng b\u1EB1ng trong m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng \u0111\u01A1n gi\u1EA3n, trong \u0111\u1EDDi s\u1ED1ng v\xE0 khi AI ph\u1EE5c v\u1EE5 con ng\u01B0\u1EDDi."
    },
    {
      "code": "5.B1.2",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c r\u1EB1ng AI c\u1EA7n ph\u1EE5c v\u1EE5 con ng\u01B0\u1EDDi m\u1ED9t c\xE1ch c\xF4ng b\u1EB1ng, kh\xF4ng ph\xE2n bi\u1EC7t gi\u1EDBi t\xEDnh, v\xF9ng mi\u1EC1n, \u0111i\u1EC1u ki\u1EC7n kinh t\u1EBF hay ho\xE0n c\u1EA3nh kh\xE1c nhau."
    },
    {
      "code": "5.B2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xE1ch gi\xFAp AI ho\u1EA1t \u0111\u1ED9ng c\xF4ng b\u1EB1ng h\u01A1n, ch\u1EB3ng h\u1EA1n nh\u01B0 s\u1EED d\u1EE5ng d\u1EEF li\u1EC7u \u0111a d\u1EA1ng, tr\xE1nh \u0111\u1ECBnh ki\u1EBFn ho\u1EB7c ki\u1EC3m tra l\u1EA1i k\u1EBFt qu\u1EA3 do AI t\u1EA1o ra."
    },
    {
      "code": "5.B2.MR1",
      "mr": true,
      "desc": "\u0110\u1EC1 xu\u1EA5t \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 quy t\u1EAFc khi thu th\u1EADp d\u1EEF li\u1EC7u \u0111\u1EC3 c\xF3 \u0111\u01B0\u1EE3c b\u1ED9 d\u1EEF li\u1EC7u \u0111a d\u1EA1ng."
    },
    {
      "code": "5.B3.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c v\xEC sao con ng\u01B0\u1EDDi c\u1EA7n hi\u1EC3u c\xE1ch AI \u0111\u01B0a ra quy\u1EBFt \u0111\u1ECBnh, nh\u1EB1m \u0111\u1EA3m b\u1EA3o t\xEDnh minh b\u1EA1ch v\xE0 \u0111\u1ED9 tin c\u1EADy khi s\u1EED d\u1EE5ng AI. C."
    },
    {
      "code": "5.C5.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1ch s\u1EED d\u1EE5ng c\u1EA5u tr\xFAc n\u1EBFu \u2026 th\xEC \u2026 trong l\u1EADp tr\xECnh AI \u0111\u01A1n gi\u1EA3n."
    },
    {
      "code": "5.C5.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh s\u1EED d\u1EE5ng c\u1EA5u tr\xFAc n\u1EBFu \u2026 th\xEC \u2026 trong l\u1EADp tr\xECnh AI \u0111\u01A1n gi\u1EA3n."
    },
    {
      "code": "5.C5.2",
      "mr": false,
      "desc": "Th\u1EF1c hi\u1EC7n \u0111\u01B0\u1EE3c c\xE1c thao t\xE1c c\u01A1 b\u1EA3n v\u1EDBi c\xF4ng c\u1EE5 tr\u1EA3i nghi\u1EC7m AI tr\u1EF1c quan d\u1EF1a tr\xEAn h\u1ECDc m\xE1y c\xF3 gi\xE1m s\xE1t nh\u01B0 Teachable Machine, ML for Kids t\xEDch h\u1EE3p trong Scratch, ho\u1EB7c m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 ph\xF9 h\u1EE3p kh\xE1c."
    },
    {
      "code": "5.C5.MR2",
      "mr": true,
      "desc": "Hu\u1EA5n luy\u1EC7n \u0111\u01B0\u1EE3c m\u1ED9t m\xF4 h\xECnh ph\xE2n lo\u1EA1i \u0111\u01A1n gi\u1EA3n tr\xEAn c\xF4ng c\u1EE5 tr\u1EA3i nghi\u1EC7m AI tr\u1EF1c quan, v\xED d\u1EE5: ph\xE2n lo\u1EA1i l\xE1 c\xE2y kho\u1EBB \u2013 l\xE1 c\xE2y b\u1ECB s\xE2u b\u1EC7nh."
    },
    {
      "code": "5.C5.MR3",
      "mr": true,
      "desc": "Ki\u1EC3m tra \u0111\u01B0\u1EE3c k\u1EBFt qu\u1EA3 do m\xF4 h\xECnh \u0111\u01B0a ra, ch\u1EC9 ra \u0111\u01B0\u1EE3c nh\u1EEFng tr\u01B0\u1EDDng h\u1EE3p m\xF4 h\xECnh ph\xE2n lo\u1EA1i sai. D."
    },
    {
      "code": "5.D1.1",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c \u0111\u1ED9 \u0111\u01A1n gi\u1EA3n v\u1EC1 c\xE1c b\u01B0\u1EDBc c\u01A1 b\u1EA3n trong qu\xE1 tr\xECnh hu\u1EA5n luy\u1EC7n m\xF4 h\xECnh AI: x\xE1c \u0111\u1ECBnh v\u1EA5n \u0111\u1EC1, thu th\u1EADp d\u1EEF li\u1EC7u, d\u1EA1y m\xE1y h\u1ECDc, ki\u1EC3m tra v\xE0 \u0111\xE1nh gi\xE1 k\u1EBFt qu\u1EA3."
    },
    {
      "code": "5.D1.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh thu th\u1EADp d\u1EEF li\u1EC7u nh\u1EB1m hu\u1EA5n luy\u1EC7n m\xF4 h\xECnh AI gi\u1EA3i quy\u1EBFt m\u1ED9t v\u1EA5n \u0111\u1EC1 \u0111\u01A1n gi\u1EA3n."
    },
    {
      "code": "5.D2.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c b\u1EB1ng v\xED d\u1EE5 r\u1EB1ng h\u1EC7 th\u1ED1ng AI c\xF3 th\u1EC3 \u0111\u01B0\u1EE3c c\u1EA3i ti\u1EBFn v\xE0 ho\u1EA1t \u0111\u1ED9ng t\u1ED1t h\u01A1n khi d\u1EEF li\u1EC7u \u0111\u01B0\u1EE3c b\u1ED5 sung v\xE0 c\u1EADp nh\u1EADt th\u01B0\u1EDDng xuy\xEAn."
    },
    {
      "code": "5.D2.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh c\u1EA3i ti\u1EBFn h\u1EC7 th\u1ED1ng AI b\u1EB1ng c\xE1ch b\u1ED5 sung d\u1EEF li\u1EC7u."
    }
  ],
  "6": [
    {
      "code": "6.A1.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c AI l\xE0 s\u1EA3n ph\u1EA9m do con ng\u01B0\u1EDDi t\u1EA1o ra, l\u1EADp tr\xECnh v\xE0 \u0111i\u1EC1u khi\u1EC3n \u0111\u1EC3 th\u1EF1c hi\u1EC7n nh\u1EEFng nhi\u1EC7m v\u1EE5 c\u1EE5 th\u1EC3; AI kh\xF4ng t\u1EF1 sinh ra v\xE0 kh\xF4ng ho\u1EA1t \u0111\u1ED9ng \u0111\u1ED9c l\u1EADp v\u1EDBi con ng\u01B0\u1EDDi. N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 AI quen thu\u1ED9c trong \u0111\u1EDDi s\u1ED1ng h\u1EB1ng ng\xE0y v\xE0 ch\u1EC9 ra \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a con ng\u01B0\u1EDDi trong vi\u1EC7c t\u1EA1o ra ch\xFAng."
    },
    {
      "code": "6.A1.2",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a AI ch\u1EC9 l\xE0 c\xF4ng c\u1EE5 h\u1ED7 tr\u1EE3 ho\u1EA1t \u0111\u1ED9ng c\u1EE7a con ng\u01B0\u1EDDi; con ng\u01B0\u1EDDi \u0111\u01B0a ra quy\u1EBFt \u0111\u1ECBnh cu\u1ED1i c\xF9ng v\xE0 ch\u1ECBu tr\xE1ch nhi\u1EC7m khi s\u1EED d\u1EE5ng AI."
    },
    {
      "code": "6.A1.3",
      "mr": false,
      "desc": "Th\u1EF1c hi\u1EC7n \u0111\u01B0\u1EE3c vi\u1EC7c ki\u1EC3m tra l\u1EA1i m\u1ED9t k\u1EBFt qu\u1EA3 do AI \u0111\u01B0a ra (\u0111\u1ED1i chi\u1EBFu v\u1EDBi s\xE1ch gi\xE1o khoa, ngu\u1ED3n tin c\u1EADy kh\xE1c ho\u1EB7c h\u1ECFi th\u1EA7y c\xF4) tr\u01B0\u1EDBc khi s\u1EED d\u1EE5ng, th\u1EC3 hi\u1EC7n th\xF3i quen \u201Ccon ng\u01B0\u1EDDi quy\u1EBFt \u0111\u1ECBnh cu\u1ED1i c\xF9ng\u201D."
    },
    {
      "code": "6.A3.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 t\xECnh hu\u1ED1ng con ng\u01B0\u1EDDi ra quy\u1EBFt \u0111\u1ECBnh v\u1EDBi s\u1EF1 h\u1ED7 tr\u1EE3 c\u1EE7a AI (v\xED d\u1EE5: trong y t\u1EBF, AI g\u1EE3i \xFD ch\u1EA9n \u0111o\xE1n, b\xE1c s\u0129 xem x\xE9t v\xE0 quy\u1EBFt \u0111\u1ECBnh ph\u01B0\u01A1ng \xE1n \u0111i\u1EC1u tr\u1ECB; trong giao th\xF4ng, AI \u0111\u1EC1 xu\u1EA5t l\u1ED9 tr\xECnh nhanh h\u01A1n, ng\u01B0\u1EDDi l\xE1i xe c\xE2n nh\u1EAFc tr\u01B0\u1EDBc khi l\u1EF1a ch\u1ECDn). Th\u1EF1c h\xE0nh ra quy\u1EBFt \u0111\u1ECBnh trong m\u1ED9t t\xECnh hu\u1ED1ng gi\u1EA3 \u0111\u1ECBnh c\xF3 g\u1EE3i \xFD c\u1EE7a AI (n\xEAu l\u1EA1i g\u1EE3i \xFD c\u1EE7a AI, c\xE2n nh\u1EAFc v\xE0 \u0111\u01B0a ra l\u1EF1a ch\u1ECDn c\u1EE7a b\u1EA3n th\xE2n k\xE8m l\xED do)."
    },
    {
      "code": "6.A3.2",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c l\u1EE3i \xEDch c\u1EE7a AI trong vi\u1EC7c h\u1ED7 tr\u1EE3 con ng\u01B0\u1EDDi h\u1ECDc h\u1ECFi, r\xE8n luy\u1EC7n k\u0129 n\u0103ng v\xE0 m\u1EDF r\u1ED9ng hi\u1EC3u bi\u1EBFt khi \u0111\u01B0\u1EE3c s\u1EED d\u1EE5ng \u0111\xFAng c\xE1ch, c\xF3 m\u1EE5c \u0111\xEDch r\xF5 r\xE0ng. S\u1EED d\u1EE5ng \u0111\u01B0\u1EE3c m\u1ED9t c\xF4ng c\u1EE5 AI ph\xF9 h\u1EE3p l\u1EE9a tu\u1ED5i (tr\u1EF1c ti\u1EBFp ho\u1EB7c theo h\u01B0\u1EDBng d\u1EABn c\u1EE7a gi\xE1o vi\xEAn) \u0111\u1EC3 h\u1ED7 tr\u1EE3 m\u1ED9t nhi\u1EC7m v\u1EE5 h\u1ECDc t\u1EADp \u0111\u01A1n gi\u1EA3n v\xE0 n\xF3i l\u1EA1i \u0111\u01B0\u1EE3c \u0111i\u1EC1u m\xECnh \u0111\xE3 h\u1ECDc th\xEAm."
    },
    {
      "code": "6.A3.3",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c d\u1EEF li\u1EC7u c\xE1 nh\xE2n (h\xECnh \u1EA3nh, gi\u1ECDng n\xF3i, h\u1ECD t\xEAn, \u0111\u1ECBa ch\u1EC9, th\xF4ng tin h\u1ECDc t\u1EADp, th\xF3i quen s\u1EED d\u1EE5ng m\u1EA1ng,...) l\xE0 t\xE0i s\u1EA3n c\u1EE7a m\u1ED7i ng\u01B0\u1EDDi; ch\u1EC9 ch\u1EE7 s\u1EDF h\u1EEFu m\u1EDBi c\xF3 quy\u1EC1n quy\u1EBFt \u0111\u1ECBnh vi\u1EC7c chia s\u1EBB ho\u1EB7c cho ph\xE9p s\u1EED d\u1EE5ng d\u1EEF li\u1EC7u \u0111\xF3."
    },
    {
      "code": "6.A3.4",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c kh\xE1i ni\u1EC7m quy\u1EC1n ri\xEAng t\u01B0 \u1EDF m\u1EE9c \u0111\u01A1n gi\u1EA3n; n\xEAu \u0111\u01B0\u1EE3c t\xE1c h\u1EA1i khi d\u1EEF li\u1EC7u c\xE1 nh\xE2n b\u1ECB s\u1EED d\u1EE5ng sai m\u1EE5c \u0111\xEDch v\xE0 \xFD ngh\u0129a c\u1EE7a vi\u1EC7c b\u1EA3o v\u1EC7 quy\u1EC1n ri\xEAng t\u01B0 \u0111\u1ED1i v\u1EDBi s\u1EF1 an to\xE0n, danh d\u1EF1 v\xE0 t\u1EF1 do c\u1EE7a con ng\u01B0\u1EDDi trong m\xF4i tr\u01B0\u1EDDng s\u1ED1. Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xE1ch \u1EE9ng ph\xF3 ban \u0111\u1EA7u khi c\xF3 nguy c\u01A1 m\u1EA5t an to\xE0n d\u1EEF li\u1EC7u c\xE1 nh\xE2n (v\xED d\u1EE5: d\u1EEBng chia s\u1EBB; b\xE1o cho cha m\u1EB9, th\u1EA7y c\xF4; b\xE1o c\xE1o tr\xEAn n\u1EC1n t\u1EA3ng). B."
    },
    {
      "code": "6.B1.1",
      "mr": false,
      "desc": "Ch\u1EC9 ra \u0111\u01B0\u1EE3c m\u1EB7t t\xEDch c\u1EF1c v\xE0 h\u1EA1n ch\u1EBF c\u1EE7a m\u1ED9t s\u1ED1 t\xEDnh n\u0103ng AI c\u1EE5 th\u1EC3 (v\xED d\u1EE5: t\xEDnh n\u0103ng thu th\u1EADp d\u1EEF li\u1EC7u ng\u01B0\u1EDDi d\xF9ng gi\xFAp AI \u0111\u01B0a ra g\u1EE3i \xFD ch\xEDnh x\xE1c h\u01A1n nh\u01B0ng c\xF3 th\u1EC3 \u1EA3nh h\u01B0\u1EDFng \u0111\u1EBFn quy\u1EC1n ri\xEAng t\u01B0)."
    },
    {
      "code": "6.B2.1",
      "mr": false,
      "desc": "\u0110\u1EB7t \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xE2u h\u1ECFi \u0111\u01A1n gi\u1EA3n \u0111\u1EC3 ki\u1EC3m tra t\xEDnh an to\xE0n v\xE0 minh b\u1EA1ch c\u1EE7a \u1EE9ng d\u1EE5ng AI (v\xED d\u1EE5: \u201CC\xF4ng c\u1EE5 n\xE0y c\xF3 an to\xE0n kh\xF4ng?\u201D, \u201CC\xF4ng c\u1EE5 n\xE0y c\xF3 thu th\u1EADp th\xF4ng tin c\xE1 nh\xE2n kh\xF4ng?\u201D, \u201CC\xF3 th\u1EC3 t\u1EAFt t\xEDnh n\u0103ng n\xE0y kh\xF4ng?\u201D) v\xE0 \xE1p d\u1EE5ng c\xE1c c\xE2u h\u1ECFi \u0111\xF3 \u0111\u1EC3 nh\u1EADn x\xE9t \u0111\u01A1n gi\u1EA3n v\u1EC1 m\u1EE9c \u0111\u1ED9 an to\xE0n, ph\xF9 h\u1EE3p c\u1EE7a m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 AI. C."
    },
    {
      "code": "6.C1.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c hai th\xE0nh ph\u1EA7n ch\xEDnh \u0111\u1EC3 hu\u1EA5n luy\u1EC7n (\u201Cd\u1EA1y\u201D) AI l\xE0 d\u1EEF li\u1EC7u (v\xED d\u1EE5: v\u0103n b\u1EA3n, h\xECnh \u1EA3nh) v\xE0 thu\u1EADt to\xE1n (ph\u01B0\u01A1ng ph\xE1p h\u1ECDc t\u1EEB d\u1EEF li\u1EC7u). M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c c\xE1c b\u01B0\u1EDBc ho\u1EA1t \u0111\u1ED9ng ch\xEDnh c\u1EE7a m\u1ED9t c\xF4ng c\u1EE5 AI qua v\xED d\u1EE5 \u0111\u01A1n gi\u1EA3n (v\xED d\u1EE5: tr\u1EE3 l\xED \u1EA3o nh\u1EADn c\xE2u h\u1ECFi b\u1EB1ng gi\u1ECDng n\xF3i \u2192 chuy\u1EC3n gi\u1ECDng n\xF3i th\xE0nh v\u0103n b\u1EA3n \u2192 x\u1EED l\xED \u0111\u1EC3 hi\u1EC3u y\xEAu c\u1EA7u \u2192 t\xECm v\xE0 \u0111\u01B0a ra c\xE2u tr\u1EA3 l\u1EDDi)."
    },
    {
      "code": "6.C1.MR1",
      "mr": true,
      "desc": "Ch\u1EC9 ra \u0111\u01B0\u1EE3c m\u1ED1i li\xEAn h\u1EC7 gi\u1EEFa qu\xE1 tr\xECnh hu\u1EA5n luy\u1EC7n v\xE0 ho\u1EA1t \u0111\u1ED9ng c\u1EE7a m\u1ED9t c\xF4ng c\u1EE5 AI qua v\xED d\u1EE5 \u0111\u01A1n gi\u1EA3n (v\xED d\u1EE5: tr\u1EE3 l\xED \u1EA3o tr\u1EA3 l\u1EDDi \u0111\u01B0\u1EE3c c\xE2u h\u1ECFi nh\u1EDD \u0111\xE3 \u0111\u01B0\u1EE3c hu\u1EA5n luy\u1EC7n t\u1EEB d\u1EEF li\u1EC7u; d\u1EEF li\u1EC7u hu\u1EA5n luy\u1EC7n c\xE0ng phong ph\xFA, c\xE2u tr\u1EA3 l\u1EDDi c\xE0ng ch\xEDnh x\xE1c)."
    },
    {
      "code": "6.C1.MR2",
      "mr": true,
      "desc": "Th\u1EF1c hi\u1EC7n \u0111\u01B0\u1EE3c th\u1EED nghi\u1EC7m \u0111\u01A1n gi\u1EA3n v\u1EDBi m\u1ED9t c\xF4ng c\u1EE5 AI v\xE0 nh\u1EADn x\xE9t \u0111\u01B0\u1EE3c c\xE1c b\u01B0\u1EDBc ho\u1EA1t \u0111\u1ED9ng c\u1EE7a c\xF4ng c\u1EE5 \u0111\xF3 (v\xED d\u1EE5: \u0111\u1EB7t c\xF9ng m\u1ED9t c\xE2u h\u1ECFi cho tr\u1EE3 l\xED \u1EA3o theo c\xE1c c\xE1ch n\xF3i kh\xE1c nhau v\xE0 quan s\xE1t k\u1EBFt qu\u1EA3)."
    },
    {
      "code": "6.C1.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 t\xE1c \u0111\u1ED9ng t\xEDch c\u1EF1c v\xE0 t\xE1c \u0111\u1ED9ng ti\xEAu c\u1EF1c c\u1EE7a AI \u0111\u1ED1i v\u1EDBi cu\u1ED9c s\u1ED1ng c\u1EE7a b\u1EA3n th\xE2n, gia \u0111\xECnh."
    },
    {
      "code": "6.C1.MR3",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c \u0111\u01A1n gi\u1EA3n t\xE1c \u0111\u1ED9ng c\u1EE7a m\u1ED9t c\xF4ng c\u1EE5 AI c\u1EE5 th\u1EC3 \u0111\u1ED1i v\u1EDBi b\u1EA3n th\xE2n v\xE0 gia \u0111\xECnh, trong \u0111\xF3 ch\u1EC9 ra c\u1EA3 m\u1EB7t t\xEDch c\u1EF1c v\xE0 m\u1EB7t ti\xEAu c\u1EF1c (v\xED d\u1EE5: t\xEDnh n\u0103ng g\u1EE3i \xFD video gi\xFAp t\xECm \u0111\u01B0\u1EE3c n\u1ED9i dung ph\xF9 h\u1EE3p nh\u01B0ng d\u1EC5 l\xF4i cu\u1ED1n ng\u01B0\u1EDDi xem, g\xE2y m\u1EA5t nhi\u1EC1u th\u1EDDi gian)."
    },
    {
      "code": "6.C2.1",
      "mr": false,
      "desc": "Ph\xE2n bi\u1EC7t \u0111\u01B0\u1EE3c c\xF4ng c\u1EE5, t\xEDnh n\u0103ng c\xF3 \u1EE9ng d\u1EE5ng AI v\xE0 kh\xF4ng \u1EE9ng d\u1EE5ng AI qua c\xE1c v\xED d\u1EE5 g\u1EA7n g\u0169i trong cu\u1ED9c s\u1ED1ng (v\xED d\u1EE5: tr\u1EE3 l\xED \u1EA3o c\xF3 \u1EE9ng d\u1EE5ng AI; m\xE1y t\xEDnh b\u1ECF t\xFAi kh\xF4ng \u1EE9ng d\u1EE5ng AI)."
    },
    {
      "code": "6.C2.2",
      "mr": false,
      "desc": "K\u1EC3 \u0111\u01B0\u1EE3c t\xEAn v\xE0 m\xF4 t\u1EA3 \u0111\u01B0\u1EE3c ch\u1EE9c n\u0103ng ch\xEDnh c\u1EE7a m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 AI th\xF4ng d\u1EE5ng (nh\u01B0 tr\u1EE3 l\xED \u1EA3o; \u1EE9ng d\u1EE5ng b\u1EA3n \u0111\u1ED3; c\xF4ng c\u1EE5 d\u1ECBch thu\u1EADt;\u2026)."
    },
    {
      "code": "6.C2.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 \u1EE9ng d\u1EE5ng AI trong cu\u1ED9c s\u1ED1ng, \u01B0u ti\xEAn c\xE1c v\u1EA5n \u0111\u1EC1 g\u1EAFn v\u1EDBi th\u1EF1c ti\u1EC5n \u1EDF Vi\u1EC7t Nam (v\xED d\u1EE5: AI trong n\xF4ng nghi\u1EC7p, gi\xE1o d\u1EE5c, d\u1EF1 b\xE1o l\u0169, d\u1ECBch ng\xF4n ng\u1EEF c\xE1c d\xE2n t\u1ED9c,...)."
    },
    {
      "code": "6.C3.1",
      "mr": false,
      "desc": "K\u1EC3 \u0111\u01B0\u1EE3c t\xEAn m\u1ED9t s\u1ED1 c\xF4ng ngh\u1EC7 AI quen thu\u1ED9c trong \u0111\u1EDDi s\u1ED1ng (v\xED d\u1EE5: nh\u1EADn d\u1EA1ng h\xECnh \u1EA3nh, chuy\u1EC3n \u0111\u1ED5i gi\u1EEFa v\u0103n b\u1EA3n v\xE0 gi\u1ECDng n\xF3i,...)."
    },
    {
      "code": "6.C3.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 t\xEDnh n\u0103ng AI t\xEDch h\u1EE3p trong c\xE1c \u1EE9ng d\u1EE5ng th\u01B0\u1EDDng g\u1EB7p (v\xED d\u1EE5: g\u1EE3i \xFD \u0111\u1EC1 xu\u1EA5t n\u1ED9i dung tr\xEAn m\u1EA1ng x\xE3 h\u1ED9i, qu\u1EA3ng c\xE1o c\xE1 nh\xE2n ho\xE1,...) v\xE0 n\xEAu \u0111\u01B0\u1EE3c \u1EA3nh h\u01B0\u1EDFng c\u1EE7a c\xE1c t\xEDnh n\u0103ng \u0111\xF3 \u0111\u1EBFn vi\u1EC7c s\u1EED d\u1EE5ng v\xE0 ra quy\u1EBFt \u0111\u1ECBnh c\u1EE7a ng\u01B0\u1EDDi d\xF9ng. D."
    },
    {
      "code": "6.D1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng quen thu\u1ED9c n\xEAn ho\u1EB7c kh\xF4ng n\xEAn s\u1EED d\u1EE5ng AI (v\xED d\u1EE5: n\xEAn d\xF9ng AI \u0111\u1EC3 luy\u1EC7n ph\xE1t \xE2m ngo\u1EA1i ng\u1EEF, t\xECm \u0111\u01B0\u1EDDng \u0111i; kh\xF4ng n\xEAn nh\u1EDD AI vi\u1EBFt h\u1ED9 b\xE0i v\u0103n, l\xE0m h\u1ED9 b\xE0i t\u1EADp v\xEC l\xE0m gi\u1EA3m c\u01A1 h\u1ED9i t\u1EF1 r\xE8n luy\u1EC7n, l\xE0m gi\u1EA3m kh\u1EA3 n\u0103ng t\u1EF1 suy ngh\u0129 c\u1EE7a b\u1EA3n th\xE2n; kh\xF4ng n\xEAn d\xF9ng AI khi c\xF3 c\xE1ch gi\u1EA3i quy\u1EBFt kh\xE1c \u0111\u01A1n gi\u1EA3n v\xE0 an to\xE0n h\u01A1n)."
    },
    {
      "code": "6.D1.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c \xFD ki\u1EBFn c\xE1 nh\xE2n v\u1EC1 vi\u1EC7c n\xEAn hay kh\xF4ng n\xEAn s\u1EED d\u1EE5ng AI trong m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng th\u1EF1c t\u1EBF; gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c l\xED do l\u1EF1a ch\u1ECDn d\u1EF1a tr\xEAn l\u1EE3i \xEDch v\xE0 t\xE1c h\u1EA1i c\u1EE7a vi\u1EC7c s\u1EED d\u1EE5ng AI trong t\xECnh hu\u1ED1ng \u0111\xF3."
    },
    {
      "code": "6.D2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 gi\u1EDBi h\u1EA1n c\u1EE7a h\u1EC7 th\u1ED1ng AI so v\u1EDBi con ng\u01B0\u1EDDi qua c\xE1c t\xECnh hu\u1ED1ng, c\xF4ng vi\u1EC7c m\xE0 con ng\u01B0\u1EDDi th\u1EF1c hi\u1EC7n ph\xF9 h\u1EE3p h\u01A1n (ch\u1EB3ng h\u1EA1n nh\u01B0 c\xE1c c\xF4ng vi\u1EC7c c\u1EA7n s\u1EF1 th\u1EA5u hi\u1EC3u c\u1EA3m x\xFAc, s\xE1ng t\u1EA1o, ra quy\u1EBFt \u0111\u1ECBnh trong t\xECnh hu\u1ED1ng ph\u1EE9c t\u1EA1p)."
    },
    {
      "code": "6.D2.MR1",
      "mr": true,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c v\xEC sao AI g\u1EB7p gi\u1EDBi h\u1EA1n trong m\u1ED9t s\u1ED1 c\xF4ng vi\u1EC7c so v\u1EDBi con ng\u01B0\u1EDDi (v\xED d\u1EE5: AI h\u1ECDc t\u1EEB d\u1EEF li\u1EC7u c\xF3 s\u1EB5n n\xEAn kh\xF3 x\u1EED l\xED t\u1ED1t c\xE1c t\xECnh hu\u1ED1ng m\u1EDBi, ch\u01B0a t\u1EEBng c\xF3 trong d\u1EEF li\u1EC7u; AI kh\xF4ng c\xF3 c\u1EA3m x\xFAc th\u1EADt n\xEAn kh\xF3 th\u1EA5u hi\u1EC3u con ng\u01B0\u1EDDi)."
    },
    {
      "code": "6.D2.MR2",
      "mr": true,
      "desc": "\u0110\u1EC1 xu\u1EA5t \u0111\u01B0\u1EE3c c\xE1ch k\u1EBFt h\u1EE3p gi\u1EEFa con ng\u01B0\u1EDDi v\xE0 AI trong m\u1ED9t c\xF4ng vi\u1EC7c c\u1EE5 th\u1EC3 \u0111\u1EC3 ph\xE1t huy th\u1EBF m\u1EA1nh c\u1EE7a m\u1ED7i b\xEAn (v\xED d\u1EE5: AI g\u1EE3i \xFD \xFD t\u01B0\u1EDFng, con ng\u01B0\u1EDDi l\u1EF1a ch\u1ECDn v\xE0 ho\xE0n thi\u1EC7n)."
    }
  ],
  "7": [
    {
      "code": "7.A1.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c l\xED do con ng\u01B0\u1EDDi c\u1EA7n gi\u1EEF quy\u1EC1n ra quy\u1EBFt \u0111\u1ECBnh khi s\u1EED d\u1EE5ng AI (v\xED d\u1EE5 nh\u01B0 b\u1EA3o \u0111\u1EA3m c\xF4ng b\u1EB1ng: con ng\u01B0\u1EDDi c\xF3 th\u1EC3 xem x\xE9t ho\xE0n c\u1EA3nh v\xE0 c\u1EA3m x\xFAc m\xE0 AI kh\xF4ng hi\u1EC3u \u0111\u01B0\u1EE3c; b\u1EA3o \u0111\u1EA3m an to\xE0n: con ng\u01B0\u1EDDi c\xF3 th\u1EC3 ng\u0103n ch\u1EB7n c\xE1c h\xE0nh \u0111\u1ED9ng sai l\u1EC7ch ho\u1EB7c nguy hi\u1EC3m do AI g\xE2y ra; b\u1EA3o v\u1EC7 quy\u1EC1n l\u1EE3i con ng\u01B0\u1EDDi: kh\xF4ng \u0111\u1EC3 AI x\xE2m ph\u1EA1m quy\u1EC1n ri\xEAng t\u01B0, t\u1EF1 do ho\u1EB7c ph\u1EA9m gi\xE1 c\xE1 nh\xE2n)."
    },
    {
      "code": "7.A1.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 h\u1EADu qu\u1EA3 c\xF3 th\u1EC3 x\u1EA3y ra khi kh\xF4ng c\xF3 s\u1EF1 x\xE1c th\u1EF1c c\u1EE7a con ng\u01B0\u1EDDi v\u1EC1 \u0111\u1ED9 ch\xEDnh x\xE1c c\u1EE7a k\u1EBFt qu\u1EA3 do AI \u0111\u01B0a ra trong m\u1ED9t s\u1ED1 tr\u01B0\u1EDDng h\u1EE3p th\u1EF1c t\u1EBF."
    },
    {
      "code": "7.A1.MR1",
      "mr": true,
      "desc": "Th\u1EF1c hi\u1EC7n \u0111\u01B0\u1EE3c vi\u1EC7c ki\u1EC3m ch\u1EE9ng m\u1ED9t th\xF4ng tin do AI cung c\u1EA5p b\u1EB1ng \xEDt nh\u1EA5t m\u1ED9t ngu\u1ED3n \u0111\xE1ng tin c\u1EADy kh\xE1c (s\xE1ch gi\xE1o khoa, trang th\xF4ng tin ch\xEDnh th\u1ED1ng, \xFD ki\u1EBFn th\u1EA7y c\xF4) tr\u01B0\u1EDBc khi s\u1EED d\u1EE5ng."
    },
    {
      "code": "7.A2.1",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c c\xE1c t\xE1c h\u1EA1i c\xF3 th\u1EC3 x\u1EA3y ra n\u1EBFu con ng\u01B0\u1EDDi cho ph\xE9p AI \u0111\u01B0a ra quy\u1EBFt \u0111\u1ECBnh cu\u1ED1i c\xF9ng trong m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng th\u1EF1c t\u1EBF."
    },
    {
      "code": "7.A2.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1c h\u1EADu qu\u1EA3 c\xF3 th\u1EC3 x\u1EA3y ra n\u1EBFu kh\xF4ng c\xF3 c\xE1c quy \u0111\u1ECBnh ph\xE1p l\xFD nh\u1EB1m ng\u0103n ch\u1EB7n vi\u1EC7c thi\u1EBFt k\u1EBF v\xE0 s\u1EA3n xu\u1EA5t c\xE1c c\xF4ng c\u1EE5 AI c\xF3 h\u1EA1i trong m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng th\u1EF1c t\u1EBF."
    },
    {
      "code": "7.A3.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 t\xECnh hu\u1ED1ng AI \u0111\u01B0\u1EE3c ph\xE9p t\u1EF1 \u0111\u1ED9ng th\u1EF1c hi\u1EC7n m\u1ED9t s\u1ED1 thao t\xE1c thay con ng\u01B0\u1EDDi (v\xED d\u1EE5: t\u1EF1 \u0111\u1ED9ng ki\u1EC3m tra v\xE0 s\u1EEDa l\u1ED7i ch\xEDnh t\u1EA3 khi so\u1EA1n th\u1EA3o v\u0103n b\u1EA3n) v\xE0 t\xECnh hu\u1ED1ng con ng\u01B0\u1EDDi c\u1EA7n tr\u1EF1c ti\u1EBFp quy\u1EBFt \u0111\u1ECBnh."
    },
    {
      "code": "7.A3.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 xung \u0111\u1ED9t gi\u1EEFa quy\u1EC1n t\u1EF1 ch\u1EE7 c\u1EE7a con ng\u01B0\u1EDDi v\xE0 m\u1EE9c \u0111\u1ED9 t\u1EF1 ch\u1EE7 c\u1EE7a AI."
    },
    {
      "code": "7.A3.MR2",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c y\xEAu c\u1EA7u \u0111\xE1nh gi\xE1 m\u1EE9c \u0111\u1ED9 t\u1EF1 ch\u1EE7 c\u1EE7a AI d\u1EF1a tr\xEAn nhu c\u1EA7u v\xE0 y\u1EBFu t\u1ED1 b\u1ED1i c\u1EA3nh c\u1EE5 th\u1EC3."
    },
    {
      "code": "7.A3.2",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c s\u1EF1 c\u1EA7n thi\u1EBFt ph\u1EA3i b\u1EA3o v\u1EC7 quy\u1EC1n t\u1EF1 ch\u1EE7 c\u1EE7a con ng\u01B0\u1EDDi khi s\u1EED d\u1EE5ng AI \u0111\u1EC3 \u0111\u01B0a ra c\xE1c quy\u1EBFt \u0111\u1ECBnh quan tr\u1ECDng. B."
    },
    {
      "code": "7.B2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 ti\xEAu ch\xED \u0111\u01A1n gi\u1EA3n (d\u1EF1a tr\xEAn c\xE1c nguy\xEAn t\u1EAFc \u0111\u1EA1o \u0111\u1EE9c) \u0111\u1EC3 \u0111\xE1nh gi\xE1 m\u1EE9c \u0111\u1ED9 ph\xF9 h\u1EE3p, an to\xE0n c\u1EE7a m\u1ED9t \u1EE9ng d\u1EE5ng AI."
    },
    {
      "code": "7.B2.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 c\xE1c h\xE0nh \u0111\u1ED9ng c\u1EE5 th\u1EC3 g\xF3p ph\u1EA7n x\xE2y d\u1EF1ng m\xF4i tr\u01B0\u1EDDng AI c\xF3 \u0111\u1EA1o \u0111\u1EE9c (v\xED d\u1EE5: b\xE1o c\xE1o l\u1ED7i, kh\xF4ng s\u1EED d\u1EE5ng \u1EE9ng d\u1EE5ng \u0111\u1ED9c h\u1EA1i, y\xEAu c\u1EA7u s\u1EF1 minh b\u1EA1ch)."
    },
    {
      "code": "7.B3.1",
      "mr": false,
      "desc": "Th\u1EC3 hi\u1EC7n \u0111\u01B0\u1EE3c th\xE1i \u0111\u1ED9 v\xE0 cam k\u1EBFt c\xE1 nh\xE2n trong vi\u1EC7c s\u1EED d\u1EE5ng AI c\xF3 tr\xE1ch nhi\u1EC7m th\xF4ng qua m\u1ED9t s\u1ED1 h\xECnh th\u1EE9c \u0111\u01A1n gi\u1EA3n (v\xED d\u1EE5: b\xE0i vi\u1EBFt ng\u1EAFn, h\xF9ng bi\u1EC7n ho\u1EB7c s\u1EA3n ph\u1EA9m h\u1ECDc t\u1EADp c\u1EE5 th\u1EC3). Th\u1EC3 hi\u1EC7n \u0111\u01B0\u1EE3c vi\u1EC7c khai b\xE1o trung th\u1EF1c khi c\xF3 s\u1EED d\u1EE5ng AI trong s\u1EA3n ph\u1EA9m h\u1ECDc t\u1EADp. C."
    },
    {
      "code": "7.C4.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1c v\u1EA5n \u0111\u1EC1 \u0111\u1EA1o \u0111\u1EE9c c\xF3 th\u1EC3 n\u1EA3y sinh t\u1EEB d\u1EEF li\u1EC7u hu\u1EA5n luy\u1EC7n AI (v\xED d\u1EE5: d\u1EEF li\u1EC7u thi\u1EBFu \u0111a d\u1EA1ng d\u1EABn \u0111\u1EBFn ph\xE2n bi\u1EC7t \u0111\u1ED1i x\u1EED, d\u1EEF li\u1EC7u ri\xEAng t\u01B0 b\u1ECB x\xE2m ph\u1EA1m)."
    },
    {
      "code": "7.C4.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c t\u1EA7m quan tr\u1ECDng c\u1EE7a vi\u1EC7c s\u1EED d\u1EE5ng b\u1ED9 d\u1EEF li\u1EC7u \u201Cs\u1EA1ch\u201D v\xE0 \u201Cc\xF4ng b\u1EB1ng\u201D trong vi\u1EC7c t\u1EA1o ra c\xE1c c\xF4ng c\u1EE5 AI c\xF3 \u0111\u1EA1o \u0111\u1EE9c."
    },
    {
      "code": "7.C5.1",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c c\xE1c b\u01B0\u1EDBc ch\xEDnh trong qu\xE1 tr\xECnh hu\u1EA5n luy\u1EC7n AI qua m\u1ED9t v\xED d\u1EE5 c\u1EE5 th\u1EC3 (nh\u01B0 thu th\u1EADp v\xE0 g\xE1n nh\xE3n d\u1EEF li\u1EC7u \u1EA3nh ch\xF3, m\xE8o; cho m\xE1y h\u1ECDc t\u1EEB d\u1EEF li\u1EC7u; ki\u1EC3m tra kh\u1EA3 n\u0103ng ph\xE2n bi\u1EC7t tr\xEAn \u1EA3nh m\u1EDBi; \u0111i\u1EC1u ch\u1EC9nh n\u1EBFu k\u1EBFt qu\u1EA3 ch\u01B0a t\u1ED1t)."
    },
    {
      "code": "7.C5.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t s\u1ED1 c\xE1ch h\u1ECDc kh\xE1c nhau c\u1EE7a AI \u1EDF m\u1EE9c \u0111\u1ED9 \u0111\u01A1n gi\u1EA3n (v\xED d\u1EE5: h\u1ECDc t\u1EEB d\u1EEF li\u1EC7u \u0111\xE3 \u0111\u01B0\u1EE3c con ng\u01B0\u1EDDi ph\xE2n lo\u1EA1i, g\xE1n nh\xE3n s\u1EB5n; t\u1EF1 t\xECm ra quy lu\u1EADt t\u1EEB d\u1EEF li\u1EC7u; h\u1ECDc qua th\u1EED nghi\u1EC7m v\xE0 r\xFAt kinh nghi\u1EC7m t\u1EEB k\u1EBFt qu\u1EA3)."
    },
    {
      "code": "7.C5.MR1",
      "mr": true,
      "desc": "Ph\xE2n bi\u1EC7t \u0111\u01B0\u1EE3c ba ph\u01B0\u01A1ng ph\xE1p h\u1ECDc m\xE1y c\u01A1 b\u1EA3n (h\u1ECDc c\xF3 gi\xE1m s\xE1t, h\u1ECDc kh\xF4ng gi\xE1m s\xE1t, h\u1ECDc t\u0103ng c\u01B0\u1EDDng); n\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng AI s\u1EED d\u1EE5ng c\xE1c ph\u01B0\u01A1ng ph\xE1p h\u1ECDc m\xE1y \u0111\xF3. D."
    },
    {
      "code": "7.D1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t v\u1EA5n \u0111\u1EC1 trong tr\u01B0\u1EDDng h\u1ECDc ho\u1EB7c c\u1ED9ng \u0111\u1ED3ng c\xF3 th\u1EC3 \u0111\u01B0\u1EE3c gi\u1EA3i quy\u1EBFt ho\u1EB7c c\u1EA3i thi\u1EC7n b\u1EB1ng AI; m\xF4 t\u1EA3 \u0111\u01B0\u1EE3c ph\u1EA1m vi c\u1EE7a v\u1EA5n \u0111\u1EC1 \u0111\xF3 (v\xED d\u1EE5: trong tr\u01B0\u1EDDng h\u1ECDc c\xF3 th\u1EC3 x\xE2y d\u1EF1ng chatbot gi\u1EA3i \u0111\xE1p th\u1EAFc m\u1EAFc th\u01B0\u1EDDng g\u1EB7p v\u1EC1 n\u1ED9i quy, l\u1ECBch h\u1ECDc cho h\u1ECDc sinh; trong c\u1ED9ng \u0111\u1ED3ng c\xF3 th\u1EC3 d\xF9ng AI nh\u1EADn d\u1EA1ng h\xECnh \u1EA3nh h\u1ED7 tr\u1EE3 ph\xE2n lo\u1EA1i r\xE1c t\xE1i ch\u1EBF)."
    },
    {
      "code": "7.D1.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c t\xEDnh kh\u1EA3 thi c\u1EE7a m\u1ED9t \xFD t\u01B0\u1EDFng d\u1EF1 \xE1n AI b\u1EB1ng c\xE1ch tr\u1EA3 l\u1EDDi c\xE1c c\xE2u h\u1ECFi (v\xED d\u1EE5: d\u1EEF li\u1EC7u c\xF3 d\u1EC5 thu th\u1EADp kh\xF4ng, c\xF3 r\u1EE7i ro \u0111\u1EA1o \u0111\u1EE9c hay kh\xF4ng, m\u1EE9c \u0111\u1ED9 ph\u1EE9c t\u1EA1p v\xE0 chi ph\xED c\xF3 ph\xF9 h\u1EE3p hay kh\xF4ng)."
    },
    {
      "code": "7.D2.1",
      "mr": false,
      "desc": "L\u1EADp \u0111\u01B0\u1EE3c k\u1EBF ho\u1EA1ch cho m\u1ED9t d\u1EF1 \xE1n s\xE1ng t\u1EA1o c\xF3 s\u1EED d\u1EE5ng AI theo nh\xF3m nh\u1ECF."
    },
    {
      "code": "7.D2.MR1",
      "mr": true,
      "desc": "Th\u1EF1c h\xE0nh t\u1EA1o \u0111\u01B0\u1EE3c s\u1EA3n ph\u1EA9m \u0111\u01A1n gi\u1EA3n theo k\u1EBF ho\u1EA1ch \u0111\xE3 x\xE2y d\u1EF1ng."
    }
  ],
  "8": [
    {
      "code": "8.A1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 l\u0129nh v\u1EF1c m\xE0 AI kh\xF4ng n\xEAn thay th\u1EBF con ng\u01B0\u1EDDi (v\xED d\u1EE5: trong gi\xE1o d\u1EE5c, gi\xE1o vi\xEAn hi\u1EC3u t\xE2m l\xED h\u1ECDc sinh, bi\u1EBFt kh\xEDch l\u1EC7, h\u01B0\u1EDBng d\u1EABn v\xE0 d\u1EA1y \u0111\u1EA1o \u0111\u1EE9c; trong y t\u1EBF, b\xE1c s\u0129 kh\xF4ng ch\u1EC9 ch\u1EEFa b\u1EC7nh m\xE0 c\xF2n \u0111\u1ED9ng vi\xEAn, l\u1EAFng nghe b\u1EC7nh nh\xE2n; trong ngh\u1EC7 thu\u1EADt, ngh\u1EC7 s\u0129 th\u1EC3 hi\u1EC7n c\u1EA3m x\xFAc v\xE0 tr\u1EA3i nghi\u1EC7m ri\xEAng m\xE0 AI kh\xF4ng c\xF3 \u0111\u01B0\u1EE3c;...)."
    },
    {
      "code": "8.A1.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c nh\u1EEFng r\u1EE7i ro c\u1EE7a vi\u1EC7c l\u1EA1m d\u1EE5ng c\xE1c c\xF4ng c\u1EE5 AI t\u1EA1o sinh, li\xEAn h\u1EC7 v\u1EDBi nguy c\u01A1 suy gi\u1EA3m t\u01B0 duy ph\u1EA3n bi\u1EC7n, k\u0129 n\u0103ng s\xE1ng t\u1EA1o, s\u1EED d\u1EE5ng th\xF4ng tin sai l\u1EC7ch; t\u1EEB \u0111\xF3 n\xEAu \u0111\u01B0\u1EE3c s\u1EF1 c\u1EA7n thi\u1EBFt c\u1EE7a vi\u1EC7c ki\u1EC3m ch\u1EE9ng ngu\u1ED3n th\xF4ng tin khi s\u1EED d\u1EE5ng AI t\u1EA1o sinh trong h\u1ECDc t\u1EADp."
    },
    {
      "code": "8.A2.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c vi\u1EC7c m\u1ED9t s\u1ED1 h\u1EC7 th\u1ED1ng, d\u1ECBch v\u1EE5 s\u1EED d\u1EE5ng AI c\xF3 th\u1EC3 thu th\u1EADp, ph\xE2n t\xEDch v\xE0 s\u1EED d\u1EE5ng d\u1EEF li\u1EC7u c\xE1 nh\xE2n c\u1EE7a ng\u01B0\u1EDDi d\xF9ng (v\u1ECB tr\xED, th\xF3i quen, h\xECnh \u1EA3nh, gi\u1ECDng n\xF3i, s\u1EDF th\xEDch,...) \u0111\u1EC3 ki\u1EC3m so\xE1t h\xE0nh vi, g\xE2y \u1EA3nh h\u01B0\u1EDFng ho\u1EB7c thao t\xFAng quy\u1EBFt \u0111\u1ECBnh c\u1EE7a con ng\u01B0\u1EDDi."
    },
    {
      "code": "8.A2.2",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c hi\u1EC7n t\u01B0\u1EE3ng s\u1EED d\u1EE5ng AI \u0111\u1EC3 ki\u1EC3m so\xE1t ng\u01B0\u1EDDi d\xF9ng (v\xED d\u1EE5: vi\u1EC7c s\u1EED d\u1EE5ng AI kh\xF4ng minh b\u1EA1ch ho\u1EB7c sai m\u1EE5c \u0111\xEDch, khi\u1EBFn ng\u01B0\u1EDDi d\xF9ng b\u1ECB theo d\xF5i m\xE0 kh\xF4ng bi\u1EBFt; b\u1ECB \u0111\u1EC1 xu\u1EA5t th\xF4ng tin, qu\u1EA3ng c\xE1o ho\u1EB7c n\u1ED9i dung m\u1ED9t chi\u1EC1u; b\u1ECB thao t\xFAng suy ngh\u0129 ho\u1EB7c h\xE0nh \u0111\u1ED9ng; b\u1ECB m\u1EA5t quy\u1EC1n ki\u1EC3m so\xE1t d\u1EEF li\u1EC7u c\xE1 nh\xE2n)."
    },
    {
      "code": "8.A3.1",
      "mr": false,
      "desc": "Ph\xE2n bi\u1EC7t \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a ng\u01B0\u1EDDi d\xF9ng v\xE0 ng\u01B0\u1EDDi ph\xE1t tri\u1EC3n khi t\u01B0\u01A1ng t\xE1c v\u1EDBi AI."
    },
    {
      "code": "8.A3.MR1",
      "mr": true,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c vi\u1EC7c ng\u01B0\u1EDDi d\xF9ng c\u0169ng g\xF3p ph\u1EA7n t\xE1c \u0111\u1ED9ng \u0111\u1EBFn c\xF4ng c\u1EE5 AI trong qu\xE1 tr\xECnh s\u1EED d\u1EE5ng (v\xED d\u1EE5: d\u1EEF li\u1EC7u v\xE0 ph\u1EA3n h\u1ED3i c\u1EE7a ng\u01B0\u1EDDi d\xF9ng \u0111\u01B0\u1EE3c thu th\u1EADp \u0111\u1EC3 ti\u1EBFp t\u1EE5c hu\u1EA5n luy\u1EC7n, c\u1EA3i thi\u1EC7n h\u1EC7 th\u1ED1ng; n\u1ED9i dung ng\u01B0\u1EDDi d\xF9ng t\u01B0\u01A1ng t\xE1c nhi\u1EC1u s\u1EBD \u0111\u01B0\u1EE3c AI \u01B0u ti\xEAn \u0111\u1EC1 xu\u1EA5t)."
    },
    {
      "code": "8.A3.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 c\xE1c b\xEAn (ng\u01B0\u1EDDi s\xE1ng t\u1EA1o, nh\xE0 cung c\u1EA5p, ng\u01B0\u1EDDi s\u1EED d\u1EE5ng) ph\u1EA3i ch\u1ECBu tr\xE1ch nhi\u1EC7m ph\xE1p l\xFD \u0111\u1ED1i v\u1EDBi h\u1EADu qu\u1EA3 do c\xF4ng c\u1EE5 AI g\xE2y ra trong m\u1ED9t s\u1ED1 t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3 (v\xED d\u1EE5: ng\u01B0\u1EDDi d\xF9ng AI t\u1EA1o h\xECnh \u1EA3nh gi\u1EA3 m\u1EA1o ng\u01B0\u1EDDi kh\xE1c; nh\xE0 cung c\u1EA5p chatbot lan truy\u1EC1n th\xF4ng tin sai l\u1EC7ch)."
    },
    {
      "code": "8.A3.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c nh\u1EEFng vi\u1EC7c th\u1EC3 hi\u1EC7n tr\xE1ch nhi\u1EC7m gi\u1EA3i tr\xECnh khi s\u1EED d\u1EE5ng AI trong h\u1ECDc t\u1EADp v\xE0 t\u1EA1o s\u1EA3n ph\u1EA9m (v\xED d\u1EE5: cho bi\u1EBFt s\u1EA3n ph\u1EA9m c\xF3 s\u1EED d\u1EE5ng AI v\xE0 s\u1EED d\u1EE5ng \u1EDF ph\u1EA7n n\xE0o; n\xF3i r\xF5 c\xF4ng c\u1EE5 AI \u0111\xE3 d\xF9ng khi \u0111\u01B0\u1EE3c h\u1ECFi; ch\u1ECBu tr\xE1ch nhi\u1EC7m ki\u1EC3m tra t\xEDnh ch\xEDnh x\xE1c c\u1EE7a n\u1ED9i dung do AI t\u1EA1o ra tr\u01B0\u1EDBc khi s\u1EED d\u1EE5ng)."
    },
    {
      "code": "8.A3.MR2",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c nh\u1EEFng vi\u1EC7c th\u1EC3 hi\u1EC7n tr\xE1ch nhi\u1EC7m gi\u1EA3i tr\xECnh khi t\u1EA1o ra ho\u1EB7c thi\u1EBFt k\u1EBF AI (v\xED d\u1EE5: cho bi\u1EBFt s\u1EA3n ph\u1EA9m s\u1EED d\u1EE5ng d\u1EEF li\u1EC7u g\xEC, gi\u1EA3i th\xEDch c\xE1ch s\u1EA3n ph\u1EA9m ho\u1EA1t \u0111\u1ED9ng, nh\u1EADn tr\xE1ch nhi\u1EC7m v\u1EC1 k\u1EBFt qu\u1EA3 s\u1EA3n ph\u1EA9m t\u1EA1o ra). B."
    },
    {
      "code": "8.B1.1",
      "mr": false,
      "desc": "Nh\u1EADn di\u1EC7n v\xE0 ph\xE2n lo\u1EA1i \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 r\u1EE7i ro ph\u1ED5 bi\u1EBFn khi s\u1EED d\u1EE5ng AI: r\u1EE7i ro v\u1EC1 d\u1EEF li\u1EC7u v\xE0 quy\u1EC1n ri\xEAng t\u01B0; r\u1EE7i ro do thu\u1EADt to\xE1n thi\xEAn v\u1ECB ho\u1EB7c \u0111\u01B0a ra k\u1EBFt lu\u1EADn sai; r\u1EE7i ro b\u1ECB l\u1EEBa \u0111\u1EA3o b\u1EB1ng n\u1ED9i dung gi\u1EA3 m\u1EA1o do AI t\u1EA1o ra (h\xECnh \u1EA3nh, gi\u1ECDng n\xF3i, tin nh\u1EAFn gi\u1EA3)."
    },
    {
      "code": "8.B2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 vi\u1EC7c l\xE0m, c\xE1ch th\u1EE9c \u0111\u01A1n gi\u1EA3n \u0111\u1EC3 b\u1EA3o v\u1EC7 d\u1EEF li\u1EC7u c\xE1 nh\xE2n, t\xF4n tr\u1ECDng b\u1EA3n quy\u1EC1n v\xE0 ph\xF2ng tr\xE1nh, gi\u1EA3m thi\u1EC3u r\u1EE7i ro khi s\u1EED d\u1EE5ng, ph\xE1t tri\u1EC3n AI trong c\xE1c d\u1EF1 \xE1n h\u1ECDc t\u1EADp."
    },
    {
      "code": "8.B3.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1c v\u1EA5n \u0111\u1EC1 \u0111\u1EA1o \u0111\u1EE9c c\u1EA7n l\u01B0u \xFD khi ph\xE1t tri\u1EC3n AI (v\xED d\u1EE5: b\u1EA3o m\u1EADt th\xF4ng tin, kh\xF4ng cung c\u1EA5p th\xF4ng tin sai l\u1EC7ch, kh\xF4ng x\xFAc ph\u1EA1m ng\u01B0\u1EDDi kh\xE1c). C. C\xE1c k\u0129 thu\u1EADt v\xE0 \u1EE9ng d\u1EE5ng AI C\xE1ch AI th\u1EF1c hi\u1EC7n m\u1ED9t s\u1ED1 ch\u1EE9c n\u0103ng c\u01A1 b\u1EA3n."
    },
    {
      "code": "8.C1.1",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c \u0111\u01A1n gi\u1EA3n c\xE1ch AI th\u1EF1c hi\u1EC7n m\u1ED9t s\u1ED1 ch\u1EE9c n\u0103ng c\u01A1 b\u1EA3n nh\u01B0 \u201C\u0111\u1ECDc\u201D, \u201Cnghe\u201D, \u201Cnh\xECn\u201D (v\xED d\u1EE5: \u0111\u1EC3 \u201Cnghe\u201D, AI thu \xE2m thanh, chuy\u1EC3n gi\u1ECDng n\xF3i th\xE0nh v\u0103n b\u1EA3n r\u1ED3i ph\xE2n t\xEDch \u0111\u1EC3 hi\u1EC3u n\u1ED9i dung; \u0111\u1EC3 \u201Cnh\xECn\u201D, AI ph\xE2n t\xEDch c\xE1c \u0111\u1EB7c \u0111i\u1EC3m trong h\xECnh \u1EA3nh v\xE0 so s\xE1nh v\u1EDBi d\u1EEF li\u1EC7u \u0111\xE3 h\u1ECDc)."
    },
    {
      "code": "8.C1.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xF4ng ngh\u1EC7, k\u0129 thu\u1EADt \u0111\u1EA3m nhi\u1EC7m c\xE1c ch\u1EE9c n\u0103ng \u201C\u0111\u1ECDc\u201D, \u201Cnghe\u201D, \u201Cnh\xECn\u201D c\u1EE7a AI (v\xED d\u1EE5: x\u1EED l\xED ng\xF4n ng\u1EEF t\u1EF1 nhi\xEAn, nh\u1EADn d\u1EA1ng gi\u1ECDng n\xF3i, th\u1ECB gi\xE1c m\xE1y t\xEDnh)."
    },
    {
      "code": "8.C5.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c c\xE1ch AI nh\u1EADn di\u1EC7n c\u1EA3m x\xFAc d\u1EF1a v\xE0o c\xE1c \u0111\u1EB7c \u0111i\u1EC3m (v\xED d\u1EE5: n\xE9t m\u1EB7t; t\u1EEB kho\xE1 trong v\u0103n b\u1EA3n v\xE0 l\u1EDDi n\xF3i; ng\u1EEF \u0111i\u1EC7u c\u1EE7a l\u1EDDi n\xF3i; c\u1EED ch\u1EC9 nh\u01B0 g\u1EADt \u0111\u1EA7u, l\u1EAFc \u0111\u1EA7u,...)."
    },
    {
      "code": "8.C5.MR1",
      "mr": true,
      "desc": "Nh\u1EADn x\xE9t \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c \u0111\u01A1n gi\u1EA3n v\u1EC1 \u0111\u1ED9 tin c\u1EADy v\xE0 gi\u1EDBi h\u1EA1n c\u1EE7a vi\u1EC7c AI nh\u1EADn di\u1EC7n c\u1EA3m x\xFAc (v\xED d\u1EE5: c\xF9ng m\u1ED9t n\xE9t m\u1EB7t c\xF3 th\u1EC3 bi\u1EC3u th\u1ECB nh\u1EEFng c\u1EA3m x\xFAc kh\xE1c nhau t\xF9y theo b\u1ED1i c\u1EA3nh). D."
    },
    {
      "code": "8.D1.1",
      "mr": false,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 v\u1EA5n \u0111\u1EC1 th\u1EF1c t\u1EBF c\xF3 th\u1EC3 gi\u1EA3i quy\u1EBFt b\u1EB1ng AI."
    },
    {
      "code": "8.D1.MR1",
      "mr": true,
      "desc": "L\u1EADp \u0111\u01B0\u1EE3c k\u1EBF ho\u1EA1ch s\u01A1 b\u1ED9 cho d\u1EF1 \xE1n AI \u0111\u1EC3 gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1 \u0111\xE3 x\xE1c \u0111\u1ECBnh."
    },
    {
      "code": "8.D2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t k\u1ECBch b\u1EA3n h\u1ED9i tho\u1EA1i cho m\u1ED9t t\xECnh hu\u1ED1ng c\u1EE5 th\u1EC3 c\xF3 \u1EE9ng d\u1EE5ng AI (v\xED d\u1EE5: chatbot, tr\u1EE3 l\xED \u1EA3o)."
    },
    {
      "code": "8.D2.MR1",
      "mr": true,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \u0111\u1EB7c \u0111i\u1EC3m c\u01A1 b\u1EA3n c\u1EE7a m\u1ED9t tr\u1EA3i nghi\u1EC7m ng\u01B0\u1EDDi d\xF9ng (UX) t\u1ED1t khi t\u01B0\u01A1ng t\xE1c v\u1EDBi AI."
    },
    {
      "code": "8.D2.MR2",
      "mr": true,
      "desc": "L\u1EADp k\u1EBF ho\u1EA1ch v\xE0 tri\u1EC3n khai \u0111\u01B0\u1EE3c ho\u1EA1t \u0111\u1ED9ng l\xE0m vi\u1EC7c nh\xF3m \u0111\u1EC3 ph\xE1t tri\u1EC3n m\u1ED9t s\u1EA3n ph\u1EA9m AI \u0111\u01A1n gi\u1EA3n (v\xED d\u1EE5: chatbot, m\xF4 h\xECnh nh\u1EADn d\u1EA1ng) b\u1EB1ng c\xE1c c\xF4ng c\u1EE5 c\xF3 s\u1EB5n."
    }
  ],
  "9": [
    {
      "code": "9.A1.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c \xFD ki\u1EBFn c\xE1 nh\xE2n v\u1EC1 m\u1ED9t th\xE1ch th\u1EE9c m\xE0 x\xE3 h\u1ED9i \u0111ang \u0111\u1ED1i m\u1EB7t trong k\u1EC9 nguy\xEAn AI (v\xED d\u1EE5: hi sinh s\u1EF1 an to\xE0n c\u1EE7a ng\u01B0\u1EDDi d\xF9ng \u0111\u1EC3 \u01B0u ti\xEAn \u0111\u1ED5i m\u1EDBi v\xE0 ph\xE1t tri\u1EC3n AI)."
    },
    {
      "code": "9.A2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c d\u1EABn ch\u1EE9ng \u0111\u1EC3 gi\u1EA3i th\xEDch v\xEC sao AI c\xF3 t\xE1c \u0111\u1ED9ng l\u1EDBn \u0111\u1EBFn x\xE3 h\u1ED9i."
    },
    {
      "code": "9.A2.2",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c c\xE1c v\u1EA5n \u0111\u1EC1 \u201Cthi\xEAn v\u1ECB\u201D, \u201Cth\xE0nh ki\u1EBFn\u201D m\xE0 AI c\xF3 th\u1EC3 g\xE2y ra \u0111\u1ED1i v\u1EDBi x\xE3 h\u1ED9i."
    },
    {
      "code": "9.A3.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c nh\u1EEFng n\u0103ng l\u1EF1c con ng\u01B0\u1EDDi c\u1EA7n r\xE8n luy\u1EC7n trong th\u1EBF gi\u1EDBi c\xF3 AI, nh\u01B0 bi\u1EBFt h\u1ECDc c\xE1ch h\u1ECDc (r\xE8n luy\u1EC7n kh\u1EA3 n\u0103ng t\u1EF1 h\u1ECDc, h\u1ECDc su\u1ED1t \u0111\u1EDDi), ph\xE1t tri\u1EC3n t\u01B0 duy ph\u1EA3n bi\u1EC7n (bi\u1EBFt ph\xE2n t\xEDch, \u0111\xE1nh gi\xE1 th\xF4ng tin do AI cung c\u1EA5p), gi\u1EEF v\u1EEFng t\xEDnh s\xE1ng t\u1EA1o v\xE0 c\u1EA3m x\xFAc con ng\u01B0\u1EDDi (l\xE0m nh\u1EEFng vi\u1EC7c AI kh\xF4ng th\u1EC3 l\xE0m); n\xE2ng cao k\u0129 n\u0103ng h\u1EE3p t\xE1c v\xE0 giao ti\u1EBFp (l\xE0m vi\u1EC7c hi\u1EC7u qu\u1EA3 v\u1EDBi con ng\u01B0\u1EDDi v\xE0 AI)."
    },
    {
      "code": "9.A3.2",
      "mr": false,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c m\u1EE5c ti\xEAu h\u1ECDc t\u1EADp c\xE1 nh\xE2n trong th\u1EBF gi\u1EDBi c\xF3 AI (v\xED d\u1EE5: h\u1ECDc c\xE1ch s\u1EED d\u1EE5ng AI \u0111\u1EC3 t\xECm ki\u1EBFm v\xE0 t\xF3m t\u1EAFt th\xF4ng tin; s\u1EED d\u1EE5ng AI \u0111\u1EC3 r\xE8n luy\u1EC7n k\u0129 n\u0103ng vi\u1EBFt, \u0111\u1ECDc, v\u1EBD, t\u01B0 duy logic; luy\u1EC7n t\u1EADp giao ti\u1EBFp ho\u1EB7c h\u1ECDc ng\xF4n ng\u1EEF m\u1EDBi v\u1EDBi c\xF4ng c\u1EE5 AI; t\u1EF1 \u0111\u1EB7t m\u1EE5c ti\xEAu ph\xE1t tri\u1EC3n k\u0129 n\u0103ng m\u1EC1m m\xE0 AI kh\xF4ng thay th\u1EBF \u0111\u01B0\u1EE3c)."
    },
    {
      "code": "9.A3.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 c\xE1ch s\u1EED d\u1EE5ng AI nh\u01B0 m\u1ED9t c\xF4ng c\u1EE5 gi\xFAp th\u1EC3 hi\u1EC7n b\u1EA3n th\xE2n v\xE0 theo \u0111u\u1ED5i \u0111am m\xEA."
    },
    {
      "code": "9.A3.4",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c v\xED d\u1EE5 m\u1ED9t s\u1ED1 thay \u0111\u1ED5i m\xE0 AI c\xF3 th\u1EC3 mang l\u1EA1i cho c\xE1c ng\xE0nh ngh\u1EC1 trong t\u01B0\u01A1ng lai, t\u1EEB \u0111\xF3 n\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 k\u0129 n\u0103ng g\u1EAFn v\u1EDBi c\xE1c ng\xE0nh ngh\u1EC1 c\u1EE5 th\u1EC3 c\u1EA7n trang b\u1ECB \u0111\u1EC3 l\xE0m vi\u1EC7c hi\u1EC7u qu\u1EA3 c\xF9ng AI. B."
    },
    {
      "code": "9.B2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a ng\u01B0\u1EDDi d\xF9ng trong vi\u1EC7c ki\u1EC3m so\xE1t v\xE0 ch\u1ECBu tr\xE1ch nhi\u1EC7m \u0111\u1ED1i v\u1EDBi k\u1EBFt qu\u1EA3 cu\u1ED1i c\xF9ng do AI t\u1EA1o ra. Khai b\xE1o \u0111\u01B0\u1EE3c vi\u1EC7c s\u1EED d\u1EE5ng AI trong s\u1EA3n ph\u1EA9m h\u1ECDc t\u1EADp c\u1EE7a b\u1EA3n th\xE2n."
    },
    {
      "code": "9.B2.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a c\xE1 nh\xE2n v\xE0 c\u1ED9ng \u0111\u1ED3ng trong vi\u1EC7c gi\xE1m s\xE1t, ph\u1EA3n h\u1ED3i v\xE0 \u0111\u1EC1 xu\u1EA5t gi\u1EA3i ph\xE1p \u0111\u1EC3 s\u1EED d\u1EE5ng AI m\u1ED9t c\xE1ch an to\xE0n, c\xF4ng b\u1EB1ng, h\u1EE3p l\xED."
    },
    {
      "code": "9.B2.3",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 d\u1EA5u hi\u1EC7u c\u1EE7a n\u1ED9i dung gi\u1EA3 m\u1EA1o do AI t\u1EA1o ra; ki\u1EC3m ch\u1EE9ng \u0111\u01B0\u1EE3c th\xF4ng tin b\u1EB1ng c\xE1c ngu\u1ED3n ph\xF9 h\u1EE3p v\xE0 \u0111\u1EC1 xu\u1EA5t \u0111\u01B0\u1EE3c c\xE1ch \u1EE9ng ph\xF3 khi ph\xE1t hi\u1EC7n n\u1ED9i dung c\xF3 d\u1EA5u hi\u1EC7u gi\u1EA3 m\u1EA1o ho\u1EB7c l\u1EEBa \u0111\u1EA3o."
    },
    {
      "code": "9.B3.1",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c t\u1EA7m quan tr\u1ECDng c\u1EE7a vi\u1EC7c hu\u1EA5n luy\u1EC7n AI theo h\u01B0\u1EDBng kh\xF4ng ph\xE2n bi\u1EC7t \u0111\u1ED1i x\u1EED v\xE0 t\xF4n tr\u1ECDng s\u1EF1 \u0111a d\u1EA1ng."
    },
    {
      "code": "9.B3.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xE1ch thu th\u1EADp d\u1EEF li\u1EC7u b\u1EA3o \u0111\u1EA3m c\xF4ng b\u1EB1ng, kh\xF4ng b\u1ECF s\xF3t hay thi\xEAn v\u1ECB c\xE1c nh\xF3m \u0111\u1ED1i t\u01B0\u1EE3ng kh\xE1c nhau. C."
    },
    {
      "code": "9.C2.1",
      "mr": false,
      "desc": "\u0110\u1EC1 xu\u1EA5t \u0111\u01B0\u1EE3c \xFD t\u01B0\u1EDFng m\u1EDBi, s\xE1ng t\u1EA1o \u0111\u1EC3 gi\u1EA3i quy\u1EBFt m\u1ED9t v\u1EA5n \u0111\u1EC1 b\u1EB1ng AI."
    },
    {
      "code": "9.C2.MR1",
      "mr": true,
      "desc": "V\u1EADn d\u1EE5ng \u0111\u01B0\u1EE3c ki\u1EBFn th\u1EE9c \u0111\xE3 h\u1ECDc \u0111\u1EC3 t\u1EA1o ra m\u1ED9t c\xF4ng c\u1EE5 AI \u0111\u01A1n gi\u1EA3n (v\xED d\u1EE5: chatbot, m\xF4 h\xECnh nh\u1EADn d\u1EA1ng h\xECnh \u1EA3nh) d\u1EF1a tr\xEAn c\xE1c c\xF4ng c\u1EE5 ho\u1EB7c n\u1EC1n t\u1EA3ng c\xF3 s\u1EB5n; \u01B0u ti\xEAn n\u1EC1n t\u1EA3ng m\u1EDF, mi\u1EC5n ph\xED, d\u1EC5 s\u1EED d\u1EE5ng, ph\xF9 h\u1EE3p \u0111\u1ED9 tu\u1ED5i v\xE0 \u0111\u01B0\u1EE3c r\xE0 so\xE1t v\u1EC1 an to\xE0n. C\xE1ch c\u1EA3i thi\u1EC7n d\u1EEF li\u1EC7u, n\xE2ng cao ch\u1EA5t l\u01B0\u1EE3ng c\u1EE7a s\u1EA3n ph\u1EA9m AI."
    },
    {
      "code": "9.C4.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xE1ch c\u1EA3i thi\u1EC7n b\u1ED9 d\u1EEF li\u1EC7u \u0111\u1EC3 n\xE2ng cao ch\u1EA5t l\u01B0\u1EE3ng c\u1EE7a s\u1EA3n ph\u1EA9m AI (v\xED d\u1EE5: b\u1ED5 sung d\u1EEF li\u1EC7u c\xF2n thi\u1EBFu, lo\u1EA1i b\u1ECF d\u1EEF li\u1EC7u tr\xF9ng l\u1EB7p, s\u1EEDa d\u1EEF li\u1EC7u sai ho\u1EB7c g\xE1n nh\xE3n l\u1EA1i cho \u0111\xFAng)."
    },
    {
      "code": "9.C4.MR1",
      "mr": true,
      "desc": "Ch\u1EC9 ra \u0111\u01B0\u1EE3c nh\u1EEFng \u0111i\u1EC3m c\u1EA7n c\u1EA3i thi\u1EC7n trong b\u1ED9 d\u1EEF li\u1EC7u c\u1EE5 th\u1EC3 \u1EA3nh h\u01B0\u1EDFng \u0111\u1EBFn ch\u1EA5t l\u01B0\u1EE3ng c\u1EE7a s\u1EA3n ph\u1EA9m AI."
    },
    {
      "code": "9.C4.MR2",
      "mr": true,
      "desc": "Th\u1EF1c hi\u1EC7n \u0111\u01B0\u1EE3c vi\u1EC7c c\u1EA3i thi\u1EC7n b\u1ED9 d\u1EEF li\u1EC7u (th\xEAm, xo\xE1, s\u1EEDa d\u1EEF li\u1EC7u) \u0111\u1EC3 n\xE2ng cao ch\u1EA5t l\u01B0\u1EE3ng c\u1EE7a s\u1EA3n ph\u1EA9m AI. D."
    },
    {
      "code": "9.D1.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a con ng\u01B0\u1EDDi l\xE0 ng\u01B0\u1EDDi \u0111\u1ED3ng s\xE1ng t\u1EA1o v\xE0 d\u1EABn d\u1EAFt trong vi\u1EC7c thi\u1EBFt k\u1EBF, v\u1EADn h\xE0nh v\xE0 ph\xE1t tri\u1EC3n c\xE1c h\u1EC7 th\u1ED1ng, c\xF4ng c\u1EE5 AI."
    },
    {
      "code": "9.D1.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c vai tr\xF2 \u0111\u1ED3ng s\xE1ng t\u1EA1o v\xE0 d\u1EABn d\u1EAFt c\u1EE7a con ng\u01B0\u1EDDi qua m\u1ED9t h\u1EC7 th\u1ED1ng ho\u1EB7c c\xF4ng c\u1EE5 AI c\u1EE5 th\u1EC3 (v\xED d\u1EE5: con ng\u01B0\u1EDDi x\xE1c \u0111\u1ECBnh m\u1EE5c ti\xEAu, l\u1EF1a ch\u1ECDn d\u1EEF li\u1EC7u, \u0111\xE1nh gi\xE1 v\xE0 \u0111i\u1EC1u ch\u1EC9nh k\u1EBFt qu\u1EA3 c\u1EE7a h\u1EC7 th\u1ED1ng)."
    },
    {
      "code": "9.D2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xE1ch ki\u1EC3m tra \u0111\u01A1n gi\u1EA3n \u0111\u1EC3 \u0111\xE1nh gi\xE1 s\u1EA3n ph\u1EA9m AI (v\xED d\u1EE5: v\u1EDBi m\xF4 h\xECnh nh\u1EADn d\u1EA1ng, th\u1EED v\u1EDBi d\u1EEF li\u1EC7u m\u1EDBi ch\u01B0a d\xF9ng khi hu\u1EA5n luy\u1EC7n, \u0111\u1EBFm s\u1ED1 l\u1EA7n cho k\u1EBFt qu\u1EA3 \u0111\xFAng v\xE0 sai; v\u1EDBi chatbot, \u0111\u1EB7t nhi\u1EC1u c\xE2u h\u1ECFi kh\xE1c nhau, ki\u1EC3m tra t\xEDnh ch\xEDnh x\xE1c v\xE0 ph\xF9 h\u1EE3p c\u1EE7a c\xE2u tr\u1EA3 l\u1EDDi)."
    },
    {
      "code": "9.D2.MR1",
      "mr": true,
      "desc": "Thi\u1EBFt k\u1EBF v\xE0 th\u1EF1c hi\u1EC7n \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 ki\u1EC3m tra \u0111\u01A1n gi\u1EA3n \u0111\u1EC3 \u0111\xE1nh gi\xE1 s\u1EA3n ph\u1EA9m AI."
    },
    {
      "code": "9.D2.MR2",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c k\u1EBFt qu\u1EA3 ki\u1EC3m tra v\xE0 ch\u1EE7 \u0111\u1ED9ng th\u1EED nghi\u1EC7m m\u1ED9t s\u1ED1 c\xE1ch c\u1EA3i ti\u1EBFn \u0111\u01A1n gi\u1EA3n, ph\xF9 h\u1EE3p nh\u1EB1m n\xE2ng cao ch\u1EA5t l\u01B0\u1EE3ng s\u1EA3n ph\u1EA9m."
    }
  ],
  "10": [
    {
      "code": "10.A1.1",
      "mr": false,
      "desc": "Th\u1EF1c h\xE0nh x\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a con ng\u01B0\u1EDDi trong s\u1EED d\u1EE5ng, v\u1EADn h\xE0nh, t\xF9y ch\u1EC9nh m\u1ED9t h\u1EC7 th\u1ED1ng AI c\u1EE5 th\u1EC3."
    },
    {
      "code": "10.A1.2",
      "mr": false,
      "desc": "Gi\u1EA3i th\xEDch \u0111\u01B0\u1EE3c t\u1EA1i sao vi\u1EC7c con ng\u01B0\u1EDDi ki\u1EC3m so\xE1t AI l\xE0 quan tr\u1ECDng, th\xF4ng qua vi\u1EC7c li\xEAn h\u1EC7 \u0111\u1EBFn c\xE1c gi\xE1 tr\u1ECB nh\u01B0 an to\xE0n, c\xF4ng b\u1EB1ng v\xE0 quy\u1EC1n l\u1EE3i con ng\u01B0\u1EDDi."
    },
    {
      "code": "10.A2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 r\u1EE7i ro \u0111\u1ED1i v\u1EDBi con ng\u01B0\u1EDDi, x\xE3 h\u1ED9i m\xE0 m\u1ED9t s\u1EA3n ph\u1EA9m AI c\xF3 th\u1EC3 \u0111em l\u1EA1i."
    },
    {
      "code": "10.A2.MR1",
      "mr": true,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 bi\u1EC7n ph\xE1p h\u1EA1n ch\u1EBF c\xE1c r\u1EE7i ro \u0111\u1ED1i v\u1EDBi con ng\u01B0\u1EDDi, x\xE3 h\u1ED9i m\xE0 m\u1ED9t s\u1EA3n ph\u1EA9m AI c\xF3 th\u1EC3 \u0111em l\u1EA1i th\xF4ng qua m\u1ED9t d\u1EF1 \xE1n s\xE1ng t\u1EA1o AI."
    },
    {
      "code": "10.A3.1",
      "mr": false,
      "desc": "K\u1EC3 t\xEAn \u0111\u01B0\u1EE3c m\u1ED9t v\xE0i quy \u0111\u1ECBnh ho\u1EB7c lu\u1EADt l\u1EC7 (\u1EDF m\u1EE9c \u0111\u1ED9 kh\xE1i ni\u1EC7m, v\xED d\u1EE5: Lu\u1EADt An ninh m\u1EA1ng, Lu\u1EADt D\u1EEF li\u1EC7u, Lu\u1EADt B\u1EA3o v\u1EC7 d\u1EEF li\u1EC7u c\xE1 nh\xE2n) c\xF3 ch\u1EE9c n\u0103ng b\u1EA3o v\u1EC7 ng\u01B0\u1EDDi d\xF9ng trong kh\xF4ng gian s\u1ED1. B."
    },
    {
      "code": "10.B2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 h\xE0nh vi s\u1EED d\u1EE5ng AI ho\u1EB7c s\u1EF1 c\u1ED1 li\xEAn quan \u0111\u1EBFn AI vi ph\u1EA1m quy \u0111\u1ECBnh c\u1EE7a nh\xE0 tr\u01B0\u1EDDng ho\u1EB7c c\xE1c v\u0103n b\u1EA3n ph\xE1p lu\u1EADt li\xEAn quan \u0111\u1EBFn s\u1EED d\u1EE5ng c\xF4ng ngh\u1EC7 th\xF4ng tin."
    },
    {
      "code": "10.B2.MR1",
      "mr": true,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 d\u1EA5u hi\u1EC7u c\u1EE7a n\u1ED9i dung do AI t\u1EA1o sinh t\u1EA1o ra; ki\u1EC3m tra v\xE0 nh\u1EADn x\xE9t \u0111\u01B0\u1EE3c m\u1EE9c \u0111\u1ED9 minh b\u1EA1ch c\u1EE7a vi\u1EC7c khai b\xE1o s\u1EED d\u1EE5ng AI trong m\u1ED9t s\u1EA3n ph\u1EA9m."
    },
    {
      "code": "10.B3.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c v\xED d\u1EE5 minh h\u1ECDa m\u1ED9t s\u1ED1 v\u1EA5n \u0111\u1EC1 \u0111\u1EA1o \u0111\u1EE9c c\xF3 th\u1EC3 ph\xE1t sinh trong qu\xE1 tr\xECnh thi\u1EBFt k\u1EBF v\xE0 v\u1EADn h\xE0nh AI (nh\u01B0 thi\xEAn v\u1ECB d\u1EEF li\u1EC7u, vi ph\u1EA1m quy\u1EC1n ri\xEAng t\u01B0 ho\u1EB7c thi\u1EBFu minh b\u1EA1ch). C."
    },
    {
      "code": "10.C2.1",
      "mr": false,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c c\xE1c v\u1EA5n \u0111\u1EC1 th\u1EF1c t\u1EBF c\xF3 th\u1EC3 \u1EE9ng d\u1EE5ng AI \u0111\u1EC3 th\u1EF1c hi\u1EC7n. \u01AFu ti\xEAn c\xE1c v\u1EA5n \u0111\u1EC1 g\u1EA7n g\u0169i, c\u1EA7n thi\u1EBFt trong b\u1ED1i c\u1EA3nh Vi\u1EC7t Nam, ch\u1EB3ng h\u1EA1n: s\u1EA3n xu\u1EA5t n\xF4ng nghi\u1EC7p, c\xE1c v\u1EA5n \u0111\u1EC1 li\xEAn quan \u0111\u1EBFn c\xE1c c\u1ED9ng \u0111\u1ED3ng thi\u1EC3u s\u1ED1, \u2026"
    },
    {
      "code": "10.C2.2",
      "mr": false,
      "desc": "Li\u1EC7t k\xEA \u0111\u01B0\u1EE3c t\xEAn c\xE1c \u1EE9ng d\u1EE5ng AI theo c\xE1c t\xEDnh n\u0103ng c\u1EE7a h\u1EC7 th\u1ED1ng."
    },
    {
      "code": "10.C2.MR1",
      "mr": true,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c c\xE1c y\xEAu c\u1EA7u c\u1EA7n c\xF3 \u0111\u1ED1i v\u1EDBi vi\u1EC7c \u1EE9ng d\u1EE5ng AI th\u1EF1c hi\u1EC7n nhi\u1EC7m v\u1EE5 c\u1EE5 th\u1EC3."
    },
    {
      "code": "10.C2.3",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 m\u1ED9t s\u1ED1 tr\u01B0\u1EDDng h\u1EE3p s\u1EED d\u1EE5ng AI h\u1ED7 tr\u1EE3 qu\xE1 tr\xECnh h\u1ECDc t\u1EADp."
    },
    {
      "code": "10.C2.MR2",
      "mr": true,
      "desc": "S\u1EED d\u1EE5ng \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng AI trong h\u1ECDc t\u1EADp."
    },
    {
      "code": "10.C3.1",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c c\xE1c y\xEAu c\u1EA7u \u0111\u1EC3 \u0111\u01B0a ra prompt ph\xF9 h\u1EE3p v\u1EDBi m\u1EE5c ti\xEAu c\u1EE5 th\u1EC3."
    },
    {
      "code": "10.C3.2",
      "mr": false,
      "desc": "Th\u1EF1c h\xE0nh \u0111\u1EB7t prompt gi\u1EA3i quy\u1EBFt m\u1ED9t s\u1ED1 v\u1EA5n \u0111\u1EC1 g\u1EA7n g\u0169i trong cu\u1ED9c s\u1ED1ng, h\u1ECDc t\u1EADp m\u1ED9t c\xE1ch hi\u1EC7u qu\u1EA3."
    },
    {
      "code": "10.C3.3",
      "mr": false,
      "desc": "Ph\xE2n bi\u1EC7t \u0111\u01B0\u1EE3c AI t\u1EA1o sinh v\u1EDBi c\xE1c h\u1EC7 th\u1ED1ng AI ph\xE2n lo\u1EA1i, d\u1EF1 \u0111o\xE1n qua v\xED d\u1EE5 c\u1EE5 th\u1EC3."
    },
    {
      "code": "10.C3.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c v\xED d\u1EE5 m\xF4 t\u1EA3 m\u1ED9t s\u1ED1 c\xF4ng ngh\u1EC7 \u0111\u1EC3 thi\u1EBFt k\u1EBF v\xE0 t\u1EA1o AI."
    },
    {
      "code": "10.C4.1",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c s\u1EF1 \u1EA3nh h\u01B0\u1EDFng c\u1EE7a ch\u1EA5t l\u01B0\u1EE3ng d\u1EEF li\u1EC7u \u0111\u1EBFn ch\u1EA5t l\u01B0\u1EE3ng AI."
    },
    {
      "code": "10.C4.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c c\xE1c d\u1EA1ng d\u1EEF li\u1EC7u (h\xECnh \u1EA3nh, \xE2m thanh, t\u1EEB ng\u1EEF, \u2026) \u0111\u01B0\u1EE3c s\u1EED d\u1EE5ng \u0111\u1EC3 hu\u1EA5n luy\u1EC7n AI. D."
    },
    {
      "code": "10.D1.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 c\u1EE5 th\u1EC3, x\xE1c \u0111\u1ECBnh nhi\u1EC7m v\u1EE5 ho\u1EB7c m\u1EE5c ti\xEAu c\u1EE5 th\u1EC3 m\xE0 m\u1ED9t h\u1EC7 th\u1ED1ng AI c\u1EA7n th\u1EF1c hi\u1EC7n, n\xEAu \u0111\u01B0\u1EE3c m\u1ED1i li\xEAn h\u1EC7 gi\u1EEFa m\u1EE5c ti\xEAu \u0111\xF3 v\u1EDBi c\xE1c th\xE0nh ph\u1EA7n ch\xEDnh c\u1EE7a h\u1EC7 th\u1ED1ng."
    },
    {
      "code": "10.D2.1",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c c\xE1c th\xE0nh ph\u1EA7n c\u01A1 b\u1EA3n c\u1EE7a h\u1EC7 th\u1ED1ng AI (d\u1EEF li\u1EC7u, m\xF4 h\xECnh, thu\u1EADt to\xE1n, \u0111\u1EA7u ra, ph\u1EA3n h\u1ED3i) ph\xF9 h\u1EE3p v\u1EDBi nhi\u1EC7m v\u1EE5 c\u1EE5 th\u1EC3."
    },
    {
      "code": "10.D2.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t s\u1ED1 v\u1EA5n \u0111\u1EC1 ph\xE1t sinh trong qu\xE1 tr\xECnh v\u1EADn h\xE0nh ho\u1EB7c t\u1ED1i \u01B0u ho\xE1 AI v\xE0 tr\xECnh b\xE0y \u0111\u01B0\u1EE3c \xFD ngh\u0129a c\u1EE7a vi\u1EC7c kh\u1EAFc ph\u1EE5c c\xE1c v\u1EA5n \u0111\u1EC1 \u0111\xF3."
    }
  ],
  "11": [
    {
      "code": "11.A1.1",
      "mr": false,
      "desc": "X\xE2y d\u1EF1ng \u0111\u01B0\u1EE3c quy tr\xECnh s\u1EED d\u1EE5ng m\u1ED9t s\u1EA3n ph\u1EA9m AI c\u1EE5 th\u1EC3 m\u1ED9t c\xE1ch th\xEDch h\u1EE3p."
    },
    {
      "code": "11.A1.2",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c c\xE1c tr\u01B0\u1EDDng h\u1EE3p th\u1EF1c t\u1EBF \u0111\u1EC3 th\u1EA5y \u0111\u01B0\u1EE3c t\u1EA7m quan tr\u1ECDng c\u1EE7a vi\u1EC7c s\u1EED d\u1EE5ng AI \u0111\u1EC3 n\xE2ng cao n\u0103ng l\u1EF1c c\u1EE7a con ng\u01B0\u1EDDi m\xE0 v\u1EABn \u0111\u1EA3m b\u1EA3o s\u1EF1 ki\u1EC3m so\xE1t c\u1EE7a con ng\u01B0\u1EDDi."
    },
    {
      "code": "11.A2.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng AI c\xF3 t\xE1c \u0111\u1ED9ng t\xEDch c\u1EF1c v\xE0 mang l\u1EA1i l\u1EE3i \xEDch x\xE3 h\u1ED9i l\xE2u d\xE0i, ch\u1EB3ng h\u1EA1n nh\u01B0: AI trong n\xF4ng nghi\u1EC7p, AI trong y t\u1EBF, \u2026"
    },
    {
      "code": "11.A2.2",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c c\xE1c y\u1EBFu t\u1ED1 th\u1EC3 hi\u1EC7n t\xEDnh b\u1EC1n v\u1EEFng v\xE0 c\xF4ng b\u1EB1ng c\u1EE7a h\u1EC7 th\u1ED1ng AI \u0111\xF3, v\xED d\u1EE5: b\u1EC1n v\u1EEFng - s\u1EED d\u1EE5ng n\u0103ng l\u01B0\u1EE3ng hi\u1EC7u qu\u1EA3, g\xF3p ph\u1EA7n b\u1EA3o v\u1EC7 m\xF4i tr\u01B0\u1EDDng, duy tr\xEC l\xE2u d\xE0i m\xE0 kh\xF4ng g\xE2y h\u1EA1i \u0111\u1EBFn con ng\u01B0\u1EDDi hay thi\xEAn nhi\xEAn; c\xF4ng b\u1EB1ng - \u0111\u1EA3m b\u1EA3o m\u1ECDi ng\u01B0\u1EDDi \u0111\u1EC1u \u0111\u01B0\u1EE3c h\u01B0\u1EDFng l\u1EE3i t\u1EEB c\xF4ng ngh\u1EC7, kh\xF4ng ph\xE2n bi\u1EC7t gi\u1EDBi t\xEDnh, \u0111\u1ECBa v\u1ECB, v\xF9ng mi\u1EC1n hay \u0111i\u1EC1u ki\u1EC7n kinh t\u1EBF."
    },
    {
      "code": "11.A3.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1c quy\u1EC1n c\u01A1 b\u1EA3n c\u1EE7a ng\u01B0\u1EDDi d\xF9ng d\u1EEF li\u1EC7u (quy\u1EC1n \u0111\u01B0\u1EE3c bi\u1EBFt, quy\u1EC1n \u0111\u01B0\u1EE3c \u0111\u1ED3ng \xFD, quy\u1EC1n y\xEAu c\u1EA7u xo\xE1 d\u1EEF li\u1EC7u\u2026)."
    },
    {
      "code": "11.A3.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c m\u1EE9c \u0111\u1ED9 \u0111\u1EA3m b\u1EA3o c\xE1c quy\u1EC1n c\u01A1 b\u1EA3n c\u1EE7a ng\u01B0\u1EDDi d\xF9ng \u0111\u1ED1i v\u1EDBi m\u1ED9t s\u1ED1 s\u1EA3n ph\u1EA9m AI th\xF4ng qua m\u1ED9t d\u1EF1 \xE1n s\xE1ng t\u1EA1o AI. B."
    },
    {
      "code": "11.B2.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt v\xE0 ph\xE2n lo\u1EA1i \u0111\u01B0\u1EE3c c\xE1c r\u1EE7i ro ho\u1EB7c s\u1EF1 c\u1ED1 li\xEAn quan \u0111\u1EBFn vi\u1EC7c s\u1EED d\u1EE5ng AI c\xF3 th\u1EC3 d\u1EABn \u0111\u1EBFn vi ph\u1EA1m quy \u0111\u1ECBnh c\u1EE7a nh\xE0 tr\u01B0\u1EDDng ho\u1EB7c c\xE1c v\u0103n b\u1EA3n ph\xE1p lu\u1EADt li\xEAn quan."
    },
    {
      "code": "11.B3.MR1",
      "mr": true,
      "desc": "X\xE1c \u0111\u1ECBnh v\xE0 s\u01A1 \u0111\u1ED3 ho\xE1 \u0111\u01B0\u1EE3c c\xE1c v\u1EA5n \u0111\u1EC1 \u0111\u1EA1o \u0111\u1EE9c c\xF3 th\u1EC3 ph\xE1t sinh trong t\u1EEBng b\u01B0\u1EDBc thi\u1EBFt k\u1EBF v\xE0 v\u1EADn h\xE0nh AI. C."
    },
    {
      "code": "11.C2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1ch AI h\u1ED7 tr\u1EE3 qu\xE1 tr\xECnh h\u1ECDc t\u1EADp v\xE0 thi\u1EBFt k\u1EBF c\xE1c c\xF4ng c\u1EE5 h\u1ED7 tr\u1EE3."
    },
    {
      "code": "11.C2.2",
      "mr": false,
      "desc": "\u0110\u1EC1 xu\u1EA5t \u0111\u01B0\u1EE3c c\xE1c t\xEDnh n\u0103ng AI h\u1ED7 tr\u1EE3 ho\u1EA1t \u0111\u1ED9ng h\u1ECDc t\u1EADp."
    },
    {
      "code": "11.C2.MR1",
      "mr": true,
      "desc": "S\u1EED d\u1EE5ng \u0111\u01B0\u1EE3c c\xF4ng c\u1EE5 AI \u0111\u1EC3 t\u1EA1o v\xE0 bi\xEAn t\u1EADp n\u1ED9i dung h\u1ECDc li\u1EC7u ph\u1EE5c v\u1EE5 h\u1ECDc t\u1EADp v\xE0 \u0111\xE1nh gi\xE1 s\u1EA3n ph\u1EA9m h\u1ECDc t\u1EADp."
    },
    {
      "code": "11.C3.1",
      "mr": false,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 k\u0129 thu\u1EADt prompt n\xE2ng cao t\u1EEB c\xE1c y\xEAu c\u1EA7u c\u1EE5 th\u1EC3 nh\u01B0: r\xE0ng bu\u1ED9c \u0111\u1ECBnh d\u1EA1ng \u0111\u1EA7u ra, chia nh\u1ECF nhi\u1EC7m v\u1EE5."
    },
    {
      "code": "11.C3.MR1",
      "mr": true,
      "desc": "V\u1EADn d\u1EE5ng \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 k\u0129 thu\u1EADt prompt n\xE2ng cao nh\u01B0: r\xE0ng bu\u1ED9c \u0111\u1ECBnh d\u1EA1ng \u0111\u1EA7u ra, chia nh\u1ECF nhi\u1EC7m v\u1EE5."
    },
    {
      "code": "11.C3.2",
      "mr": false,
      "desc": "M\xF4 t\u1EA3 \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xF4ng ngh\u1EC7 AI c\u01A1 b\u1EA3n nh\u01B0: chatbot, x\u1EED l\xED ng\xF4n ng\u1EEF t\u1EF1 nhi\xEAn, th\u1ECB gi\xE1c m\xE1y t\xEDnh, c\u1EA3m bi\u1EBFn."
    },
    {
      "code": "11.C3.MR2",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c c\xE1ch c\xE1c c\xF4ng ngh\u1EC7 \u0111\xF3 v\u1EADn h\xE0nh trong h\u1EC7 th\u1ED1ng AI."
    },
    {
      "code": "11.C3.MR3",
      "mr": true,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 ph\u01B0\u01A1ng ph\xE1p \u0111\u1EC3 t\xF9y ch\u1EC9nh h\u1EC7 th\u1ED1ng AI nh\u01B0: b\u1ED5 sung v\xE0 \u0111i\u1EC1u ch\u1EC9nh d\u1EEF li\u1EC7u, \u0111i\u1EC1u ch\u1EC9nh c\xE1c tham s\u1ED1 \u0111\u01B0\u1EE3c ph\xE9p, cung c\u1EA5p h\u01B0\u1EDBng d\u1EABn h\u1EC7 th\u1ED1ng, s\u1EED d\u1EE5ng k\u0129 thu\u1EADt sinh n\u1ED9i dung t\u0103ng c\u01B0\u1EDDng b\u1EB1ng truy xu\u1EA5t (RAG), \u2026"
    },
    {
      "code": "11.C3.MR4",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c \u1EDF m\u1EE9c kh\xE1i ni\u1EC7m c\xE1ch ho\u1EA1t \u0111\u1ED9ng c\u1EE7a k\u0129 thu\u1EADt sinh n\u1ED9i dung t\u0103ng c\u01B0\u1EDDng b\u1EB1ng truy xu\u1EA5t (RAG) v\xE0 v\xEC sao n\xF3 gi\xFAp gi\u1EA3m sai l\u1EC7ch th\xF4ng tin."
    },
    {
      "code": "11.C5.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng m\u1EA1ng n\u01A1-ron nh\xE2n t\u1EA1o."
    },
    {
      "code": "11.C5.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c ki\u1EBFn th\u1EE9c c\u01A1 b\u1EA3n v\u1EC1 m\u1EA1ng n\u01A1-ron nh\xE2n t\u1EA1o."
    },
    {
      "code": "11.C5.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 \u1EE9ng d\u1EE5ng thu\u1EADt to\xE1n ph\xE2n c\u1EE5m, ph\xE2n l\u1EDBp."
    },
    {
      "code": "11.C5.MR2",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c ki\u1EBFn th\u1EE9c c\u01A1 b\u1EA3n v\u1EC1 c\xE1c thu\u1EADt to\xE1n ph\xE2n c\u1EE5m, ph\xE2n l\u1EDBp v\xE0 m\u1ED9t s\u1ED1 \xFD t\u01B0\u1EDFng th\u1EF1c hi\u1EC7n. D."
    },
    {
      "code": "11.D1.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1ch th\u1EE9c thi\u1EBFt k\u1EBF v\xE0 v\u1EADn h\xE0nh t\u1ED5ng th\u1EC3 c\u1EE7a m\u1ED9t h\u1EC7 th\u1ED1ng AI, th\u1EC3 hi\u1EC7n \u0111\u01B0\u1EE3c m\u1ED1i quan h\u1EC7 gi\u1EEFa m\u1EE5c ti\xEAu, d\u1EEF li\u1EC7u v\xE0 c\xE1c th\xE0nh ph\u1EA7n trong h\u1EC7 th\u1ED1ng."
    },
    {
      "code": "11.D2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1ch th\u1EE9c v\u1EADn h\xE0nh c\u1EE7a c\xF4ng ngh\u1EC7 trong h\u1EC7 th\u1ED1ng AI, th\u1EC3 hi\u1EC7n \u0111\u01B0\u1EE3c m\u1ED1i li\xEAn h\u1EC7 gi\u1EEFa c\xE1c th\xE0nh ph\u1EA7n c\u1EE7a h\u1EC7 th\u1ED1ng trong vi\u1EC7c th\u1EF1c hi\u1EC7n m\u1ED9t nhi\u1EC7m v\u1EE5 c\u1EE5 th\u1EC3."
    },
    {
      "code": "11.D2.MR1",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c c\xE1c c\xE1ch th\u1EE9c gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1 ph\xE1t sinh c\u1EE7a h\u1EC7 th\u1ED1ng AI nh\u1EB1m t\u1ED1i \u01B0u ho\xE1 hi\u1EC7u qu\u1EA3 ho\u1EA1t \u0111\u1ED9ng c\u1EE7a h\u1EC7 th\u1ED1ng."
    }
  ],
  "12": [
    {
      "code": "12.A1.1",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c m\u1ED9t h\u1EC7 th\u1ED1ng AI nh\u1EB1m \u0111\u1EA3m b\u1EA3o con ng\u01B0\u1EDDi c\xF3 quy\u1EC1n ki\u1EC3m so\xE1t v\xE0 ch\u1ECBu tr\xE1ch nhi\u1EC7m \u0111\u1ED1i v\u1EDBi t\u1EA5t c\u1EA3 c\xE1c b\u01B0\u1EDBc quan tr\u1ECDng trong v\xF2ng \u0111\u1EDDi AI."
    },
    {
      "code": "12.A1.MR1",
      "mr": true,
      "desc": "Th\u1EF1c hi\u1EC7n \u0111\u01B0\u1EE3c vi\u1EC7c ph\xE2n t\xEDch quy\u1EC1n ki\u1EC3m so\xE1t v\xE0 tr\xE1ch nhi\u1EC7m c\u1EE7a con ng\u01B0\u1EDDi trong v\xF2ng \u0111\u1EDDi AI th\xF4ng qua m\u1ED9t d\u1EF1 \xE1n s\xE1ng t\u1EA1o AI."
    },
    {
      "code": "12.A1.2",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c vai tr\xF2 c\u1EE7a con ng\u01B0\u1EDDi v\xE0 AI trong c\xE1c b\u01B0\u1EDBc ch\xEDnh c\u1EE7a qu\xE1 tr\xECnh ra quy\u1EBFt \u0111\u1ECBnh."
    },
    {
      "code": "12.A1.3",
      "mr": false,
      "desc": "Ki\u1EC3m tra \u0111\u01B0\u1EE3c vi\u1EC7c th\u1EF1c hi\u1EC7n tr\xE1ch nhi\u1EC7m gi\u1EA3i tr\xECnh c\u1EE7a con ng\u01B0\u1EDDi \u0111\u1ED1i v\u1EDBi c\xE1c quy\u1EBFt \u0111\u1ECBnh, \u0111\u1ED1i chi\u1EBFu v\u1EDBi c\xE1c quy \u0111\u1ECBnh hi\u1EC7n h\xE0nh trong n\u01B0\u1EDBc v\xE0 qu\u1ED1c t\u1EBF."
    },
    {
      "code": "12.A2.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 nguy\xEAn t\u1EAFc \u0111\u1EA1o \u0111\u1EE9c c\u01A1 b\u1EA3n trong thi\u1EBFt k\u1EBF, ph\xE1t tri\u1EC3n s\u1EA3n ph\u1EA9m AI, g\u1ED3m c\xF3: an to\xE0n \u2013 kh\xF4ng g\xE2y nguy hi\u1EC3m cho con ng\u01B0\u1EDDi v\xE0 m\xF4i tr\u01B0\u1EDDng; c\xF4ng b\u1EB1ng \u2013 tr\xE1nh thi\xEAn v\u1ECB, ph\xE2n bi\u1EC7t \u0111\u1ED1i x\u1EED trong d\u1EEF li\u1EC7u v\xE0 k\u1EBFt qu\u1EA3; minh b\u1EA1ch \u2013 c\xF4ng khai c\xE1ch ho\u1EA1t \u0111\u1ED9ng, kh\xF4ng che gi\u1EA5u th\xF4ng tin; t\xF4n tr\u1ECDng quy\u1EC1n ri\xEAng t\u01B0 \u2013 kh\xF4ng s\u1EED d\u1EE5ng, chia s\u1EBB d\u1EEF li\u1EC7u c\xE1 nh\xE2n tr\xE1i ph\xE9p; tr\xE1ch nhi\u1EC7m \u2013 ki\u1EC3m tra, \u0111\xE1nh gi\xE1 v\xE0 ch\u1ECBu tr\xE1ch nhi\u1EC7m v\u1EC1 s\u1EA3n ph\u1EA9m t\u1EA1o ra; l\u1EE3i \xEDch x\xE3 h\u1ED9i \u2013 ph\u1EE5c v\u1EE5 con ng\u01B0\u1EDDi, c\u1ED9ng \u0111\u1ED3ng v\xE0 ph\xE1t tri\u1EC3n b\u1EC1n v\u1EEFng."
    },
    {
      "code": "12.A2.MR1",
      "mr": true,
      "desc": "V\u1EADn d\u1EE5ng \u0111\u01B0\u1EE3c c\xE1c nguy\xEAn t\u1EAFc \u0111\u1EC3 so\u1EA1n th\u1EA3o \u0111\u01B0\u1EE3c b\u1ED9 nguy\xEAn t\u1EAFc c\xE1 nh\xE2n cho m\u1ED9t d\u1EF1 \xE1n thi\u1EBFt k\u1EBF, ph\xE1t tri\u1EC3n s\u1EA3n ph\u1EA9m AI c\u1EE5 th\u1EC3; \u0111\u1ED1i chi\u1EBFu c\xE1c quy\u1EBFt \u0111\u1ECBnh thi\u1EBFt k\u1EBF (m\u1EE5c ti\xEAu, d\u1EEF li\u1EC7u, t\xEDnh n\u0103ng, ki\u1EC3m th\u1EED) v\u1EDBi b\u1ED9 nguy\xEAn t\u1EAFc \u0111\xF3 v\xE0 \u0111i\u1EC1u ch\u1EC9nh khi ph\xE1t hi\u1EC7n nguy c\u01A1 vi ph\u1EA1m."
    },
    {
      "code": "12.A3.1",
      "mr": false,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c n\u1ED9i h\xE0m c\u1EE7a \u201Ctr\xE1ch nhi\u1EC7m c\xF4ng d\xE2n trong x\xE3 h\u1ED9i AI\u201D: s\u1EED d\u1EE5ng AI m\u1ED9t c\xE1ch an to\xE0n, trung th\u1EF1c v\xE0 c\xF3 \u0111\u1EA1o \u0111\u1EE9c; t\xF4n tr\u1ECDng quy\u1EC1n ri\xEAng t\u01B0 v\xE0 d\u1EEF li\u1EC7u c\u1EE7a ng\u01B0\u1EDDi kh\xE1c; kh\xF4ng s\u1EED d\u1EE5ng AI \u0111\u1EC3 lan truy\u1EC1n th\xF4ng tin sai l\u1EC7ch, gian l\u1EADn hay g\xE2y h\u1EA1i; \u0111\xF3ng g\xF3p v\xE0o vi\u1EC7c x\xE2y d\u1EF1ng m\xF4i tr\u01B0\u1EDDng s\u1ED1 t\xEDch c\u1EF1c, c\xF4ng b\u1EB1ng v\xE0 nh\xE2n v\u0103n. B."
    },
    {
      "code": "12.B1.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c nguy\xEAn nh\xE2n d\u1EABn \u0111\u1EBFn c\xE1c v\u1EA5n \u0111\u1EC1 \u0111\u1EA1o \u0111\u1EE9c ho\u1EB7c sai l\u1EC7ch trong qu\xE1 tr\xECnh ho\u1EA1t \u0111\u1ED9ng c\u1EE7a h\u1EC7 th\u1ED1ng AI."
    },
    {
      "code": "12.B2.1",
      "mr": false,
      "desc": "X\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c m\u1EE9c \u0111\u1ED9 r\u1EE7i ro khi s\u1EED d\u1EE5ng AI c\xF3 th\u1EC3 d\u1EABn \u0111\u1EBFn vi ph\u1EA1m quy \u0111\u1ECBnh c\u1EE7a nh\xE0 tr\u01B0\u1EDDng ho\u1EB7c c\xE1c quy \u0111\u1ECBnh ph\xE1p lu\u1EADt li\xEAn quan."
    },
    {
      "code": "12.B3.1",
      "mr": false,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c quy\u1EC1n v\xE0 tr\xE1ch nhi\u1EC7m c\u1EE7a ng\u01B0\u1EDDi ph\xE1t tri\u1EC3n, ng\u01B0\u1EDDi s\u1EED d\u1EE5ng AI, c\u0169ng nh\u01B0 vai tr\xF2 c\u1EE7a c\xE1 nh\xE2n trong vi\u1EC7c g\xF3p \xFD, \u0111\u1EC1 xu\u1EA5t x\xE2y d\u1EF1ng ch\xEDnh s\xE1ch v\xE0 quy \u0111\u1ECBnh li\xEAn quan \u0111\u1EBFn AI. C."
    },
    {
      "code": "12.C2.1",
      "mr": false,
      "desc": "L\u1EF1a ch\u1ECDn \u0111\u01B0\u1EE3c \xFD t\u01B0\u1EDFng thi\u1EBFt k\u1EBF m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 AI \u0111\u1EC3 th\u1EF1c hi\u1EC7n c\xE1c c\xF4ng vi\u1EC7c kh\xE1c nhau."
    },
    {
      "code": "12.C2.MR1",
      "mr": true,
      "desc": "T\xF9y ch\u1EC9nh \u0111\u01B0\u1EE3c c\xE1c y\xEAu c\u1EA7u h\u1EC7 th\u1ED1ng AI \u0111\u1EC3 h\u1ED7 tr\u1EE3 c\xE1c ho\u1EA1t \u0111\u1ED9ng h\u1ECDc t\u1EADp v\xE0 ho\u1EA1t \u0111\u1ED9ng x\xE3 h\u1ED9i."
    },
    {
      "code": "12.C3.1",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 m\xE3 ngu\u1ED3n m\u1EDF ho\u1EB7c mi\u1EC5n ph\xED d\xF9ng \u0111\u1EC3 thi\u1EBFt k\u1EBF, hu\u1EA5n luy\u1EC7n v\xE0 ph\xE1t tri\u1EC3n h\u1EC7 th\u1ED1ng AI, nh\u01B0: Teachable Machine, ML5.js, TensorFlow.js, MIT App Inventor ho\u1EB7c c\xF4ng c\u1EE5 ph\xF9 h\u1EE3p kh\xE1c."
    },
    {
      "code": "12.C3.MR1",
      "mr": true,
      "desc": "S\u1EED d\u1EE5ng \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 c\xF4ng c\u1EE5 m\xE3 ngu\u1ED3n m\u1EDF ho\u1EB7c mi\u1EC5n ph\xED d\xF9ng \u0111\u1EC3 thi\u1EBFt k\u1EBF, hu\u1EA5n luy\u1EC7n v\xE0 ph\xE1t tri\u1EC3n h\u1EC7 th\u1ED1ng AI."
    },
    {
      "code": "12.C3.2",
      "mr": false,
      "desc": "N\xEAu \u0111\u01B0\u1EE3c v\xED d\u1EE5 v\u1EC1 c\xE1ch th\u1EE9c \u0111\xE1nh gi\xE1 hi\u1EC7u qu\u1EA3 c\u1EE7a h\u1EC7 th\u1ED1ng AI."
    },
    {
      "code": "12.C3.MR2",
      "mr": true,
      "desc": "\u0110\xE1nh gi\xE1 \u0111\u01B0\u1EE3c kh\u1EA3 n\u0103ng t\u1ED1i \u01B0u h\u1EC7 th\u1ED1ng AI th\xF4ng qua c\u1EADp nh\u1EADt c\xF4ng ngh\u1EC7, k\u0129 thu\u1EADt m\u1EDBi."
    },
    {
      "code": "12.C3.MR3",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 kh\xE1i ni\u1EC7m c\u01A1 b\u1EA3n c\u1EE7a h\u1EC7 th\u1ED1ng \u1EE9ng d\u1EE5ng h\u1ECDc m\xE1y nh\u01B0: h\xE0m m\u1EE5c ti\xEAu, t\u1ED1i \u01B0u ho\xE1 h\u1EC7 th\u1ED1ng, m\xF4 h\xECnh qu\xE1 kh\u1EDBp d\u1EEF li\u1EC7u (overfitting)."
    },
    {
      "code": "12.C4.MR1",
      "mr": true,
      "desc": "Thu th\u1EADp v\xE0 t\u1ED5 ch\u1EE9c \u0111\u01B0\u1EE3c d\u1EEF li\u1EC7u \u0111\xE1p \u1EE9ng y\xEAu c\u1EA7u c\u1EE7a vi\u1EC7c ph\xE1t tri\u1EC3n h\u1EC7 th\u1ED1ng AI."
    },
    {
      "code": "12.C4.MR2",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch v\xE0 x\xE1c \u0111\u1ECBnh \u0111\u01B0\u1EE3c c\xE1c n\u1EC1n t\u1EA3ng ho\u1EB7c b\u1ED9 c\xF4ng c\u1EE5 ph\xE1t tri\u1EC3n AI, c\u1EA3i thi\u1EC7n c\xE1c b\u1ED9 d\u1EEF li\u1EC7u \u0111\xE1p \u1EE9ng qu\xE1 tr\xECnh thi\u1EBFt k\u1EBF, ph\xE1t tri\u1EC3n AI. D."
    },
    {
      "code": "12.D1.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 ph\u01B0\u01A1ng \xE1n thi\u1EBFt k\u1EBF v\xE0 v\u1EADn h\xE0nh h\u1EC7 th\u1ED1ng AI ph\xF9 h\u1EE3p \u0111\u1EC3 \u0111\u1EA1t hi\u1EC7u qu\u1EA3 cao trong m\u1ED9t s\u1ED1 nhi\u1EC7m v\u1EE5 c\u1EE5 th\u1EC3."
    },
    {
      "code": "12.D1.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c m\u1ED9t s\u1ED1 ph\u01B0\u01A1ng \xE1n thi\u1EBFt k\u1EBF v\xE0 v\u1EADn h\xE0nh h\u1EC7 th\u1ED1ng AI ph\xF9 h\u1EE3p \u0111\u1EC3 \u0111\u1EA1t hi\u1EC7u qu\u1EA3 cao trong m\u1ED9t s\u1ED1 nhi\u1EC7m v\u1EE5 c\u1EE5 th\u1EC3."
    },
    {
      "code": "12.D2.1",
      "mr": false,
      "desc": "Nh\u1EADn bi\u1EBFt \u0111\u01B0\u1EE3c c\xE1c vai tr\xF2 kh\xE1c nhau trong qu\xE1 tr\xECnh ph\xE1t tri\u1EC3n m\u1ED9t s\u1EA3n ph\u1EA9m AI (nh\u01B0 ng\u01B0\u1EDDi \u0111\u1EC1 xu\u1EA5t \xFD t\u01B0\u1EDFng, l\u1EADp tr\xECnh, hu\u1EA5n luy\u1EC7n, ki\u1EC3m th\u1EED) v\xE0 vi\u1EC7c t\u1EA1o ra s\u1EA3n ph\u1EA9m AI c\u1EA7n c\xF3 s\u1EF1 h\u1EE3p t\xE1c gi\u1EEFa nhi\u1EC1u ng\u01B0\u1EDDi v\u1EDBi chuy\xEAn m\xF4n kh\xE1c nhau."
    },
    {
      "code": "12.D2.MR1",
      "mr": true,
      "desc": "Ph\xE2n t\xEDch \u0111\u01B0\u1EE3c nguy\xEAn nh\xE2n c\u1EE7a c\xE1c v\u1EA5n \u0111\u1EC1 ph\xE1t sinh trong h\u1EC7 th\u1ED1ng AI v\xE0 l\u1EF1a ch\u1ECDn \u0111\u01B0\u1EE3c c\xE1ch gi\u1EA3i quy\u1EBFt ph\xF9 h\u1EE3p \u0111\u1EC3 h\u1EC7 th\u1ED1ng ho\u1EA1t \u0111\u1ED9ng \u1ED5n \u0111\u1ECBnh v\xE0 hi\u1EC7u qu\u1EA3 h\u01A1n."
    },
    {
      "code": "12.D2.MR2",
      "mr": true,
      "desc": "Tr\xECnh b\xE0y \u0111\u01B0\u1EE3c kh\u1EA3 n\u0103ng v\xE0 c\u1EA5u tr\xFAc c\u01A1 b\u1EA3n c\u1EE7a m\u1ED9t h\u1EC7 th\u1ED1ng t\xE1c nh\xE2n AI (AI agent)."
    },
    {
      "code": "12.D2.MR3",
      "mr": true,
      "desc": "X\xE2y d\u1EF1ng v\xE0 ki\u1EC3m th\u1EED \u0111\u01B0\u1EE3c h\u1EC7 th\u1ED1ng t\xE1c nh\xE2n AI \u0111\u01A1n gi\u1EA3n ph\u1EE5c v\u1EE5 m\u1ED9t nhi\u1EC7m v\u1EE5 h\u1ECDc t\u1EADp ho\u1EB7c c\u1ED9ng \u0111\u1ED3ng."
    }
  ]
};
var NLS_TT02_COMPONENTS = [
  { id: "1.1", domain: "I. Khai th\xE1c d\u1EEF li\u1EC7u v\xE0 th\xF4ng tin", desc: "Duy\u1EC7t, t\xECm ki\u1EBFm v\xE0 l\u1ECDc d\u1EEF li\u1EC7u, th\xF4ng tin v\xE0 n\u1ED9i dung s\u1ED1." },
  { id: "1.2", domain: "I. Khai th\xE1c d\u1EEF li\u1EC7u v\xE0 th\xF4ng tin", desc: "\u0110\xE1nh gi\xE1 d\u1EEF li\u1EC7u, th\xF4ng tin v\xE0 n\u1ED9i dung s\u1ED1." },
  { id: "1.3", domain: "I. Khai th\xE1c d\u1EEF li\u1EC7u v\xE0 th\xF4ng tin", desc: "Qu\u1EA3n l\xFD d\u1EEF li\u1EC7u, th\xF4ng tin v\xE0 n\u1ED9i dung s\u1ED1." },
  { id: "2.1", domain: "II. Giao ti\u1EBFp v\xE0 h\u1EE3p t\xE1c trong m\xF4i tr\u01B0\u1EDDng s\u1ED1", desc: "T\u01B0\u01A1ng t\xE1c th\xF4ng qua c\xF4ng ngh\u1EC7 s\u1ED1." },
  { id: "2.2", domain: "II. Giao ti\u1EBFp v\xE0 h\u1EE3p t\xE1c trong m\xF4i tr\u01B0\u1EDDng s\u1ED1", desc: "Chia s\u1EBB th\xF4ng tin v\xE0 n\u1ED9i dung th\xF4ng qua c\xF4ng ngh\u1EC7 s\u1ED1." },
  { id: "2.3", domain: "II. Giao ti\u1EBFp v\xE0 h\u1EE3p t\xE1c trong m\xF4i tr\u01B0\u1EDDng s\u1ED1", desc: "S\u1EED d\u1EE5ng c\xF4ng ngh\u1EC7 s\u1ED1 \u0111\u1EC3 th\u1EF1c hi\u1EC7n tr\xE1ch nhi\u1EC7m c\xF4ng d\xE2n." },
  { id: "2.4", domain: "II. Giao ti\u1EBFp v\xE0 h\u1EE3p t\xE1c trong m\xF4i tr\u01B0\u1EDDng s\u1ED1", desc: "H\u1EE3p t\xE1c th\xF4ng qua c\xF4ng ngh\u1EC7 s\u1ED1." },
  { id: "2.5", domain: "II. Giao ti\u1EBFp v\xE0 h\u1EE3p t\xE1c trong m\xF4i tr\u01B0\u1EDDng s\u1ED1", desc: "Th\u1EF1c hi\u1EC7n quy t\u1EAFc \u1EE9ng x\u1EED tr\xEAn m\u1EA1ng." },
  { id: "2.6", domain: "II. Giao ti\u1EBFp v\xE0 h\u1EE3p t\xE1c trong m\xF4i tr\u01B0\u1EDDng s\u1ED1", desc: "Qu\u1EA3n l\xFD danh t\xEDnh s\u1ED1." },
  { id: "3.1", domain: "III. S\xE1ng t\u1EA1o n\u1ED9i dung s\u1ED1", desc: "Ph\xE1t tri\u1EC3n n\u1ED9i dung s\u1ED1." },
  { id: "3.2", domain: "III. S\xE1ng t\u1EA1o n\u1ED9i dung s\u1ED1", desc: "T\xEDch h\u1EE3p v\xE0 t\u1EA1o l\u1EADp l\u1EA1i n\u1ED9i dung s\u1ED1." },
  { id: "3.3", domain: "III. S\xE1ng t\u1EA1o n\u1ED9i dung s\u1ED1", desc: "Th\u1EF1c thi b\u1EA3n quy\u1EC1n v\xE0 gi\u1EA5y ph\xE9p." },
  { id: "3.4", domain: "III. S\xE1ng t\u1EA1o n\u1ED9i dung s\u1ED1", desc: "L\u1EADp tr\xECnh." },
  { id: "4.1", domain: "IV. An to\xE0n", desc: "B\u1EA3o v\u1EC7 thi\u1EBFt b\u1ECB." },
  { id: "4.2", domain: "IV. An to\xE0n", desc: "B\u1EA3o v\u1EC7 d\u1EEF li\u1EC7u c\xE1 nh\xE2n v\xE0 quy\u1EC1n ri\xEAng t\u01B0." },
  { id: "4.3", domain: "IV. An to\xE0n", desc: "B\u1EA3o v\u1EC7 s\u1EE9c kh\u1ECFe v\xE0 an sinh s\u1ED1." },
  { id: "4.4", domain: "IV. An to\xE0n", desc: "B\u1EA3o v\u1EC7 m\xF4i tr\u01B0\u1EDDng." },
  { id: "5.1", domain: "V. Gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1", desc: "Gi\u1EA3i quy\u1EBFt c\xE1c v\u1EA5n \u0111\u1EC1 k\u1EF9 thu\u1EADt." },
  { id: "5.2", domain: "V. Gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1", desc: "X\xE1c \u0111\u1ECBnh nhu c\u1EA7u v\xE0 gi\u1EA3i ph\xE1p c\xF4ng ngh\u1EC7." },
  { id: "5.3", domain: "V. Gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1", desc: "S\u1EED d\u1EE5ng s\xE1ng t\u1EA1o c\xF4ng ngh\u1EC7 s\u1ED1." },
  { id: "5.4", domain: "V. Gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1", desc: "X\xE1c \u0111\u1ECBnh c\xE1c v\u1EA5n \u0111\u1EC1 c\u1EA7n c\u1EA3i thi\u1EC7n v\u1EC1 n\u0103ng l\u1EF1c s\u1ED1." },
  { id: "6.1", domain: "VI. \u1EE8ng d\u1EE5ng tr\xED tu\u1EC7 nh\xE2n t\u1EA1o", desc: "Hi\u1EC3u bi\u1EBFt v\u1EC1 AI (trong \u0111\xF3 c\xF3 AI t\u1EA1o sinh)." },
  { id: "6.2", domain: "VI. \u1EE8ng d\u1EE5ng tr\xED tu\u1EC7 nh\xE2n t\u1EA1o", desc: "S\u1EED d\u1EE5ng AI c\xF3 \u0111\u1EA1o \u0111\u1EE9c v\xE0 tr\xE1ch nhi\u1EC7m." },
  { id: "6.3", domain: "VI. \u1EE8ng d\u1EE5ng tr\xED tu\u1EC7 nh\xE2n t\u1EA1o", desc: "\u0110\xE1nh gi\xE1 c\xE1c c\xF4ng c\u1EE5 AI." }
];
var ATM_CATALOG = [
  {
    code: "ATM1",
    title: "Khai th\xE1c, \u0111\xE1nh gi\xE1 th\xF4ng tin s\u1ED1 v\xE0 nh\u1EADn di\u1EC7n r\u1EE7i ro th\xF4ng tin",
    requiredNls: ["1.1", "1.2"],
    recommendedNls: ["1.3", "2.5", "4.1", "4.2", "6.3"],
    desc: "T\xECm ki\u1EBFm, \u0111\xE1nh gi\xE1, ki\u1EC3m ch\u1EE9ng ngu\u1ED3n tin; nh\u1EADn di\u1EC7n tin gi\u1EA3, qu\u1EA3ng c\xE1o tr\xE1 h\xECnh, thao t\xFAng c\u1EA3m x\xFAc, link l\u1EA1, m\xE3 QR gi\u1EA3, tin nh\u1EAFn tr\xFAng th\u01B0\u1EDFng, gi\u1EA3 danh, l\u1EEBa \u0111\u1EA3o tr\u1EF1c tuy\u1EBFn."
  },
  {
    code: "ATM2",
    title: "Giao ti\u1EBFp, h\u1EE3p t\xE1c, \u1EE9ng x\u1EED v\xE0 c\xF4ng d\xE2n s\u1ED1",
    requiredNls: ["2.1", "2.5"],
    recommendedNls: ["2.2", "2.3", "2.4", "4.3"],
    desc: "\u1EE8ng x\u1EED v\u0103n minh trong nh\xF3m chat, m\u1EA1ng x\xE3 h\u1ED9i, tr\xF2 ch\u01A1i tr\u1EF1c tuy\u1EBFn; t\xF4n tr\u1ECDng danh d\u1EF1, nh\xE2n ph\u1EA9m, quy\u1EC1n ri\xEAng t\u01B0; ph\xF2ng ch\u1ED1ng b\u1EAFt n\u1EA1t tr\xEAn m\xF4i tr\u01B0\u1EDDng m\u1EA1ng, d\u1EE5 d\u1ED7 tr\u1EF1c tuy\u1EBFn v\xE0 h\xE0nh vi lan truy\u1EC1n n\u1ED9i dung g\xE2y h\u1EA1i"
  },
  {
    code: "ATM3",
    title: "Qu\u1EA3n l\xFD danh t\xEDnh s\u1ED1, d\u1EEF li\u1EC7u c\xE1 nh\xE2n v\xE0 quy\u1EC1n ri\xEAng t\u01B0",
    requiredNls: ["2.6"],
    recommendedNls: ["1.3", "2.2", "4.1", "4.2", "5.2"],
    desc: "Nh\u1EADn bi\u1EBFt d\u1EEF li\u1EC7u c\xE1 nh\xE2n, d\u1EEF li\u1EC7u nh\u1EA1y c\u1EA3m; qu\u1EA3n l\xFD h\u1ED3 s\u01A1, h\xECnh \u1EA3nh, d\u1EA5u v\u1EBFt s\u1ED1; xin ph\xE9p tr\u01B0\u1EDBc khi chia s\u1EBB h\xECnh \u1EA3nh, th\xF4ng tin c\u1EE7a ng\u01B0\u1EDDi kh\xE1c; ki\u1EC3m so\xE1t quy\u1EC1n truy c\u1EADp \u1EE9ng d\u1EE5ng."
  },
  {
    code: "ATM4",
    title: "An to\xE0n thi\u1EBFt b\u1ECB, t\xE0i kho\u1EA3n v\xE0 \u1EE9ng ph\xF3 s\u1EF1 c\u1ED1",
    requiredNls: ["4.1", "4.2"],
    recommendedNls: ["2.6", "5.1", "5.2"],
    desc: "M\u1EADt kh\u1EA9u m\u1EA1nh, x\xE1c th\u1EF1c nhi\u1EC1u l\u1EDBp, c\u1EADp nh\u1EADt ph\u1EA7n m\u1EC1m, nh\u1EADn di\u1EC7n m\xE3 \u0111\u1ED9c, b\u1EA3o v\u1EC7 thi\u1EBFt b\u1ECB d\xF9ng chung; l\u01B0u b\u1EB1ng ch\u1EE9ng, ch\u1EB7n, b\xE1o c\xE1o, t\xECm ki\u1EBFm h\u1ED7 tr\u1EE3 khi b\u1ECB l\u1EEBa \u0111\u1EA3o, chi\u1EBFm t\xE0i kho\u1EA3n, qu\u1EA5y r\u1ED1i, ph\xE1t t\xE1n th\xF4ng tin."
  },
  {
    code: "ATM5",
    title: "S\u1EE9c kh\u1ECFe s\u1ED1, an sinh s\u1ED1 v\xE0 c\xE2n b\u1EB1ng \u0111\u1EDDi s\u1ED1ng s\u1ED1",
    requiredNls: ["4.3"],
    recommendedNls: ["4.4", "5.2", "5.4"],
    desc: "Qu\u1EA3n l\xFD th\u1EDDi gian s\u1EED d\u1EE5ng thi\u1EBFt b\u1ECB, gi\u1EA5c ng\u1EE7, th\u1ECB l\u1EF1c, t\u01B0 th\u1EBF, v\u1EADn \u0111\u1ED9ng, c\u1EA3m x\xFAc, \xE1p l\u1EF1c m\u1EA1ng x\xE3 h\u1ED9i, l\u1EC7 thu\u1ED9c tr\xF2 ch\u01A1i tr\u1EF1c tuy\u1EBFn, m\u1EA1ng x\xE3 h\u1ED9i v\xE0 h\xECnh th\xE0nh th\xF3i quen s\u1ED1 l\xE0nh m\u1EA1nh."
  },
  {
    code: "ATM6",
    title: "S\xE1ng t\u1EA1o n\u1ED9i dung s\u1ED1, b\u1EA3n quy\u1EC1n v\xE0 tr\xE1ch nhi\u1EC7m s\u1ED1",
    requiredNls: ["2.2", "3.1", "3.2", "3.3"],
    recommendedNls: ["2.5"],
    desc: "T\u1EA1o l\u1EADp, chia s\u1EBB c\xE1c s\u1EA3n ph\u1EA9m s\u1ED1 an to\xE0n; t\xF4n tr\u1ECDng b\u1EA3n quy\u1EC1n, tr\xEDch d\u1EABn ngu\u1ED3n, kh\xF4ng sao ch\xE9p, kh\xF4ng ph\xE1t t\xE1n n\u1ED9i dung vi ph\u1EA1m ph\xE1p lu\u1EADt ho\u1EB7c g\xE2y t\u1ED5n h\u1EA1i \u0111\u1EBFn b\u1EA3n th\xE2n v\xE0 ng\u01B0\u1EDDi kh\xE1c."
  },
  {
    code: "ATM7",
    title: "S\u1EED d\u1EE5ng AI an to\xE0n trong h\u1ECDc t\u1EADp v\xE0 sinh ho\u1EA1t",
    requiredNls: ["6.1", "6.2"],
    recommendedNls: ["1.2", "3.3", "4.2", "6.3"],
    desc: "Hi\u1EC3u \u0111\u01B0\u1EE3c AI c\xF3 th\u1EC3 sai, thi\xEAn l\u1EC7ch ho\u1EB7c t\u1EA1o n\u1ED9i dung gi\u1EA3 m\u1EA1o; kh\xF4ng cung c\u1EA5p d\u1EEF li\u1EC7u c\xE1 nh\xE2n/d\u1EEF li\u1EC7u h\u1ECDc sinh v\xE0o c\xF4ng c\u1EE5 AI khi ch\u01B0a \u0111\u01B0\u1EE3c ph\xE9p; s\u1EED d\u1EE5ng AI minh b\u1EA1ch, trung th\u1EF1c, c\xF3 \u0111\u1EA1o \u0111\u1EE9c, c\xF3 ki\u1EC3m ch\u1EE9ng, kh\xF4ng ph\u1EE5 thu\u1ED9c v\xE0 kh\xF4ng d\xF9ng AI \u0111\u1EC3 gian l\u1EADn ho\u1EB7c g\xE2y h\u1EA1i."
  },
  {
    code: "ATM8",
    title: "T\xECm ki\u1EBFm h\u1ED7 tr\u1EE3, ph\u1EE5c h\u1ED3i v\xE0 \u0111\xF3ng g\xF3p c\u1ED9ng \u0111\u1ED3ng s\u1ED1",
    requiredNls: ["2.3", "4.3"],
    recommendedNls: ["2.2", "2.5"],
    desc: "Bi\u1EBFt cung c\u1EA5p th\xF4ng tin v\u1EDBi ng\u01B0\u1EDDi tin c\u1EADy, gi\xE1o vi\xEAn, c\xE1n b\u1ED9 t\u01B0 v\u1EA5n, t\u1ED5ng \u0111\xE0i ho\u1EB7c c\u01A1 quan ch\u1EE9c n\u0103ng; bi\u1EBFt h\u1ED7 tr\u1EE3 b\u1EA1n b\xE8 \u0111\xFAng c\xE1ch; g\xF3p ph\u1EA7n x\xE2y d\u1EF1ng m\xF4i tr\u01B0\u1EDDng s\u1ED1 t\xEDch c\u1EF1c, nh\xE2n v\u0103n trong l\u1EDBp h\u1ECDc v\xE0 nh\xE0 tr\u01B0\u1EDDng, gia \u0111\xECnh v\xE0 x\xE3 h\u1ED9i."
  }
];

// tich-hop-giao-an-standalone/app.js
var { useState, useEffect, useMemo } = React;
var SUBJECTS = [
  "Tin h\u1ECDc",
  "To\xE1n",
  "Ng\u1EEF v\u0103n",
  "Ti\u1EBFng Anh",
  "Khoa h\u1ECDc t\u1EF1 nhi\xEAn",
  "L\u1ECBch s\u1EED v\xE0 \u0110\u1ECBa l\xED",
  "Gi\xE1o d\u1EE5c c\xF4ng d\xE2n",
  "C\xF4ng ngh\u1EC7",
  "Gi\xE1o d\u1EE5c th\u1EC3 ch\u1EA5t",
  "Ngh\u1EC7 thu\u1EADt",
  "\xC2m nh\u1EA1c",
  "M\u0129 thu\u1EADt",
  "Ho\u1EA1t \u0111\u1ED9ng tr\u1EA3i nghi\u1EC7m",
  "V\u1EADt l\xED",
  "Ho\xE1 h\u1ECDc",
  "Sinh h\u1ECDc",
  "L\u1ECBch s\u1EED",
  "\u0110\u1ECBa l\xED",
  "Gi\xE1o d\u1EE5c Kinh t\u1EBF v\xE0 Ph\xE1p lu\u1EADt",
  "Gi\xE1o d\u1EE5c qu\u1ED1c ph\xF2ng v\xE0 an ninh",
  "Ti\u1EBFng Vi\u1EC7t",
  "T\u1EF1 nhi\xEAn v\xE0 X\xE3 h\u1ED9i",
  "Khoa h\u1ECDc"
];
var DISABILITY_TYPES = [
  { key: "chung", label: "Khuy\u1EBFt t\u1EADt chung" },
  { key: "nhin", label: "Khuy\u1EBFt t\u1EADt v\u1EC1 m\u1EAFt (nh\xECn)" },
  { key: "vandong", label: "Khuy\u1EBFt t\u1EADt v\u1EC1 x\u01B0\u01A1ng (v\u1EADn \u0111\u1ED9ng)" },
  { key: "tritue", label: "Khuy\u1EBFt t\u1EADt v\u1EC1 tr\xED tu\u1EC7" },
  { key: "nghenoi", label: "Khuy\u1EBFt t\u1EADt nghe, n\xF3i" },
  { key: "tuky", label: "R\u1ED1i lo\u1EA1n ph\u1ED5 t\u1EF1 k\u1EF7" },
  { key: "rlht", label: "R\u1ED1i lo\u1EA1n h\u1ECDc t\u1EADp \u0111\u1EB7c th\xF9" },
  { key: "thankinh", label: "Khuy\u1EBFt t\u1EADt th\u1EA7n kinh" },
  { key: "tamthan", label: "Khuy\u1EBFt t\u1EADt t\xE2m th\u1EA7n" },
  { key: "khac", label: "Khuy\u1EBFt t\u1EADt kh\xE1c" }
];
function getNlsLevelLabel(grade) {
  if (grade <= 3) return "m\u1EE9c CB1 \u2014 l\u1EDBp 1\u20133";
  if (grade <= 5) return "m\u1EE9c CB2 \u2014 l\u1EDBp 4\u20135";
  if (grade <= 7) return "m\u1EE9c TC1 \u2014 l\u1EDBp 6\u20137";
  if (grade <= 9) return "m\u1EE9c TC2 \u2014 l\u1EDBp 8\u20139";
  return "m\u1EE9c NC1 \u2014 l\u1EDBp 10\u201312";
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
  return NLS_TT02_COMPONENTS.map((c) => ({
    code: `${c.id}.${lvl}a`,
    desc: `${c.domain}: ${c.desc}`
  }));
}
function getAiCatalog(grade) {
  return AI_CATALOG_BY_GRADE[Number(grade)] || AI_CATALOG_BY_GRADE[10];
}
var GDQPAN_SUGGESTIONS = [
  "Li\xEAn h\u1EC7 b\xE0i to\xE1n s\u1ED1 li\u1EC7u b\u1EA3o v\u1EC7 ch\u1EE7 quy\u1EC1n bi\xEAn gi\u1EDBi v\xE0 bi\u1EC3n \u0111\u1EA3o T\u1ED5 qu\u1ED1c",
  "Gi\xE1o d\u1EE5c \xFD th\u1EE9c an to\xE0n th\xF4ng tin c\xE1 nh\xE2n v\xE0 an ninh m\u1EA1ng",
  "L\u1ED3ng gh\xE9p truy\u1EC1n th\u1ED1ng v\u1EBB vang, g\u01B0\u01A1ng anh h\xF9ng li\u1EC7t s\u0129 Q\u0110ND v\xE0 CAND",
  "Nh\u1EAFc nh\u1EDF tinh th\u1EA7n k\u1EF7 lu\u1EADt, \u0111o\xE0n k\u1EBFt v\xE0 tr\xE1ch nhi\u1EC7m c\u1EE7a h\u1ECDc sinh v\u1EDBi T\u1ED5 qu\u1ED1c"
];
var BDKH_SUGGESTIONS = [
  "Li\xEAn h\u1EC7 b\xE0i to\xE1n s\u1ED1 li\u1EC7u v\u1EC1 l\u01B0\u1EE3ng r\xE1c th\u1EA3i, ti\u1EBFt ki\u1EC7m n\u0103ng l\u01B0\u1EE3ng \u0111i\u1EC7n n\u01B0\u1EDBc",
  "T\u1ED5 ch\u1EE9c ho\u1EA1t \u0111\u1ED9ng ph\xE2n lo\u1EA1i r\xE1c v\xE0 gi\u1EA3m thi\u1EC3u \u0111\u1ED3 nh\u1EF1a d\xF9ng m\u1ED9t l\u1EA7n",
  "Th\u1EF1c h\xE0nh l\u1ED1i s\u1ED1ng xanh, ch\u0103m s\xF3c v\xE0 b\u1EA3o v\u1EC7 c\xE2y xanh h\u1ECDc \u0111\u01B0\u1EDDng",
  "\xDD th\u1EE9c th\xEDch \u1EE9ng v\xE0 ph\xF2ng tr\xE1nh tai n\u1EA1n th\u01B0\u01A1ng t\xEDch khi thi\xEAn tai b\xE3o l\u0169"
];
var INTEGRATION_CARDS = [
  {
    key: "nls",
    title: "N\u0103ng l\u1EF1c s\u1ED1",
    desc: "M\u1EE5c ti\xEAu + ho\u1EA1t \u0111\u1ED9ng, ch\u1EEF t\xEDm, m\xE3 ch\u1EC9 b\xE1o \u0111\u1EA7u d\xF2ng \u1EDF m\u1EE5c ti\xEAu chung, cu\u1ED1i c\xE2u trong ho\u1EA1t \u0111\u1ED9ng (vd ... (1.2.TC1a))",
    dot: "#7C3AED"
  },
  {
    key: "ai",
    title: "N\u0103ng l\u1EF1c tr\xED tu\u1EC7 nh\xE2n t\u1EA1o",
    desc: "Theo khung Q\u0110 2422, ch\u1EEF \u0111\u1ECF",
    dot: "#DC2626"
  },
  {
    key: "kt",
    title: "H\u1ECDc sinh khuy\u1EBFt t\u1EADt",
    desc: "M\u1EE5c ti\xEAu ri\xEAng + nhi\u1EC7m v\u1EE5 \u0111i\u1EC1u ch\u1EC9nh \u1EDF t\u1EEBng ho\u1EA1t \u0111\u1ED9ng, ch\u1EEF xanh l\xE1",
    dot: "#15803D"
  },
  {
    key: "qp",
    title: "Gi\xE1o d\u1EE5c qu\u1ED1c ph\xF2ng v\xE0 an ninh",
    desc: "Ch\u1EEF xanh d\u01B0\u01A1ng, c\xF3 ghi [T\xEDch h\u1EE3p GDQPAN]",
    dot: "#1D4ED8"
  },
  {
    key: "bd",
    title: "Bi\u1EBFn \u0111\u1ED5i kh\xED h\u1EADu - ph\xF2ng ch\u1ED1ng thi\xEAn tai",
    desc: "Ch\u1EEF cam, c\xF3 ghi [T\xEDch h\u1EE3p B\u0110KH-PCTT]",
    dot: "#C2410C"
  },
  {
    key: "ltcm",
    title: "Gi\xE1o d\u1EE5c l\xFD t\u01B0\u1EDFng c\xE1ch m\u1EA1ng, \u0111\u1EA1o \u0111\u1EE9c, l\u1ED1i s\u1ED1ng",
    desc: "L\u1EDBp 10\u201312 (Q\u0110 2546) v\xE0 Ng\u1EEF v\u0103n l\u1EDBp 6\u20139, ch\u1EEF v\xE0ng, c\xF3 ghi [T\xEDch h\u1EE3p GD LTCM, \u0110\u0110LS]",
    dot: "#CA8A04"
  },
  {
    key: "atm",
    title: "Gi\xE1o d\u1EE5c an to\xE0n m\u1EA1ng to\xE0n di\u1EC7n",
    desc: "Theo CV 6206, m\u1ECDi l\u1EDBp, m\u1ECDi m\xF4n, ch\u1EEF xanh ng\u1ECDc, m\xE3 ATM1\u2013ATM8 \u0111\u1EA7u d\xF2ng \u1EDF m\u1EE5c ti\xEAu chung, cu\u1ED1i c\xE2u trong ho\u1EA1t \u0111\u1ED9ng (vd ... (ATM4))",
    dot: "#0D9488"
  },
  {
    key: "hcm",
    title: "Gi\xE1o d\u1EE5c t\u01B0 t\u01B0\u1EDFng, \u0111\u1EA1o \u0111\u1EE9c, phong c\xE1ch H\u1ED3 Ch\xED Minh",
    desc: "Theo Q\u0110 2995, m\u1ECDi l\u1EDBp, m\u1ECDi m\xF4n, ch\u1EEF h\u1ED3ng c\xE1nh sen, c\xF3 ghi [T\xEDch h\u1EE3p GD t\u01B0 t\u01B0\u1EDFng, \u0111\u1EA1o \u0111\u1EE9c, phong c\xE1ch HCM]",
    dot: "#DB2777"
  }
];
var DB_NAME = "mhedu_sgk_cache_v1";
var STORE_NAME = "textbooks";
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
  } catch {
    return [];
  }
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
  } catch {
  }
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
  } catch {
  }
}
function guessSubjectAndGradeFromName(fileName) {
  const lower = fileName.toLowerCase();
  let g = null;
  const gm = /(?:^|[^0-9])(1[0-2]|[1-9])(?:[^0-9]|$)/.exec(lower);
  if (gm) g = Number(gm[1]);
  const subjectHints = [
    ["tin", "Tin h\u1ECDc"],
    ["toan", "To\xE1n"],
    ["to\xE1n", "To\xE1n"],
    ["van", "Ng\u1EEF v\u0103n"],
    ["v\u0103n", "Ng\u1EEF v\u0103n"],
    ["anh", "Ti\u1EBFng Anh"],
    ["khtn", "Khoa h\u1ECDc t\u1EF1 nhi\xEAn"],
    ["gdcd", "Gi\xE1o d\u1EE5c c\xF4ng d\xE2n"],
    ["cong nghe", "C\xF4ng ngh\u1EC7"],
    ["vat li", "V\u1EADt l\xED"],
    ["hoa hoc", "Ho\xE1 h\u1ECDc"],
    ["sinh hoc", "Sinh h\u1ECDc"],
    ["lich su", "L\u1ECBch s\u1EED"],
    ["dia li", "\u0110\u1ECBa l\xED"],
    ["tieng viet", "Ti\u1EBFng Vi\u1EC7t"],
    ["khoa hoc", "Khoa h\u1ECDc"]
  ];
  let s = null;
  for (const [kw, label] of subjectHints) {
    if (lower.includes(kw)) {
      s = label;
      break;
    }
  }
  return { grade: g, subject: s };
}
function extractCodesFromText(text, pattern) {
  const matches = String(text || "").match(pattern) || [];
  return [...new Set(matches)];
}
function App() {
  const [mode, setMode] = useState("existing");
  const [step, setStep] = useState(0);
  const [subject, setSubject] = useState("To\xE1n");
  const [grade, setGrade] = useState(10);
  const [title, setTitle] = useState("");
  const [ppctName, setPpctName] = useState("");
  const [ppctText, setPpctText] = useState("");
  const [docInfo, setDocInfo] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [savedTextbooks, setSavedTextbooks] = useState([]);
  const [sgkFile, setSgkFile] = useState(null);
  const [pageFrom, setPageFrom] = useState(5);
  const [pageTo, setPageTo] = useState(8);
  const [pagePreviews, setPagePreviews] = useState([]);
  const [periods, setPeriods] = useState(1);
  const [activityList, setActivityList] = useState([]);
  const [renderingPages, setRenderingPages] = useState(false);
  const [opts, setOpts] = useState({
    nls: true,
    ai: false,
    kt: false,
    qp: false,
    bd: false,
    ltcm: false,
    atm: false,
    hcm: false
  });
  const [selectedNlsItems, setSelectedNlsItems] = useState([]);
  const [selectedAiItems, setSelectedAiItems] = useState([]);
  const [selectedAtmItems, setSelectedAtmItems] = useState([]);
  const [disabilityRows, setDisabilityRows] = useState([{ type: "Khuy\u1EBFt t\u1EADt chung", className: "" }]);
  const [disabilityNote, setDisabilityNote] = useState("");
  const [qpNote, setQpNote] = useState("");
  const [bdNote, setBdNote] = useState("");
  const [ltcmNote, setLtcmNote] = useState("");
  const [hcmNote, setHcmNote] = useState("");
  const [lessonLang, setLessonLang] = useState("Ti\u1EBFng Vi\u1EC7t");
  const [extraInstructions, setExtraInstructions] = useState("");
  const [strongerAi, setStrongerAi] = useState(false);
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
  const handleFileUpload = async (file) => {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setError("Vui l\xF2ng ch\u1ECDn file Word \u0111\u1ECBnh d\u1EA1ng .docx (Word 2007 tr\u1EDF l\xEAn). V\u1EDBi file .doc c\u0169, h\xE3y m\u1EDF b\u1EB1ng Word r\u1ED3i L\u01B0u th\xE0nh .docx.");
      return;
    }
    setError(null);
    setLoading(true);
    setProgress(30);
    setProgressStatus("\u0110ang \u0111\u1ECDc v\xE0 ph\xE2n t\xEDch c\u1EA5u tr\xFAc gi\xE1o \xE1n...");
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
      setError(err.message || "Kh\xF4ng th\u1EC3 \u0111\u1ECDc file Word. Vui l\xF2ng ki\u1EC3m tra l\u1EA1i file.");
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
        setPpctText(parsed.text.slice(0, 6e3));
      } else if (lower.endsWith(".xlsx")) {
        const text = await readXlsxFile(file);
        setPpctName(file.name);
        setPpctText(text.slice(0, 6e3));
      } else {
        setError("PPCT ch\u1EC9 h\u1ED7 tr\u1EE3 file .docx ho\u1EB7c .xlsx.");
      }
    } catch (err) {
      setError("Kh\xF4ng th\u1EC3 \u0111\u1ECDc file PPCT: " + (err.message || "L\u1ED7i \u0111\u1ECBnh d\u1EA1ng"));
    }
  };
  const loadPdfBuffer = async (buf, fileName, idForSave = null) => {
    if (!window.pdfjsLib) throw new Error("Th\u01B0 vi\u1EC7n \u0111\u1ECDc PDF ch\u01B0a s\u1EB5n s\xE0ng.");
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
        id: recordId,
        name: fileName,
        numPages,
        grade: guessed.grade || grade,
        subject: guessed.subject || subject,
        buf,
        updatedAt: Date.now()
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
      setError("Ch\u1EC9 ch\u1ECDn t\u1ED1i \u0111a 20 trang SGK cho m\u1ED7i b\xE0i d\u1EA1y.");
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
          text = tc.items.map((it) => it.str).join(" ").replace(/\s+/g, " ").trim();
        } catch {
        }
        previews.push({ pageNum: p, dataUrl, text });
      }
      setPagePreviews(previews);
    } catch (e) {
      setError("Kh\xF4ng th\u1EC3 hi\u1EC3n th\u1ECB trang SGK: " + (e.message || ""));
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
        setError("L\u1ED7i \u0111\u1ECDc file PDF SGK: " + (e.message || ""));
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
      setSgkFile({ id: `imgs-${Date.now()}`, name: `${imgs.length} \u1EA3nh ch\u1EE5p SGK`, type: "images", numPages: imgs.length });
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
      setError("Kh\xF4ng th\u1EC3 m\u1EDF SGK \u0111\xE3 l\u01B0u: " + (e.message || ""));
    } finally {
      setRenderingPages(false);
    }
  };
  const callGeminiJson = async (parts) => {
    let userKey = "";
    try {
      userKey = (localStorage.getItem("smartedu_gemini_api_key") || "").trim();
    } catch {
    }
    const BACKUP_KEYS = [];
    const keysToTry = (userKey ? [userKey, ...BACKUP_KEYS] : BACKUP_KEYS).filter(Boolean);
    if (keysToTry.length === 0) {
      if (typeof window.openKeyModal === "function") window.openKeyModal();
      throw new Error("Ch\u01B0a c\xF3 Gemini API Key. Vui l\xF2ng l\u1EA5y key mi\u1EC5n ph\xED t\u1EA1i https://aistudio.google.com/apikey v\xE0 d\xE1n v\xE0o h\u1ED9p tho\u1EA1i '\u{1F511} \u0110\u1ED5i API Key'.");
    }
    const MODELS = strongerAi ? ["gemini-3.8-flash", "gemini-3.5-flash", "gemini-3-flash-preview", "gemini-3.5-flash-lite"] : ["gemini-3.5-flash", "gemini-3-flash-preview", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-3.8-flash"];
    let lastErr = null;
    for (const key of keysToTry) {
      for (const model of MODELS) {
        try {
          setProgressStatus(`AI (${model}) \u0111ang ph\xE2n t\xEDch v\xE0 x\xE2y d\u1EF1ng n\u1ED9i dung t\xEDch h\u1EE3p...`);
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
          console.warn(`Th\u1EED model ${model} th\u1EA5t b\u1EA1i:`, e.message);
        }
      }
    }
    throw lastErr || new Error("AI kh\xF4ng ph\u1EA3n h\u1ED3i.");
  };
  const buildIntegrationRequirementsText = () => {
    const hsktList = disabilityRows.filter((r) => r.type).map((r) => r.className ? `${r.type} (L\u1EDBp ${r.className})` : r.type);
    return `
NG\xD4N NG\u1EEE HI\u1EC2N TH\u1ECA C\u1EE6A GI\xC1O \xC1N: ${lessonLang}.
${ppctText ? `C\u0102N C\u1EE8 PH\xC2N PH\u1ED0I CH\u01AF\u01A0NG TR\xCCNH (PPCT) \u0110\xC3 T\u1EA2I (d\xF9ng chung NLS, NL AI, An to\xE0n m\u1EA1ng):
"""
${ppctText.slice(0, 3500)}
"""
` : ""}
C\xC1C N\u1ED8I DUNG Y\xCAU C\u1EA6U T\xCDCH H\u1EE2P:
${opts.nls ? `- N\u0102NG L\u1EF0C S\u1ED0 (NLS, m\xE3 "nls", theo TT 02/2025/TT-BGD\u0110T g\u1ED3m 6 mi\u1EC1n: I. Khai th\xE1c d\u1EEF li\u1EC7u & th\xF4ng tin [1.1\u20131.3]; II. Giao ti\u1EBFp & h\u1EE3p t\xE1c [2.1\u20132.6]; III. S\xE1ng t\u1EA1o n\u1ED9i dung s\u1ED1 [3.1\u20133.4]; IV. An to\xE0n [4.1\u20134.4]; V. Gi\u1EA3i quy\u1EBFt v\u1EA5n \u0111\u1EC1 [5.1\u20135.4]; VI. \u1EE8ng d\u1EE5ng AI [6.1\u20136.3]): L\u1ED3ng gh\xE9p m\u1EE5c ti\xEAu + ho\u1EA1t \u0111\u1ED9ng; m\xE3 ch\u1EC9 b\xE1o \u0111\u1EB7t \u0111\u1EA7u d\xF2ng \u1EDF m\u1EE5c ti\xEAu chung (vd: "1.2.${getNlsLevelTag(grade)}a \u2013 ..."), ho\u1EB7c cu\u1ED1i c\xE2u trong ho\u1EA1t \u0111\u1ED9ng (vd: "... (1.2.${getNlsLevelTag(grade)}a)"). ${selectedNlsItems.length ? `B\u1EAFt bu\u1ED9c \u01B0u ti\xEAn c\xE1c m\xE3 \u0111\xE3 ch\u1ECDn: ${selectedNlsItems.map((x) => `${x.code} (${x.desc})`).join("; ")}.` : `T\u1EF1 ch\u1ECDn m\xE3 ph\xF9 h\u1EE3p trong khung l\u1EDBp ${grade} (${getNlsLevelLabel(grade)}).`}` : ""}
${opts.ai ? `- N\u0102NG L\u1EF0C TR\xCD TU\u1EC6 NH\xC2N T\u1EA0O (NL AI, m\xE3 "ai"): Theo Khung n\u1ED9i dung gi\xE1o d\u1EE5c AI Q\u0110 2422/Q\u0110-BGD\u0110T (18/8/2026) g\u1ED3m 4 m\u1EA1ch A (T\u01B0 duy l\u1EA5y con ng\u01B0\u1EDDi l\xE0m trung t\xE2m), B (\u0110\u1EA1o \u0111\u1EE9c AI), C (K\u0129 thu\u1EADt & \u1EE9ng d\u1EE5ng AI), D (Thi\u1EBFt k\u1EBF h\u1EC7 th\u1ED1ng AI); m\xE3 h\xF3a chu\u1EA9n [L\u1EDBp].[M\xE3 ch\u1EE7 \u0111\u1EC1].[STT] (m\u1EDF r\u1ED9ng c\xF3 MR). L\u1ED3ng gh\xE9p nh\u1EB9 nh\xE0ng, kh\xF4ng g\xE2y qu\xE1 t\u1EA3i m\xF4n h\u1ECDc, ch\xFA tr\u1ECDng ki\u1EC3m ch\u1EE9ng th\xF4ng tin v\xE0 khai b\xE1o minh b\u1EA1ch khi d\xF9ng AI. ${selectedAiItems.length ? `B\u1EAFt bu\u1ED9c \u01B0u ti\xEAn c\xE1c m\xE3: ${selectedAiItems.map((x) => `${x.code} (${x.desc})`).join("; ")}.` : `Ch\u1ECDn \u0111\xFAng m\xE3 YCC\u0110 ph\xF9 h\u1EE3p trong khung l\u1EDBp ${grade} d\u01B0\u1EDBi \u0111\xE2y:
${aiCatalog.map((x) => `  + ${x.code}: ${x.desc}`).join("\n")}`}` : ""}
${opts.kt ? `- H\u1ECCC SINH KHUY\u1EBET T\u1EACT H\xD2A NH\u1EACP (m\xE3 "kt"): L\u1ED3ng gh\xE9p m\u1EE5c ti\xEAu ri\xEAng + nhi\u1EC7m v\u1EE5 \u0111i\u1EC1u ch\u1EC9nh gi\u1EA3m t\u1EA3i \u1EDF t\u1EEBng ho\u1EA1t \u0111\u1ED9ng cho: ${hsktList.join("; ") || "Khuy\u1EBFt t\u1EADt chung"}. Ch\u1EC9 ghi t\xEAn l\u1EDBp (n\u1EBFu c\xF3), KH\xD4NG ghi t\xEAn h\u1ECDc sinh. ${disabilityNote ? `[Ghi ch\xFA: ${disabilityNote}]` : ""}` : ""}
${opts.qp ? `- GI\xC1O D\u1EE4C QU\u1ED0C PH\xD2NG V\xC0 AN NINH (GDQPAN, m\xE3 "qp"): L\u1ED3ng gh\xE9p theo TT 08/2024/TT-BGD\u0110T. ${qpNote ? `[Ghi ch\xFA: ${qpNote}]` : ""}` : ""}
${opts.bd ? `- BI\u1EBEN \u0110\u1ED4I KH\xCD H\u1EACU - PH\xD2NG CH\u1ED0NG THI\xCAN TAI (B\u0110KH-PCTT, m\xE3 "bd"): ${bdNote ? `[Ghi ch\xFA: ${bdNote}]` : "L\u1ED3ng gh\xE9p gi\xE1o d\u1EE5c m\xF4i tr\u01B0\u1EDDng, l\u1ED1i s\u1ED1ng xanh, ti\u1EBFt ki\u1EC7m n\u0103ng l\u01B0\u1EE3ng ho\u1EB7c ph\xF2ng tr\xE1nh thi\xEAn tai."}` : ""}
${opts.ltcm ? `- GI\xC1O D\u1EE4C L\xDD T\u01AF\u1EDENG C\xC1CH M\u1EA0NG, \u0110\u1EA0O \u0110\u1EE8C, L\u1ED0I S\u1ED0NG (m\xE3 "ltcm"): Theo Q\u0110 2546 (l\u1EDBp 10\u201312) ho\u1EB7c Ng\u1EEF v\u0103n 6\u20139. ${ltcmNote ? `[Ghi ch\xFA: ${ltcmNote}]` : ""}` : ""}
${opts.atm ? `- GI\xC1O D\u1EE4C AN TO\xC0N M\u1EA0NG TO\xC0N DI\u1EC6N (m\xE3 "atm"): Theo C\xF4ng v\u0103n 6206/BGD\u0110T-GDPT (14/9/2026) v\xE0 S\u01A1 \u0111\u1ED3 tra c\u1EE9u t\xEDch h\u1EE3p An to\xE0n m\u1EA1ng \u2013 N\u0103ng l\u1EF1c s\u1ED1 TT 02/2025:
${ATM_CATALOG.map((a) => `  + ${a.code} (${a.title}): NL b\u1EAFt bu\u1ED9c [${a.requiredNls.join(", ")}]; NL khuy\u1EBFn kh\xEDch [${a.recommendedNls.join(", ")}] \u2014 ${a.desc}`).join("\n")}
  \u0110\u1EB7t m\xE3 ATM1\u2013ATM8 \u0111\u1EA7u d\xF2ng \u1EDF m\u1EE5c ti\xEAu chung, cu\u1ED1i c\xE2u trong ho\u1EA1t \u0111\u1ED9ng (vd "... (ATM4)"). Khi t\xEDch h\u1EE3p c\xF9ng NLS, \u0111\u1ED1i chi\u1EBFu \u0111\xFAng c\xE1c m\xE3 NL TT02 li\xEAn k\u1EBFt \u1EDF tr\xEAn. ${selectedAtmItems.length ? `\u01AFu ti\xEAn c\xE1c m\xE3: ${selectedAtmItems.map((x) => `${x.code} (${x.title}: ${x.desc})`).join("; ")}.` : "T\u1EF1 ch\u1ECDn 1\u20132 m\xE3 ATM1\u2013ATM8 ph\xF9 h\u1EE3p b\xE0i h\u1ECDc."}` : ""}
${opts.hcm ? `- GI\xC1O D\u1EE4C T\u01AF T\u01AF\u1EDENG, \u0110\u1EA0O \u0110\u1EE8C, PHONG C\xC1CH H\u1ED2 CH\xCD MINH (m\xE3 "hcm"): Theo Q\u0110 2995. ${hcmNote ? `[Ghi ch\xFA: ${hcmNote}]` : ""}` : ""}
${extraInstructions.trim() ? `- \u0110I\u1EC0U CH\u1EC8NH TH\xCAM T\u1EEA GI\xC1O VI\xCAN:
${extraInstructions.trim()}` : ""}
`.trim();
  };
  const buildAutoIntegrationRequirementsText = () => {
    return `
CH\u1EBE \u0110\u1ED8 T\u1EF0 \u0110\u1ED8NG NH\u1EACN DI\u1EC6N & T\xCDCH H\u1EE2P TH\xD4NG MINH:
- Ng\xF4n ng\u1EEF hi\u1EC3n th\u1ECB c\u1EE7a gi\xE1o \xE1n: ${lessonLang}.
${ppctText ? `- C\u0103n c\u1EE9 PPCT \u0111\xE3 t\u1EA3i:
"""
${ppctText.slice(0, 3500)}
"""
` : ""}
- H\xE3y t\u1EF1 ph\xE2n t\xEDch n\u1ED9i dung b\xE0i h\u1ECDc, m\xF4n h\u1ECDc (${subject}), kh\u1ED1i l\u1EDBp (L\u1EDBp ${grade}) \u0111\u1EC3 T\u1EF0 \u0110\u1ED8NG CH\u1ECCN t\u1EEB 2 \u0111\u1EBFn 4 n\u1ED9i dung t\xEDch h\u1EE3p PH\xD9 H\u1EE2P NH\u1EA4T, T\u1EF0 NHI\xCAN NH\u1EA4T (kh\xF4ng nh\u1ED3i nh\xE9t khi\xEAn c\u01B0\u1EE1ng) trong 8 lo\u1EA1i sau:
  1. "nls" \u2014 N\u0103ng l\u1EF1c s\u1ED1 (TT 02/2025/TT-BGD\u0110T, m\u1EE9c ${getNlsLevelLabel(grade)}, m\xE3 ${getNlsLevelTag(grade)}, vd 1.2.${getNlsLevelTag(grade)}a \u0111\u1EB7t \u0111\u1EA7u d\xF2ng m\u1EE5c ti\xEAu chung ho\u1EB7c cu\u1ED1i c\xE2u ho\u1EA1t \u0111\u1ED9ng).
  2. "ai" \u2014 N\u0103ng l\u1EF1c Tr\xED tu\u1EC7 nh\xE2n t\u1EA1o (Q\u0110 2422/Q\u0110-BGD\u0110T), d\xF9ng \u0111\xFAng m\xE3 l\u1EDBp ${grade} trong danh s\xE1ch:
${aiCatalog.map((x) => `     + ${x.code}: ${x.desc}`).join("\n")}
  3. "atm" \u2014 Gi\xE1o d\u1EE5c an to\xE0n m\u1EA1ng to\xE0n di\u1EC7n (CV 6206/BGD\u0110T-GDPT), \u0111\u1EB7t m\xE3 ATM1\u2013ATM8 \u0111\u1EA7u d\xF2ng m\u1EE5c ti\xEAu ho\u1EB7c cu\u1ED1i c\xE2u ho\u1EA1t \u0111\u1ED9ng, li\xEAn k\u1EBFt \u0111\xFAng m\xE3 TT 02/2025:
${ATM_CATALOG.map((a) => `     + ${a.code} (${a.title}) -> NL b\u1EAFt bu\u1ED9c [${a.requiredNls.join(", ")}], khuy\u1EBFn kh\xEDch [${a.recommendedNls.join(", ")}]`).join("\n")}
  4. "kt" \u2014 H\u1ECDc sinh khuy\u1EBFt t\u1EADt h\xF2a nh\u1EADp (ch\u1EC9 ch\u1ECDn n\u1EBFu gi\xE1o \xE1n c\xF3 nh\u1EAFc t\u1EDBi HSKT ho\u1EB7c ho\u1EA1t \u0111\u1ED9ng c\u1EA7n \u0111i\u1EC1u ch\u1EC9nh gi\u1EA3m t\u1EA3i).
  5. "qp" \u2014 Gi\xE1o d\u1EE5c Qu\u1ED1c ph\xF2ng & An ninh (TT 08/2024/TT-BGD\u0110T).
  6. "bd" \u2014 Bi\u1EBFn \u0111\u1ED5i kh\xED h\u1EADu & Ph\xF2ng ch\u1ED1ng thi\xEAn tai (m\xF4i tr\u01B0\u1EDDng, ti\u1EBFt ki\u1EC7m n\u0103ng l\u01B0\u1EE3ng, l\u1ED1i s\u1ED1ng xanh).
  7. "ltcm" \u2014 Gi\xE1o d\u1EE5c l\xFD t\u01B0\u1EDFng c\xE1ch m\u1EA1ng, \u0111\u1EA1o \u0111\u1EE9c, l\u1ED1i s\u1ED1ng (Q\u0110 2546 l\u1EDBp 10\u201312 ho\u1EB7c Ng\u1EEF v\u0103n 6\u20139).
  8. "hcm" \u2014 Gi\xE1o d\u1EE5c t\u01B0 t\u01B0\u1EDFng, \u0111\u1EA1o \u0111\u1EE9c, phong c\xE1ch H\u1ED3 Ch\xED Minh (Q\u0110 2995).
${extraInstructions.trim() ? `- \u0110i\u1EC1u ch\u1EC9nh th\xEAm t\u1EEB gi\xE1o vi\xEAn:
${extraInstructions.trim()}` : ""}
`.trim();
  };
  const handleRunAi = async (autoMode = false) => {
    const activeOpts = Object.entries(opts).filter(([_, v]) => v).map(([k]) => k);
    if (!autoMode && activeOpts.length === 0) {
      setError("Vui l\xF2ng b\u1EADt \xEDt nh\u1EA5t 1 n\u1ED9i dung c\u1EA7n t\xEDch h\u1EE3p ho\u1EB7c d\xF9ng n\xFAt T\u1EF1 \u0111\u1ED9ng nh\u1EADn di\u1EC7n.");
      return;
    }
    if (mode === "existing" && !docInfo) {
      setError("Ch\u01B0a c\xF3 gi\xE1o \xE1n \u0111\u1EC3 t\xEDch h\u1EE3p. H\xE3y ch\u1ECDn file Word tr\u01B0\u1EDBc.");
      setStep(0);
      return;
    }
    if (mode === "new" && pagePreviews.length === 0) {
      setError("Ch\u01B0a c\xF3 trang SGK n\xE0o \u0111\u01B0\u1EE3c ch\u1ECDn. Vui l\xF2ng t\u1EA3i SGK v\xE0 b\u1EA5m OK \u1EDF B\u01B0\u1EDBc 1.");
      setStep(0);
      return;
    }
    setError(null);
    setLoading(true);
    setProgress(15);
    setProgressStatus(autoMode ? "AI \u0111ang t\u1EF1 \u0111\u1ED9ng ph\xE2n t\xEDch b\xE0i d\u1EA1y & ch\u1ECDn n\u1ED9i dung t\xEDch h\u1EE3p ph\xF9 h\u1EE3p..." : "\u0110ang chu\u1EA9n b\u1ECB d\u1EEF li\u1EC7u v\xE0 k\u1EBFt n\u1ED1i Gemini AI...");
    const timer = setInterval(() => {
      setProgress((p) => p < 88 ? p + Math.floor(Math.random() * 7) + 3 : p);
    }, 450);
    try {
      if (mode === "existing") {
        const parasSlice = docInfo.text.split("\n").slice(0, 180).join("\n");
        const promptText = `B\u1EA1n l\xE0 Chuy\xEAn gia Ph\u01B0\u01A1ng ph\xE1p D\u1EA1y h\u1ECDc v\xE0 Tr\u1EE3 l\xFD AI Gi\xE1o d\u1EE5c c\u1EE7a MHEdu.
Nhi\u1EC7m v\u1EE5: Ph\xE2n t\xEDch b\xE0i gi\xE1o \xE1n d\u01B0\u1EDBi \u0111\xE2y v\xE0 \u0111\u1EC1 xu\u1EA5t c\xE1c \u0111o\u1EA1n ch\xE8n t\xEDch h\u1EE3p v\xE0o \u0111\xFAng c\xE1c ho\u1EA1t \u0111\u1ED9ng (M\u1EE5c ti\xEAu, Kh\u1EDFi \u0111\u1ED9ng, H\xECnh th\xE0nh ki\u1EBFn th\u1EE9c, Luy\u1EC7n t\u1EADp, V\u1EADn d\u1EE5ng).

TH\xD4NG TIN B\xC0I H\u1ECCC:
- M\xF4n h\u1ECDc: ${subject}
- Kh\u1ED1i l\u1EDBp: L\u1EDBp ${grade}
- T\xEAn b\xE0i: ${title || docInfo.fileName}

${autoMode ? buildAutoIntegrationRequirementsText() : buildIntegrationRequirementsText()}

DANH S\xC1CH C\xC1C \u0110O\u1EA0N V\u0102N C\u1EE6A GI\xC1O \xC1N G\u1ED0C (c\xF3 g\u1EAFn ID \u0111o\u1EA1n [p1], [p2], ...):
"""
${parasSlice}
"""

QUY T\u1EAEC B\u1EAET BU\u1ED8C:
1. \u0110\u1ECDc k\u1EF9 c\xE1c \u0111o\u1EA1n v\u0103n, x\xE1c \u0111\u1ECBnh \u0111\xFAng c\xE1c ph\u1EA7n M\u1EE5c ti\xEAu, Kh\u1EDFi \u0111\u1ED9ng, H\xECnh th\xE0nh ki\u1EBFn th\u1EE9c, Luy\u1EC7n t\u1EADp, V\u1EADn d\u1EE5ng.
2. Ch\u1ECDn \u0111\xFAng m\xE3 \u0111o\u1EA1n "a" (v\xED d\u1EE5 "p4", "p12", "p25") c\u1EA7n ch\xE8n th\xEAm n\u1ED9i dung.
3. Tr\u1EA3 v\u1EC1 \u0111\xFAng JSON thu\u1EA7n t\xFAy:
{
  "items": [
    { "a": "p4", "w": "below", "k": "nls", "t": "1.1.TC2b \u2013 N\u1ED9i dung t\xEDch h\u1EE3p c\u1EE5 th\u1EC3..." }
  ]
}
Trong \u0111\xF3 "w" l\xE0 "below" ho\u1EB7c "end", "k" thu\u1ED9c ["nls","ai","kt","qp","bd","ltcm","atm","hcm"]. T\u1EA1o 6 \u0111\u1EBFn 16 \u0111o\u1EA1n ch\xE8n ph\xE2n b\u1ED5 h\u1EE3p l\xFD cho t\u1EA5t c\u1EA3 c\xE1c n\u1ED9i dung \u0111\xE3 b\u1EADt.`;
        const parsed = await callGeminiJson([{ text: promptText }]);
        if (!parsed || !Array.isArray(parsed.items) || parsed.items.length === 0) {
          throw new Error("AI tr\u1EA3 v\u1EC1 d\u1EEF li\u1EC7u tr\u1ED1ng.");
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
          const used = new Set(mappedItems.map((x) => x.k));
          setOpts({
            nls: used.has("nls"),
            ai: used.has("ai"),
            kt: used.has("kt"),
            qp: used.has("qp"),
            bd: used.has("bd"),
            ltcm: used.has("ltcm"),
            atm: used.has("atm"),
            hcm: used.has("hcm")
          });
        }
        setItems(mappedItems);
        setProgress(100);
        setStep(2);
      } else {
        const extractedTexts = pagePreviews.map((p) => `--- TRANG ${p.pageNum} ---
${p.text || "(Trang \u1EA3nh ch\u1EE5p - xem h\xECnh \u0111\xEDnh k\xE8m)"}`).join("\n\n");
        const promptText = `B\u1EA1n l\xE0 Chuy\xEAn gia Ph\u01B0\u01A1ng ph\xE1p D\u1EA1y h\u1ECDc (C\xF4ng v\u0103n 5512/BGD\u0110T) v\xE0 Tr\u1EE3 l\xFD AI Gi\xE1o d\u1EE5c MHEdu.
Nhi\u1EC7m v\u1EE5: D\u1EF1a v\xE0o n\u1ED9i dung c\xE1c trang S\xE1ch gi\xE1o khoa (SGK) \u0111\xEDnh k\xE8m, h\xE3y so\u1EA1n m\u1ED9t K\u1EBF ho\u1EA1ch b\xE0i d\u1EA1y (Gi\xE1o \xE1n) ho\xE0n ch\u1EC9nh, chu\u1EA9n s\u01B0 ph\u1EA1m v\xE0 t\xEDch h\u1EE3p \u0111\u1EA7y \u0111\u1EE7 c\xE1c n\u1ED9i dung \u0111\u01B0\u1EE3c y\xEAu c\u1EA7u.

TH\xD4NG TIN B\xC0I D\u1EA0Y:
- M\xF4n h\u1ECDc: ${subject}
- Kh\u1ED1i l\u1EDBp: L\u1EDBp ${grade}
- T\xEAn b\xE0i: ${title || "B\xE0i h\u1ECDc t\u1EEB SGK"}
- Th\u1EDDi l\u01B0\u1EE3ng: ${periods} ti\u1EBFt
${activityList.filter((a) => a.trim()).length ? `- K\u1EBF ho\u1EA1ch ho\u1EA1t \u0111\u1ED9ng theo th\u1EE9 t\u1EF1:
${activityList.filter((a) => a.trim()).map((a, i) => `  ${i + 1}. ${a.trim()}`).join("\n")}` : ""}

${autoMode ? buildAutoIntegrationRequirementsText() : buildIntegrationRequirementsText()}

N\u1ED8I DUNG CH\u1EEE TR\xCDCH XU\u1EA4T T\u1EEA SGK:
"""
${extractedTexts.slice(0, 8e3)}
"""

Tr\u1EA3 v\u1EC1 \u0111\xFAng JSON thu\u1EA7n t\xFAy:
{
  "title": "T\xEAn b\xE0i h\u1ECDc chu\u1EA9n",
  "sections": [
    {
      "heading": "I. M\u1EE4C TI\xCAU",
      "blocks": [
        { "type": "normal", "text": "1. Ki\u1EBFn th\u1EE9c: ..." },
        { "type": "nls", "text": "1.1.TC2b \u2013 ..." }
      ]
    },
    { "heading": "II. THI\u1EBET B\u1ECA D\u1EA0Y H\u1ECCC V\xC0 H\u1ECCC LI\u1EC6U", "blocks": [...] },
    { "heading": "III. TI\u1EBEN TR\xCCNH D\u1EA0Y H\u1ECCC", "blocks": [...] }
  ]
}
Trong \u0111\xF3 "type" ch\u1EC9 nh\u1EADn: "normal", "nls", "ai", "kt", "qp", "bd", "ltcm", "atm", "hcm".`;
        const parts = [{ text: promptText }];
        for (const p of pagePreviews.slice(0, 6)) {
          if (p.dataUrl && p.dataUrl.includes(",")) {
            parts.push({ inlineData: { mimeType: "image/jpeg", data: p.dataUrl.split(",")[1] } });
          }
        }
        const parsed = await callGeminiJson(parts);
        if (!parsed || !Array.isArray(parsed.sections) || parsed.sections.length === 0) {
          throw new Error("AI tr\u1EA3 v\u1EC1 c\u1EA5u tr\xFAc gi\xE1o \xE1n ch\u01B0a h\u1EE3p l\u1EC7.");
        }
        const builtLesson = {
          title: parsed.title || title || "K\u1EBF ho\u1EA1ch b\xE0i d\u1EA1y t\xEDch h\u1EE3p",
          subject,
          grade,
          periods,
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
          const used = new Set(builtLesson.sections.flatMap((s) => s.blocks.map((b) => b.type)).filter((t) => t !== "normal"));
          setOpts({
            nls: used.has("nls"),
            ai: used.has("ai"),
            kt: used.has("kt"),
            qp: used.has("qp"),
            bd: used.has("bd"),
            ltcm: used.has("ltcm"),
            atm: used.has("atm"),
            hcm: used.has("hcm")
          });
        }
        setNewLessonData(builtLesson);
        setProgress(100);
        setStep(2);
      }
    } catch (e) {
      setError("AI kh\xF4ng th\u1EC3 ph\u1EA3n h\u1ED3i: " + (e?.message || "Ki\u1EC3m tra l\u1EA1i k\u1EBFt n\u1ED1i m\u1EA1ng ho\u1EB7c Gemini API Key."));
    } finally {
      clearInterval(timer);
      setLoading(false);
    }
  };
  const integrationSummary = useMemo(() => {
    const list = mode === "existing" ? items.filter((i) => i.on).map((i) => ({ k: i.k, t: i.t })) : (newLessonData?.sections || []).flatMap((s) => s.blocks.filter((b) => b.type !== "normal").map((b) => ({ k: b.type, t: b.text })));
    const byCat = {};
    for (const k of Object.keys(CATEGORIES)) {
      byCat[k] = list.filter((x) => x.k === k);
    }
    const nlsCodes = extractCodesFromText((byCat.nls || []).map((x) => x.t).join(" "), /\b\d+\.\d+\.[A-Z]{2}\d+[a-z]?\b/g);
    const aiCodes = extractCodesFromText((byCat.ai || []).map((x) => x.t).join(" "), /\b(?:\d+\.[A-D]\d+\.(?:MR)?\d+|[A-D]\d+)\b/g);
    const atmCodes = extractCodesFromText((byCat.atm || []).map((x) => x.t).join(" "), /\bATM[1-8]\b/g);
    return { total: list.length, byCat, nlsCodes, aiCodes, atmCodes };
  }, [mode, items, newLessonData]);
  const handleCopyReport = () => {
    const s = integrationSummary;
    const rpt = `\u{1F4CB} B\xC1O C\xC1O T\xCDCH H\u1EE2P GI\xC1O \xC1N
\u2022 N\u1EC1n t\u1EA3ng: MHEdu AI
\u2022 B\xE0i h\u1ECDc: ${title || docInfo?.fileName || newLessonData?.title || "K\u1EBF ho\u1EA1ch b\xE0i d\u1EA1y"}
\u2022 M\xF4n h\u1ECDc: ${subject} - Kh\u1ED1i l\u1EDBp: L\u1EDBp ${grade}
${ppctName ? `\u2022 Ngu\u1ED3n PPCT: ${ppctName}
` : ""}\u2022 T\u1ED5ng s\u1ED1 v\u1ECB tr\xED t\xEDch h\u1EE3p: ${s.total} v\u1ECB tr\xED
  + N\u0103ng l\u1EF1c s\u1ED1 (NLS): ${s.byCat.nls.length} v\u1ECB tr\xED${s.nlsCodes.length ? ` (${s.nlsCodes.join(", ")})` : ""}
  + N\u0103ng l\u1EF1c AI (Q\u0110 2422): ${s.byCat.ai.length} v\u1ECB tr\xED${s.aiCodes.length ? ` (${s.aiCodes.join(", ")})` : ""}
  + H\u1ECDc sinh khuy\u1EBFt t\u1EADt: ${s.byCat.kt.length} v\u1ECB tr\xED
  + GDQP&AN (TT 08/2024): ${s.byCat.qp.length} v\u1ECB tr\xED
  + B\u0110KH-PCTT: ${s.byCat.bd.length} v\u1ECB tr\xED
  + GD LTCM, \u0110\u0110LS: ${s.byCat.ltcm.length} v\u1ECB tr\xED
  + An to\xE0n m\u1EA1ng (CV 6206): ${s.byCat.atm.length} v\u1ECB tr\xED${s.atmCodes.length ? ` (${s.atmCodes.join(", ")})` : ""}
  + T\u01B0 t\u01B0\u1EDFng HCM (Q\u0110 2995): ${s.byCat.hcm.length} v\u1ECB tr\xED
\u2022 Th\u1EDDi gian: ${(/* @__PURE__ */ new Date()).toLocaleString("vi-VN")}`;
    navigator.clipboard.writeText(rpt).then(() => {
      setCopiedReport(true);
      setTimeout(() => setCopiedReport(false), 3e3);
    }).catch(() => alert("Kh\xF4ng th\u1EC3 sao ch\xE9p t\u1EF1 \u0111\u1ED9ng."));
  };
  const handleDownloadWord = async () => {
    setDownloading(true);
    try {
      let blob;
      let outName;
      if (mode === "existing") {
        if (!docInfo || items.length === 0) return;
        blob = await generatePatchedDocx(docInfo, items.filter((it) => it.on), colorEnabled);
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
      alert("L\u1ED7i khi t\u1EA1o file Word: " + (err.message || "Kh\xF4ng x\xE1c \u0111\u1ECBnh"));
    } finally {
      setDownloading(false);
    }
  };
  const filteredItems = useMemo(() => {
    if (activeFilter === "all") return items;
    return items.filter((it) => it.k === activeFilter);
  }, [items, activeFilter]);
  const activeCount = useMemo(() => items.filter((it) => it.on).length, [items]);
  const step1Ready = mode === "existing" ? !!docInfo : pagePreviews.length > 0;
  const step3Ready = mode === "existing" ? items.length > 0 : !!newLessonData;
  return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-slate-50 text-slate-800 font-sans pb-24" }, /* @__PURE__ */ React.createElement("div", { className: "max-w-6xl mx-auto px-4 pt-6" }, /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-3xl p-6 shadow-sm border border-slate-200/80 mb-5" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-4" }, /* @__PURE__ */ React.createElement("div", { className: "w-12 h-12 rounded-2xl bg-gradient-to-tr from-purple-600 to-indigo-500 flex items-center justify-center text-white text-2xl shadow-md shrink-0" }, "\u{1F9E9}"), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2.5 flex-wrap" }, /* @__PURE__ */ React.createElement("h1", { className: "text-xl sm:text-2xl font-black text-slate-900 tracking-tight" }, "T\xCDCH H\u1EE2P T\u1ED4NG H\u1EE2P V\xC0O GI\xC1O \xC1N"), /* @__PURE__ */ React.createElement("span", { className: "px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 text-xs font-bold uppercase" }, "\u2728 B\u1EA2N \u0110\u1EA6Y \u0110\u1EE6")), /* @__PURE__ */ React.createElement("p", { className: "text-xs sm:text-sm text-slate-500 mt-1 font-medium" }, "H\u1ED7 tr\u1EE3 8 n\u1ED9i dung t\xEDch h\u1EE3p chu\u1EA9n B\u1ED9 GD&\u0110T (NLS, NL AI, HSKT, GDQPAN, B\u0110KH, LTCM-\u0110\u0110LS, An to\xE0n m\u1EA1ng CV 6206, T\u01B0 t\u01B0\u1EDFng HCM Q\u0110 2995)."))), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3 mt-5 pt-4 border-t border-slate-100" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => switchMode("existing"),
      className: `p-3 rounded-2xl border-2 text-center font-extrabold text-sm transition-all ${mode === "existing" ? "border-purple-600 bg-purple-600 text-white shadow-sm" : "border-slate-200 bg-white text-slate-600 hover:border-purple-300"}`
    },
    "\u{1F4C4} T\xEDch h\u1EE3p v\xE0o gi\xE1o \xE1n c\xF3 s\u1EB5n"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => switchMode("new"),
      className: `p-3 rounded-2xl border-2 text-center font-extrabold text-sm transition-all ${mode === "new" ? "border-purple-600 bg-purple-600 text-white shadow-sm" : "border-slate-200 bg-white text-slate-600 hover:border-purple-300"}`
    },
    "\u{1F4D6} T\u1EA1o m\u1ED9t gi\xE1o \xE1n m\u1EDBi c\xF3 t\xEDch h\u1EE3p"
  ))), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-3 gap-2 sm:gap-4 mb-6" }, [
    { id: 0, num: "1", title: mode === "existing" ? "Gi\xE1o \xE1n" : "B\xE0i d\u1EA1y & SGK" },
    { id: 1, num: "2", title: "N\u1ED9i dung t\xEDch h\u1EE3p" },
    { id: 2, num: "3", title: mode === "existing" ? "T\u1EA3i Word" : "Xem \xB7 T\u1EA3i Word" }
  ].map((s) => {
    const isActive = step === s.id;
    const isDone = s.id === 0 && step1Ready || s.id === 1 && step3Ready || s.id === 2 && step3Ready;
    const canClick = s.id === 0 || s.id === 1 && step1Ready || s.id === 2 && step3Ready;
    return /* @__PURE__ */ React.createElement(
      "button",
      {
        key: s.id,
        onClick: () => canClick && setStep(s.id),
        className: `p-3 rounded-2xl border text-center transition-all flex items-center justify-center gap-2 ${isActive ? "bg-purple-600 text-white border-purple-600 shadow-md" : isDone ? "bg-white text-emerald-700 border-emerald-300" : "bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed"}`
      },
      /* @__PURE__ */ React.createElement("span", { className: `w-6 h-6 rounded-full flex items-center justify-center text-xs font-black ${isActive ? "bg-white text-purple-700" : isDone ? "bg-emerald-600 text-white" : "bg-slate-200 text-slate-500"}` }, isDone && !isActive ? "\u2713" : s.num),
      /* @__PURE__ */ React.createElement("span", { className: "font-bold text-sm" }, s.title)
    );
  })), error && /* @__PURE__ */ React.createElement("div", { className: "bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-2xl text-sm mb-6 flex items-start gap-3" }, /* @__PURE__ */ React.createElement("span", { className: "text-xl" }, "\u26A0\uFE0F"), /* @__PURE__ */ React.createElement("div", { className: "flex-1 font-medium" }, error), /* @__PURE__ */ React.createElement("button", { onClick: () => setError(null), className: "text-red-400 hover:text-red-700 font-bold" }, "\u2715")), step === 0 && mode === "existing" && /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("h2", { className: "text-lg font-bold text-slate-800 flex items-center gap-2" }, /* @__PURE__ */ React.createElement("span", null, "\u{1F4C4}"), " Gi\xE1o \xE1n c\u1EA7n t\xEDch h\u1EE3p"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-slate-500" }, "T\u1EA3i file Word v\xE0 ki\u1EC3m tra th\xF4ng tin b\xE0i d\u1EA1y")), !docInfo ? /* @__PURE__ */ React.createElement(
    "div",
    {
      onDragOver: (e) => {
        e.preventDefault();
        setIsDragging(true);
      },
      onDragLeave: () => setIsDragging(false),
      onDrop: (e) => {
        e.preventDefault();
        setIsDragging(false);
        if (e.dataTransfer.files?.[0]) handleFileUpload(e.dataTransfer.files[0]);
      },
      onClick: () => document.getElementById("file-lesson-input").click(),
      className: `border-2 border-dashed rounded-3xl p-8 sm:p-12 text-center transition-all cursor-pointer ${isDragging ? "border-purple-500 bg-purple-50" : "border-purple-300 bg-purple-50/40 hover:bg-purple-50"}`
    },
    /* @__PURE__ */ React.createElement("input", { id: "file-lesson-input", type: "file", accept: ".docx", className: "hidden", onChange: (e) => e.target.files?.[0] && handleFileUpload(e.target.files[0]) }),
    /* @__PURE__ */ React.createElement("div", { className: "w-12 h-12 rounded-full bg-white text-purple-600 flex items-center justify-center text-2xl mx-auto mb-3 shadow-sm" }, "\u{1F4E4}"),
    /* @__PURE__ */ React.createElement("div", { className: "inline-block px-5 py-2.5 bg-purple-600 text-white font-bold text-sm rounded-xl shadow-sm mb-2" }, "Ch\u1ECDn file gi\xE1o \xE1n Word (.docx)"),
    /* @__PURE__ */ React.createElement("p", { className: "text-xs font-semibold text-slate-700" }, "ho\u1EB7c k\xE9o th\u1EA3 file v\xE0o khung n\xE0y"),
    /* @__PURE__ */ React.createElement("p", { className: "text-xs text-slate-500 mt-1" }, "1 b\xE0i ho\u1EB7c c\u1EA3 tu\u1EA7n (t\u1ED1i \u0111a kho\u1EA3ng 100 trang). File .doc c\u0169: m\u1EDF b\u1EB1ng Word r\u1ED3i L\u01B0u th\xE0nh .docx.")
  ) : /* @__PURE__ */ React.createElement("div", { className: "bg-emerald-50/60 rounded-2xl p-5 border-2 border-dashed border-emerald-400 text-center" }, /* @__PURE__ */ React.createElement("h4", { className: "font-bold text-slate-900 text-base" }, docInfo.fileName), /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-center gap-3 flex-wrap mt-2 text-xs font-semibold text-slate-600" }, /* @__PURE__ */ React.createElement("span", null, docInfo.paraCount, " \u0111o\u1EA1n"), " \u2022 ", /* @__PURE__ */ React.createElement("span", null, docInfo.tables, " b\u1EA3ng"), " \u2022 ", /* @__PURE__ */ React.createElement("span", null, docInfo.images, " h\xECnh"), " \u2022 ", /* @__PURE__ */ React.createElement("span", null, docInfo.maths, " c\xF4ng th\u1EE9c")), /* @__PURE__ */ React.createElement("button", { onClick: () => {
    setDocInfo(null);
    setItems([]);
  }, className: "mt-3 text-xs font-bold text-blue-700 bg-white border border-blue-200 hover:bg-blue-50 px-4 py-1.5 rounded-xl" }, "\u{1F4E4} CH\u1ECCN FILE KH\xC1C")), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-1 sm:grid-cols-3 gap-4" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-slate-600 mb-1 uppercase" }, "L\u1EDBp"), /* @__PURE__ */ React.createElement("select", { value: grade, onChange: (e) => setGrade(Number(e.target.value)), className: "w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold" }, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((g) => /* @__PURE__ */ React.createElement("option", { key: g, value: g }, "L\u1EDBp ", g)))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-slate-600 mb-1 uppercase" }, "M\xF4n h\u1ECDc (Kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement("select", { value: subject, onChange: (e) => setSubject(e.target.value), className: "w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold" }, SUBJECTS.map((s) => /* @__PURE__ */ React.createElement("option", { key: s, value: s }, s)))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-slate-600 mb-1 uppercase" }, "T\xEAn b\xE0i (Kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement("input", { type: "text", value: title, onChange: (e) => setTitle(e.target.value), placeholder: "VD: B\xC0I 2: T\u1EACP H\u1EE2P", className: "w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold" }))), !docInfo && /* @__PURE__ */ React.createElement("div", { className: "bg-red-50 border border-red-200 text-red-700 rounded-2xl p-4 text-xs font-medium" }, /* @__PURE__ */ React.createElement("strong", null, 'C\u1EA7n b\u1ED5 sung tr\u01B0\u1EDBc khi sang b\u01B0\u1EDBc "N\u1ED9i dung t\xEDch h\u1EE3p":'), /* @__PURE__ */ React.createElement("ul", { className: "list-disc ml-5 mt-1" }, /* @__PURE__ */ React.createElement("li", null, "Ch\u1ECDn file gi\xE1o \xE1n Word (.docx)."))), /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap justify-end gap-3" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      disabled: !docInfo,
      onClick: () => handleRunAi(true),
      className: "px-6 py-3 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 disabled:from-slate-300 disabled:to-slate-300 text-white font-extrabold rounded-xl text-sm shadow-sm"
    },
    "\u2728 T\u1EF0 \u0110\u1ED8NG NH\u1EACN DI\u1EC6N & T\xCDCH H\u1EE2P NGAY"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      disabled: !docInfo,
      onClick: () => setStep(1),
      className: "px-6 py-3 bg-purple-600 hover:bg-purple-700 disabled:bg-slate-300 text-white font-bold rounded-xl text-sm shadow-sm"
    },
    "TI\u1EBEP: N\u1ED8I DUNG T\xCDCH H\u1EE2P \u2192"
  ))), step === 0 && mode === "new" && /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("h2", { className: "text-lg font-bold text-slate-800 flex items-center gap-2" }, /* @__PURE__ */ React.createElement("span", null, "\u{1F4D6}"), " B\u01B0\u1EDBc 1: Ch\u1ECDn S\xE1ch gi\xE1o khoa & Trang b\xE0i d\u1EA1y"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-slate-500" }, "T\u1EA3i file SGK (.pdf) ho\u1EB7c \u1EA3nh ch\u1EE5p trang s\xE1ch. SGK \u0111\u01B0\u1EE3c l\u01B0u tr\xEAn m\xE1y \u0111\u1EC3 d\xF9ng l\u1EA1i nhanh.")), savedTextbooks.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "bg-slate-50 rounded-2xl p-4 border border-slate-200" }, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-bold text-slate-600 uppercase mb-2.5" }, "\u{1F4DA} SGK \u0111\xE3 c\xF3 tr\xEAn m\xE1y"), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-2.5" }, savedTextbooks.map((tb) => /* @__PURE__ */ React.createElement("div", { key: tb.id, className: "bg-white p-3 rounded-xl border border-slate-200 flex items-center justify-between gap-2" }, /* @__PURE__ */ React.createElement("div", { className: "min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-bold text-slate-800 truncate" }, tb.name), /* @__PURE__ */ React.createElement("div", { className: "text-[11px] text-slate-500" }, tb.numPages, " trang \u2022 L\u1EDBp ", tb.grade, " \u2022 ", tb.subject)), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1.5 shrink-0" }, /* @__PURE__ */ React.createElement("button", { onClick: () => handleUseSavedSgk(tb), className: "px-2.5 py-1 bg-purple-600 text-white text-xs font-bold rounded-lg" }, "D\xD9NG"), /* @__PURE__ */ React.createElement("button", { onClick: () => deleteTextbookFromDb(tb.id).then(() => listSavedTextbooks().then(setSavedTextbooks)), className: "px-2 py-1 text-slate-400 hover:text-red-600 text-xs font-bold" }, "\u2715")))))), /* @__PURE__ */ React.createElement("div", { className: "border-2 border-dashed border-purple-300 bg-purple-50/40 hover:bg-purple-50 rounded-2xl p-6 text-center cursor-pointer", onClick: () => document.getElementById("sgk-file-input").click() }, /* @__PURE__ */ React.createElement("input", { id: "sgk-file-input", type: "file", accept: ".pdf,image/*", multiple: true, className: "hidden", onChange: (e) => handleSgkUpload(e.target.files) }), /* @__PURE__ */ React.createElement("div", { className: "text-2xl mb-1" }, "\u{1F4E4}"), /* @__PURE__ */ React.createElement("div", { className: "text-sm font-bold text-slate-800" }, sgkFile ? `\u0110ang ch\u1ECDn: ${sgkFile.name} (${sgkFile.numPages} trang)` : "T\u1EA3i SGK kh\xE1c (file PDF ho\u1EB7c \u1EA3nh ch\u1EE5p)")), sgkFile && sgkFile.type === "pdf" && /* @__PURE__ */ React.createElement("div", { className: "bg-purple-50/60 border border-purple-200 rounded-2xl p-4 flex flex-wrap items-end gap-3" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-purple-900 mb-1" }, "T\u1EEB trang"), /* @__PURE__ */ React.createElement("input", { type: "number", min: 1, max: sgkFile.numPages, value: pageFrom, onChange: (e) => setPageFrom(Number(e.target.value)), className: "w-24 p-2 bg-white border border-purple-300 rounded-xl text-sm font-bold text-center" })), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-purple-900 mb-1" }, "\u0110\u1EBFn trang (t\u1ED1i \u0111a 20)"), /* @__PURE__ */ React.createElement("input", { type: "number", min: 1, max: sgkFile.numPages, value: pageTo, onChange: (e) => setPageTo(Number(e.target.value)), className: "w-24 p-2 bg-white border border-purple-300 rounded-xl text-sm font-bold text-center" })), /* @__PURE__ */ React.createElement("button", { onClick: () => renderPdfRange(sgkFile.pdfDoc, pageFrom, pageTo), disabled: renderingPages, className: "px-5 py-2.5 bg-purple-600 text-white font-bold text-sm rounded-xl" }, renderingPages ? "\u0110ang t\u1EA3i..." : "OK \u2022 Xem trang")), pagePreviews.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-2 sm:grid-cols-4 gap-3" }, pagePreviews.map((p) => /* @__PURE__ */ React.createElement("div", { key: p.pageNum, className: "border border-slate-200 rounded-xl overflow-hidden bg-white" }, /* @__PURE__ */ React.createElement("img", { src: p.dataUrl, alt: `Trang ${p.pageNum}`, className: "w-full h-40 object-cover object-top" }), /* @__PURE__ */ React.createElement("div", { className: "text-center py-1 text-xs font-bold text-slate-600 bg-slate-50 border-t" }, "Trang ", p.pageNum)))), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-1 sm:grid-cols-4 gap-4 pt-4 border-t" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-slate-600 mb-1 uppercase" }, "L\u1EDBp"), /* @__PURE__ */ React.createElement("select", { value: grade, onChange: (e) => setGrade(Number(e.target.value)), className: "w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold" }, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((g) => /* @__PURE__ */ React.createElement("option", { key: g, value: g }, "L\u1EDBp ", g)))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-slate-600 mb-1 uppercase" }, "M\xF4n h\u1ECDc"), /* @__PURE__ */ React.createElement("select", { value: subject, onChange: (e) => setSubject(e.target.value), className: "w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold" }, SUBJECTS.map((s) => /* @__PURE__ */ React.createElement("option", { key: s, value: s }, s)))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-slate-600 mb-1 uppercase" }, "S\u1ED1 ti\u1EBFt"), /* @__PURE__ */ React.createElement("select", { value: periods, onChange: (e) => setPeriods(Number(e.target.value)), className: "w-full p-2.5 bg-white border border-slate-300 rounded-xl text-sm font-semibold" }, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((p) => /* @__PURE__ */ React.createElement("option", { key: p, value: p }, p, " ti\u1EBFt")))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-xs font-bold text-slate-600 mb-1 uppercase" }, "T\xEAn b\xE0i ", /* @__PURE__ */ React.createElement("span", { className: "text-red-500" }, "*")), /* @__PURE__ */ React.createElement("input", { type: "text", value: title, onChange: (e) => setTitle(e.target.value), placeholder: "VD: B\xE0i 1...", className: `w-full p-2.5 bg-white border rounded-xl text-sm font-semibold ${!title.trim() ? "border-red-300" : "border-slate-300"}` }))), /* @__PURE__ */ React.createElement("div", { className: "space-y-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-bold text-slate-700" }, "K\u1EBF ho\u1EA1ch ho\u1EA1t \u0111\u1ED9ng (kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement("div", { className: "text-[11px] text-slate-500" }, "\u0110\u1EC3 tr\u1ED1ng = AI t\u1EF1 thi\u1EBFt k\u1EBF theo n\u1ED9i dung b\xE0i. Ghi ho\u1EA1t \u0111\u1ED9ng th\xEC AI so\u1EA1n \u0111\xFAng theo th\u1EE9 t\u1EF1 \u0111\xF3."), activityList.map((act, idx) => /* @__PURE__ */ React.createElement("div", { key: idx, className: "flex items-center gap-2" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "text",
      value: act,
      onChange: (e) => setActivityList(activityList.map((x, i) => i === idx ? e.target.value : x)),
      placeholder: `Ho\u1EA1t \u0111\u1ED9ng ${idx + 1}`,
      className: "flex-1 p-2 bg-white border border-slate-300 rounded-xl text-xs"
    }
  ), /* @__PURE__ */ React.createElement("button", { type: "button", onClick: () => setActivityList(activityList.filter((_, i) => i !== idx)), className: "text-slate-400 hover:text-red-600 px-2 font-bold" }, "\u{1F5D1}"))), /* @__PURE__ */ React.createElement(
    "button",
    {
      type: "button",
      onClick: () => setActivityList([...activityList, ""]),
      className: "text-xs font-extrabold text-blue-700 hover:underline"
    },
    "+ TH\xCAM HO\u1EA0T \u0110\u1ED8NG"
  )), (pagePreviews.length === 0 || !title.trim()) && /* @__PURE__ */ React.createElement("div", { className: "bg-red-50 border border-red-200 text-red-700 rounded-2xl p-4 text-xs font-medium" }, /* @__PURE__ */ React.createElement("strong", null, 'C\u1EA7n b\u1ED5 sung tr\u01B0\u1EDBc khi sang b\u01B0\u1EDBc "N\u1ED9i dung t\xEDch h\u1EE3p":'), /* @__PURE__ */ React.createElement("ul", { className: "list-disc ml-5 mt-1" }, pagePreviews.length === 0 && /* @__PURE__ */ React.createElement("li", null, "Ch\u1ECDn SGK v\xE0 c\xE1c trang c\u1EE7a b\xE0i (file PDF ho\u1EB7c \u1EA3nh ch\u1EE5p)."), !title.trim() && /* @__PURE__ */ React.createElement("li", null, "Nh\u1EADp t\xEAn b\xE0i."))), /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap justify-end gap-3" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      disabled: pagePreviews.length === 0 || !title.trim(),
      onClick: () => handleRunAi(true),
      className: "px-6 py-3 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 disabled:from-slate-300 disabled:to-slate-300 text-white font-extrabold rounded-xl text-sm shadow-sm"
    },
    "\u2728 T\u1EF0 \u0110\u1ED8NG NH\u1EACN DI\u1EC6N & SO\u1EA0N GI\xC1O \xC1N NGAY"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      disabled: pagePreviews.length === 0 || !title.trim(),
      onClick: () => setStep(1),
      className: "px-6 py-3 bg-purple-600 hover:bg-purple-700 disabled:bg-slate-300 text-white font-bold rounded-xl text-sm"
    },
    "TI\u1EBEP: N\u1ED8I DUNG T\xCDCH H\u1EE2P \u2192"
  ))), step === 1 && /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-3" }, /* @__PURE__ */ React.createElement("div", { className: "w-10 h-10 rounded-xl bg-purple-600 text-white flex items-center justify-center text-lg font-bold" }, "\u{1F4DA}"), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("h2", { className: "text-lg font-bold text-slate-900" }, "N\u1ED9i dung t\xEDch h\u1EE3p"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-slate-500" }, "Ch\u1ECDn c\xE1c n\u1ED9i dung mu\u1ED1n l\u1ED3ng gh\xE9p v\xE0o gi\xE1o \xE1n"))), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-1 lg:grid-cols-12 gap-5 items-start" }, /* @__PURE__ */ React.createElement("div", { className: "lg:col-span-5 space-y-3" }, INTEGRATION_CARDS.map((card) => {
    const checked = !!opts[card.key];
    return /* @__PURE__ */ React.createElement(
      "div",
      {
        key: card.key,
        onClick: () => setOpts((prev) => ({ ...prev, [card.key]: !prev[card.key] })),
        style: { borderColor: checked ? card.dot : "#e2e8f0", backgroundColor: checked ? `${card.dot}0A` : "#fff" },
        className: "p-3.5 rounded-2xl border-2 cursor-pointer transition-all hover:shadow-xs"
      },
      /* @__PURE__ */ React.createElement("div", { className: "flex items-start gap-2.5" }, /* @__PURE__ */ React.createElement(
        "input",
        {
          type: "checkbox",
          checked,
          onChange: () => {
          },
          style: { accentColor: card.dot },
          className: "w-4 h-4 mt-0.5 rounded cursor-pointer shrink-0"
        }
      ), /* @__PURE__ */ React.createElement("div", { className: "min-w-0" }, /* @__PURE__ */ React.createElement("div", { style: { color: card.dot }, className: "font-extrabold text-sm leading-snug" }, card.title), /* @__PURE__ */ React.createElement("div", { className: "text-[11px] text-slate-500 mt-1 leading-relaxed" }, card.desc), checked && /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1.5 mt-2 text-xs font-bold text-slate-700" }, /* @__PURE__ */ React.createElement("span", null, "M\xE0u ch\u1EEF:"), /* @__PURE__ */ React.createElement("span", { style: { backgroundColor: card.dot }, className: "w-4 h-4 rounded-full inline-block border border-white shadow-xs ring-1 ring-slate-300" }))))
    );
  })), /* @__PURE__ */ React.createElement("div", { className: "lg:col-span-7 space-y-4" }, !Object.values(opts).some(Boolean) && /* @__PURE__ */ React.createElement("div", { className: "p-8 rounded-2xl border-2 border-dashed border-slate-200 text-center text-xs text-slate-400" }, "T\xEDch ch\u1ECDn \xEDt nh\u1EA5t m\u1ED9t n\u1ED9i dung \u1EDF c\u1ED9t b\xEAn tr\xE1i \u0111\u1EC3 c\u1EA5u h\xECnh chi ti\u1EBFt."), opts.nls && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-violet-300 bg-violet-50/40 space-y-3" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-slate-800" }, "Ph\xE2n ph\u1ED1i ch\u01B0\u01A1ng tr\xECnh (kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-slate-500 mt-0.5" }, "Ch\u1EC9 t\u1EA3i khi PPCT c\u1EE7a tr\u01B0\u1EDDng C\xD3 c\u1ED9t N\u0103ng l\u1EF1c s\u1ED1 / N\u0103ng l\u1EF1c AI / An to\xE0n m\u1EA1ng. File Excel (.xlsx) ho\u1EB7c Word (.docx)."), /* @__PURE__ */ React.createElement("div", { className: "mt-2 flex items-center gap-2" }, /* @__PURE__ */ React.createElement("label", { className: "px-3.5 py-2 bg-white border border-violet-300 text-violet-700 text-xs font-bold rounded-xl cursor-pointer hover:bg-violet-50" }, "\u{1F4E4} ", ppctName ? `\u0110\xE3 t\u1EA3i: ${ppctName}` : "CH\u1ECCN FILE PPCT", /* @__PURE__ */ React.createElement("input", { type: "file", accept: ".docx,.xlsx", className: "hidden", onChange: (e) => e.target.files?.[0] && handlePpctUpload(e.target.files[0]) })), ppctName && /* @__PURE__ */ React.createElement("button", { onClick: () => {
    setPpctName("");
    setPpctText("");
  }, className: "text-xs text-red-600 font-bold" }, "X\xF3a"))), /* @__PURE__ */ React.createElement("div", { className: "pt-3 border-t border-violet-200/70" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-slate-800" }, "Ch\u1ECDn m\xE3 N\u0103ng l\u1EF1c s\u1ED1 c\u1EE5 th\u1EC3 (kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-slate-500 mb-2" }, "\u0110\u1EC3 tr\u1ED1ng = AI t\u1EF1 ch\u1ECDn m\xE3 ph\xF9 h\u1EE3p trong khung c\u1EE7a l\u1EDBp ", grade, ". Ch\u1ECDn m\xE3 = AI ch\u1EC9 d\xF9ng c\xE1c m\xE3 n\xE0y (t\u1ED1i \u0111a 10)."), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: "",
      onChange: (e) => {
        const code = e.target.value;
        if (!code || selectedNlsItems.some((x) => x.code === code) || selectedNlsItems.length >= 10) return;
        const found = nlsCatalog.find((x) => x.code === code);
        if (found) setSelectedNlsItems([...selectedNlsItems, { ...found }]);
      },
      className: "w-full p-2.5 bg-white border border-violet-300 rounded-xl text-xs font-semibold"
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "M\xE3 N\u0103ng l\u1EF1c s\u1ED1 (", getNlsLevelLabel(grade), ")"),
    nlsCatalog.map((c) => /* @__PURE__ */ React.createElement("option", { key: c.code, value: c.code }, c.code, " \u2014 ", c.desc))
  ), selectedNlsItems.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "mt-2.5 space-y-2" }, selectedNlsItems.map((it, idx) => /* @__PURE__ */ React.createElement("div", { key: it.code, className: "bg-white p-2.5 rounded-xl border border-violet-200" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-between mb-1" }, /* @__PURE__ */ React.createElement("span", { className: "text-xs font-extrabold text-violet-700" }, it.code, " \u2014 Y\xEAu c\u1EA7u c\u1EA7n \u0111\u1EA1t (Bi\u1EC3u hi\u1EC7n)"), /* @__PURE__ */ React.createElement("button", { onClick: () => setSelectedNlsItems(selectedNlsItems.filter((_, i) => i !== idx)), className: "text-xs text-slate-400 hover:text-red-600 font-bold" }, "\u2715")), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "text",
      value: it.desc,
      onChange: (e) => setSelectedNlsItems(selectedNlsItems.map((x, i) => i === idx ? { ...x, desc: e.target.value } : x)),
      className: "w-full p-1.5 border border-slate-200 rounded-lg text-xs"
    }
  )))))), opts.ai && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-red-300 bg-red-50/40 space-y-2.5" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-red-700" }, "Ch\u1ECDn y\xEAu c\u1EA7u N\u0103ng l\u1EF1c AI c\u1EE5 th\u1EC3 (kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-slate-500" }, "\u0110\u1EC3 tr\u1ED1ng = AI t\u1EF1 ch\u1ECDn y\xEAu c\u1EA7u ph\xF9 h\u1EE3p trong khung l\u1EDBp ", grade, ". Ch\u1ECDn m\xE3 = AI ch\u1EC9 d\xF9ng c\xE1c m\xE3 n\xE0y (t\u1ED1i \u0111a 10)."), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: "",
      onChange: (e) => {
        const code = e.target.value;
        if (!code || selectedAiItems.some((x) => x.code === code) || selectedAiItems.length >= 10) return;
        const found = aiCatalog.find((x) => x.code === code);
        if (found) setSelectedAiItems([...selectedAiItems, { ...found }]);
      },
      className: "w-full p-2.5 bg-white border border-red-300 rounded-xl text-xs font-semibold"
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "Y\xEAu c\u1EA7u c\u1EA7n \u0111\u1EA1t N\u0103ng l\u1EF1c AI (l\u1EDBp ", grade, ")"),
    aiCatalog.map((c) => /* @__PURE__ */ React.createElement("option", { key: c.code, value: c.code }, c.code, " \u2014 ", c.desc))
  ), selectedAiItems.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "space-y-2" }, selectedAiItems.map((it, idx) => /* @__PURE__ */ React.createElement("div", { key: it.code, className: "bg-white p-2.5 rounded-xl border border-red-200" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-between mb-1" }, /* @__PURE__ */ React.createElement("span", { className: "text-xs font-extrabold text-red-700" }, it.code, " \u2014 Y\xEAu c\u1EA7u c\u1EA7n \u0111\u1EA1t"), /* @__PURE__ */ React.createElement("button", { onClick: () => setSelectedAiItems(selectedAiItems.filter((_, i) => i !== idx)), className: "text-xs text-slate-400 hover:text-red-600 font-bold" }, "\u2715")), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "text",
      value: it.desc,
      onChange: (e) => setSelectedAiItems(selectedAiItems.map((x, i) => i === idx ? { ...x, desc: e.target.value } : x)),
      className: "w-full p-1.5 border border-slate-200 rounded-lg text-xs"
    }
  )))), /* @__PURE__ */ React.createElement("div", { className: "text-[11px] text-slate-500" }, "Ph\xE2n ph\u1ED1i ch\u01B0\u01A1ng tr\xECnh (n\u1EBFu c\xF3) t\u1EA3i \u1EDF khung N\u0103ng l\u1EF1c s\u1ED1 ph\xEDa tr\xEAn \u2014 d\xF9ng chung.")), opts.kt && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-emerald-300 bg-emerald-50/40 space-y-2.5" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-emerald-800" }, "H\u1ECDc sinh khuy\u1EBFt t\u1EADt trong l\u1EDBp"), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-slate-500" }, "Ch\u1ECDn d\u1EA1ng khuy\u1EBFt t\u1EADt (t\u1ED1i \u0111a 5). Ch\u1EC9 ghi t\xEAn l\u1EDBp, KH\xD4NG ghi t\xEAn h\u1ECDc sinh."), /* @__PURE__ */ React.createElement("div", { className: "space-y-2" }, disabilityRows.map((row, idx) => /* @__PURE__ */ React.createElement("div", { key: idx, className: "flex flex-wrap sm:flex-nowrap items-center gap-2" }, /* @__PURE__ */ React.createElement(
    "select",
    {
      value: row.type,
      onChange: (e) => setDisabilityRows(disabilityRows.map((r, i) => i === idx ? { ...r, type: e.target.value } : r)),
      className: "flex-1 p-2 bg-white border border-emerald-300 rounded-xl text-xs font-semibold"
    },
    DISABILITY_TYPES.map((d) => /* @__PURE__ */ React.createElement("option", { key: d.key, value: d.label }, d.label))
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "text",
      value: row.className,
      onChange: (e) => setDisabilityRows(disabilityRows.map((r, i) => i === idx ? { ...r, className: e.target.value } : r)),
      placeholder: "L\u1EDBp (kh\xF4ng b\u1EAFt bu\u1ED9c)",
      className: "w-36 p-2 bg-white border border-slate-300 rounded-xl text-xs"
    }
  ), disabilityRows.length > 1 && /* @__PURE__ */ React.createElement("button", { onClick: () => setDisabilityRows(disabilityRows.filter((_, i) => i !== idx)), className: "text-slate-400 hover:text-red-600 px-2 font-bold" }, "\u{1F5D1}")))), disabilityRows.length < 5 && /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => setDisabilityRows([...disabilityRows, { type: "Khuy\u1EBFt t\u1EADt chung", className: "" }]),
      className: "text-xs font-extrabold text-emerald-700 hover:underline"
    },
    "+ TH\xCAM D\u1EA0NG KHUY\u1EBET T\u1EACT"
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "text",
      value: disabilityNote,
      onChange: (e) => setDisabilityNote(e.target.value),
      placeholder: "Ghi ch\xFA bi\u1EC7n ph\xE1p h\u1ED7 tr\u1EE3 th\xEAm (n\u1EBFu c\xF3)...",
      className: "w-full p-2 bg-white border border-emerald-200 rounded-xl text-xs"
    }
  )), opts.qp && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-blue-300 bg-blue-50/40 space-y-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-blue-800" }, "Gi\xE1o d\u1EE5c qu\u1ED1c ph\xF2ng & an ninh (TT 08/2024)"), /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap gap-1.5" }, GDQPAN_SUGGESTIONS.map((s, i) => /* @__PURE__ */ React.createElement("button", { key: i, type: "button", onClick: () => setQpNote((p) => p ? p + "; " + s : s), className: "text-[11px] bg-white border border-blue-200 text-blue-700 px-2 py-1 rounded-lg" }, "+ ", s))), /* @__PURE__ */ React.createElement("textarea", { value: qpNote, onChange: (e) => setQpNote(e.target.value), rows: 2, placeholder: "Ghi ch\xFA l\u1ED3ng gh\xE9p GDQPAN (kh\xF4ng b\u1EAFt bu\u1ED9c)...", className: "w-full p-2 bg-white border border-blue-200 rounded-xl text-xs" })), opts.bd && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-orange-300 bg-orange-50/40 space-y-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-orange-800" }, "Bi\u1EBFn \u0111\u1ED5i kh\xED h\u1EADu - Ph\xF2ng ch\u1ED1ng thi\xEAn tai"), /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap gap-1.5" }, BDKH_SUGGESTIONS.map((s, i) => /* @__PURE__ */ React.createElement("button", { key: i, type: "button", onClick: () => setBdNote((p) => p ? p + "; " + s : s), className: "text-[11px] bg-white border border-orange-200 text-orange-700 px-2 py-1 rounded-lg" }, "+ ", s))), /* @__PURE__ */ React.createElement("textarea", { value: bdNote, onChange: (e) => setBdNote(e.target.value), rows: 2, placeholder: "Ghi ch\xFA l\u1ED3ng gh\xE9p B\u0110KH-PCTT (kh\xF4ng b\u1EAFt bu\u1ED9c)...", className: "w-full p-2 bg-white border border-orange-200 rounded-xl text-xs" })), opts.ltcm && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-yellow-400 bg-yellow-50/50 space-y-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-yellow-800" }, "Gi\xE1o d\u1EE5c l\xFD t\u01B0\u1EDFng c\xE1ch m\u1EA1ng, \u0111\u1EA1o \u0111\u1EE9c, l\u1ED1i s\u1ED1ng (Q\u0110 2546)"), /* @__PURE__ */ React.createElement("textarea", { value: ltcmNote, onChange: (e) => setLtcmNote(e.target.value), rows: 2, placeholder: "Y\xEAu c\u1EA7u l\u1ED3ng gh\xE9p l\xFD t\u01B0\u1EDFng c\xE1ch m\u1EA1ng, \u0111\u1EA1o \u0111\u1EE9c, l\u1ED1i s\u1ED1ng (kh\xF4ng b\u1EAFt bu\u1ED9c)...", className: "w-full p-2 bg-white border border-yellow-300 rounded-xl text-xs" })), opts.atm && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-teal-300 bg-teal-50/40 space-y-2.5" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-teal-800" }, "Ch\u1ECDn m\xE3 An to\xE0n m\u1EA1ng c\u1EE5 th\u1EC3 (kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-slate-500" }, "\u0110\u1EC3 tr\u1ED1ng = AI t\u1EF1 ch\u1ECDn 1\u20132 m\xE3 ph\xF9 h\u1EE3p v\u1EDBi b\xE0i. Ch\u1ECDn m\xE3 = AI ch\u1EC9 d\xF9ng c\xE1c m\xE3 n\xE0y (t\u1ED1i \u0111a 10)."), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: "",
      onChange: (e) => {
        const code = e.target.value;
        if (!code || selectedAtmItems.some((x) => x.code === code) || selectedAtmItems.length >= 10) return;
        const found = ATM_CATALOG.find((x) => x.code === code);
        if (found) setSelectedAtmItems([...selectedAtmItems, { ...found }]);
      },
      className: "w-full p-2.5 bg-white border border-teal-300 rounded-xl text-xs font-semibold"
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "M\xE3 An to\xE0n m\u1EA1ng (CV 6206)"),
    ATM_CATALOG.map((c) => /* @__PURE__ */ React.createElement("option", { key: c.code, value: c.code }, c.code, " \u2014 ", c.title))
  ), selectedAtmItems.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "space-y-1.5" }, selectedAtmItems.map((it, idx) => /* @__PURE__ */ React.createElement("div", { key: it.code, className: "bg-white p-2 rounded-xl border border-teal-200 flex items-start justify-between gap-2 text-xs" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("span", { className: "font-extrabold text-teal-700" }, it.code, " \u2014 ", it.title), /* @__PURE__ */ React.createElement("div", { className: "text-[11px] text-slate-500 mt-0.5" }, "Li\xEAn k\u1EBFt TT 02/2025: ", /* @__PURE__ */ React.createElement("strong", null, "B\u1EAFt bu\u1ED9c"), " (", (it.requiredNls || []).join(", "), ") \u2022 ", /* @__PURE__ */ React.createElement("strong", null, "Khuy\u1EBFn kh\xEDch"), " (", (it.recommendedNls || []).join(", "), ")")), /* @__PURE__ */ React.createElement("button", { onClick: () => setSelectedAtmItems(selectedAtmItems.filter((_, i) => i !== idx)), className: "text-slate-400 hover:text-red-600 font-bold" }, "\u2715")))), /* @__PURE__ */ React.createElement("div", { className: "text-[11px] text-slate-500" }, "Ph\xE2n ph\u1ED1i ch\u01B0\u01A1ng tr\xECnh (n\u1EBFu c\xF3) t\u1EA3i \u1EDF khung N\u0103ng l\u1EF1c s\u1ED1 ph\xEDa tr\xEAn \u2014 d\xF9ng chung.")), opts.hcm && /* @__PURE__ */ React.createElement("div", { className: "p-4 rounded-2xl border border-pink-300 bg-pink-50/40 space-y-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-pink-800" }, "Gi\xE1o d\u1EE5c t\u01B0 t\u01B0\u1EDFng, \u0111\u1EA1o \u0111\u1EE9c, phong c\xE1ch H\u1ED3 Ch\xED Minh (Q\u0110 2995)"), /* @__PURE__ */ React.createElement("textarea", { value: hcmNote, onChange: (e) => setHcmNote(e.target.value), rows: 2, placeholder: "Ghi ch\xFA n\u1ED9i dung h\u1ECDc t\u1EADp v\xE0 l\xE0m theo t\u01B0 t\u01B0\u1EDFng, \u0111\u1EA1o \u0111\u1EE9c, phong c\xE1ch HCM (kh\xF4ng b\u1EAFt bu\u1ED9c)...", className: "w-full p-2 bg-white border border-pink-200 rounded-xl text-xs" })))), /* @__PURE__ */ React.createElement("div", { className: "pt-4 border-t border-slate-200 space-y-4" }, /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap items-center justify-between gap-4" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-bold text-slate-700 mb-1.5" }, "Ng\xF4n ng\u1EEF hi\u1EC3n th\u1ECB c\u1EE7a gi\xE1o \xE1n"), /* @__PURE__ */ React.createElement("div", { className: "flex gap-2" }, ["Ti\u1EBFng Vi\u1EC7t", "Ti\u1EBFng Anh"].map((lang) => /* @__PURE__ */ React.createElement(
    "button",
    {
      key: lang,
      type: "button",
      onClick: () => setLessonLang(lang),
      className: `px-4 py-1.5 rounded-full text-xs font-bold border transition-all ${lessonLang === lang ? "bg-purple-600 text-white border-purple-600" : "bg-white text-slate-600 border-slate-300"}`
    },
    lang
  )))), /* @__PURE__ */ React.createElement("label", { className: "flex items-center gap-2.5 px-4 py-2.5 rounded-full bg-purple-50 border border-purple-200 cursor-pointer" }, /* @__PURE__ */ React.createElement("input", { type: "checkbox", checked: strongerAi, onChange: (e) => setStrongerAi(e.target.checked), className: "w-4 h-4 accent-purple-600" }), /* @__PURE__ */ React.createElement("span", { className: "text-xs font-extrabold text-purple-900" }, "\u{1F9E0} AI m\u1EA1nh h\u01A1n (Model Pro)"))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-bold text-slate-700 mb-1" }, "\u0110i\u1EC1u ch\u1EC9nh th\xEAm (kh\xF4ng b\u1EAFt bu\u1ED9c)"), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: extraInstructions,
      maxLength: 3e3,
      onChange: (e) => setExtraInstructions(e.target.value),
      rows: 2,
      placeholder: `N\u1ED9i dung \u0111i\u1EC1u ch\u1EC9nh (${extraInstructions.length}/3.000 k\xFD t\u1EF1) \u2014 M\xF4 t\u1EA3 \u0111i\u1EC1u ki\u1EC7n th\u1EF1c t\u1EBF c\u1EE7a l\u1EDBp ho\u1EB7c mong mu\u1ED1n ri\xEAng`,
      className: "w-full p-3 bg-white border border-slate-300 rounded-xl text-xs outline-none focus:ring-2 focus:ring-purple-400"
    }
  ))), /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap justify-between items-center gap-3 pt-4 border-t border-slate-200" }, /* @__PURE__ */ React.createElement("button", { onClick: () => setStep(0), className: "px-4 py-2.5 text-slate-600 hover:text-slate-900 font-bold text-sm" }, "\u2190 GI\xC1O \xC1N"), /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap items-center gap-3" }, /* @__PURE__ */ React.createElement("button", { onClick: () => handleRunAi(true), className: "px-5 py-3.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-extrabold rounded-xl text-sm shadow-md" }, "\u2728 AI T\u1EF0 CH\u1ECCN & T\xCDCH H\u1EE2P"), /* @__PURE__ */ React.createElement("button", { onClick: () => handleRunAi(false), className: "px-7 py-3.5 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-extrabold rounded-xl text-sm shadow-md" }, mode === "existing" ? "\u26A1 T\xCDCH H\u1EE2P V\xC0O GI\xC1O \xC1N \u2192" : "\u26A1 T\u1EA0O GI\xC1O \xC1N M\u1EDAI \u2192")))), step === 2 && /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-3xl p-6 sm:p-8 shadow-sm border border-slate-200/80 space-y-6" }, /* @__PURE__ */ React.createElement("div", { className: "flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("h2", { className: "text-lg font-bold text-slate-900" }, "\u{1F4E5} ", mode === "existing" ? "T\u1EA3i Word \u0111\xE3 t\xEDch h\u1EE3p" : "Xem \xB7 T\u1EA3i Word"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-slate-500 mt-0.5" }, "Ki\u1EC3m tra b\xE1o c\xE1o t\xEDch h\u1EE3p, s\u1EEDa n\u1ED9i dung n\u1EBFu mu\u1ED1n v\xE0 t\u1EA3i file .docx.")), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2" }, /* @__PURE__ */ React.createElement("button", { onClick: handleCopyReport, className: "px-3.5 py-2 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-300 font-bold text-xs rounded-xl" }, copiedReport ? "\u2713 \u0110\xE3 sao ch\xE9p!" : "\u{1F4CB} Sao ch\xE9p b\xE1o c\xE1o"), /* @__PURE__ */ React.createElement("button", { onClick: () => setStep(1), className: "px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs rounded-xl" }, "\u{1F504} T\xEDch h\u1EE3p l\u1EA1i"))), /* @__PURE__ */ React.createElement("div", { className: "bg-slate-50 border border-slate-200 rounded-2xl p-4" }, /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap items-center justify-between gap-2 mb-3" }, /* @__PURE__ */ React.createElement("div", { className: "text-xs font-extrabold text-slate-800 uppercase" }, "\u{1F4CA} B\xE1o c\xE1o t\xEDch h\u1EE3p (", integrationSummary.total, " v\u1ECB tr\xED)"), ppctName && /* @__PURE__ */ React.createElement("span", { className: "text-[11px] font-semibold text-violet-700 bg-violet-100 px-2.5 py-0.5 rounded-full" }, "PPCT: ", ppctName)), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs" }, Object.entries(CATEGORIES).map(([k, cat]) => {
    const cnt = (integrationSummary.byCat[k] || []).length;
    if (cnt === 0 && !opts[k]) return null;
    return /* @__PURE__ */ React.createElement("div", { key: k, style: { borderColor: cat.border }, className: "bg-white p-2.5 rounded-xl border" }, /* @__PURE__ */ React.createElement("div", { style: { color: cat.textCol }, className: "font-bold" }, cat.label, ": ", cnt, " v\u1ECB tr\xED"));
  }))), /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap items-center justify-between gap-3 py-2 border-b border-slate-100" }, mode === "existing" ? /* @__PURE__ */ React.createElement("div", { className: "flex flex-wrap items-center gap-1.5" }, /* @__PURE__ */ React.createElement("button", { onClick: () => setActiveFilter("all"), className: `px-3 py-1.5 rounded-xl text-xs font-bold ${activeFilter === "all" ? "bg-purple-600 text-white" : "bg-slate-100 text-slate-600"}` }, "T\u1EA5t c\u1EA3 (", items.length, ")"), Object.entries(CATEGORIES).map(([k, cat]) => {
    const cnt = items.filter((i) => i.k === k).length;
    if (cnt === 0) return null;
    return /* @__PURE__ */ React.createElement("button", { key: k, onClick: () => setActiveFilter(k), className: `px-3 py-1.5 rounded-xl text-xs font-bold ${activeFilter === k ? "bg-purple-600 text-white" : "bg-slate-100 text-slate-600"}` }, cat.label, " (", cnt, ")");
  })) : /* @__PURE__ */ React.createElement("div", { className: "text-sm font-extrabold text-slate-800" }, newLessonData?.title), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-3" }, mode === "existing" && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("button", { onClick: () => setItems((prev) => prev.map((it) => ({ ...it, on: true }))), className: "text-xs font-bold text-purple-700 hover:underline" }, "Ch\u1ECDn h\u1EBFt"), /* @__PURE__ */ React.createElement("button", { onClick: () => setItems((prev) => prev.map((it) => ({ ...it, on: false }))), className: "text-xs font-bold text-slate-500 hover:underline" }, "B\u1ECF h\u1EBFt")), /* @__PURE__ */ React.createElement("label", { className: "flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-700" }, /* @__PURE__ */ React.createElement("input", { type: "checkbox", checked: colorEnabled, onChange: (e) => setColorEnabled(e.target.checked), className: "w-4 h-4 accent-purple-600 rounded" }), /* @__PURE__ */ React.createElement("span", null, "T\xF4 m\xE0u ph\u1EA7n t\xEDch h\u1EE3p")))), mode === "existing" && /* @__PURE__ */ React.createElement("div", { className: "space-y-3" }, filteredItems.map((item) => {
    const cat = CATEGORIES[item.k] || CATEGORIES.nls;
    return /* @__PURE__ */ React.createElement("div", { key: item.id, className: `p-4 rounded-2xl border ${item.on ? "bg-white border-slate-300" : "bg-slate-50/60 border-slate-200 opacity-60"}` }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2 mb-2" }, /* @__PURE__ */ React.createElement("input", { type: "checkbox", checked: item.on, onChange: (e) => setItems((prev) => prev.map((it) => it.id === item.id ? { ...it, on: e.target.checked } : it)), className: "w-4 h-4 accent-purple-600 rounded" }), /* @__PURE__ */ React.createElement("span", { style: { backgroundColor: cat.bg, color: cat.textCol, borderColor: cat.border }, className: "px-2.5 py-0.5 rounded-lg border text-xs font-extrabold uppercase" }, cat.label), /* @__PURE__ */ React.createElement("span", { className: "text-xs text-slate-400 font-medium" }, "V\u1ECB tr\xED: ", item.a, " \u2022 ", item.w === "end" ? "n\u1ED1i cu\u1ED1i \u0111o\u1EA1n" : "d\xF2ng m\u1EDBi b\xEAn d\u01B0\u1EDBi")), /* @__PURE__ */ React.createElement(
      "textarea",
      {
        value: item.t,
        onChange: (e) => setItems((prev) => prev.map((it) => it.id === item.id ? { ...it, t: e.target.value } : it)),
        rows: 2,
        style: { color: colorEnabled ? cat.textCol : "#1e293b" },
        className: "w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium outline-none focus:bg-white focus:ring-2 focus:ring-purple-400"
      }
    ));
  })), mode === "new" && newLessonData && /* @__PURE__ */ React.createElement("div", { className: "space-y-4" }, newLessonData.sections.map((sec, sIdx) => /* @__PURE__ */ React.createElement("div", { key: sec.id, className: "border border-slate-200 rounded-2xl p-4 bg-slate-50/40" }, /* @__PURE__ */ React.createElement("h3", { className: "font-extrabold text-sm text-slate-900 mb-3" }, sec.heading), /* @__PURE__ */ React.createElement("div", { className: "space-y-2" }, sec.blocks.map((blk, bIdx) => {
    const cat = blk.type !== "normal" ? CATEGORIES[blk.type] : null;
    return /* @__PURE__ */ React.createElement("div", { key: blk.id, className: "flex items-start gap-2" }, cat && /* @__PURE__ */ React.createElement("span", { style: { backgroundColor: cat.bg, color: cat.textCol, borderColor: cat.border }, className: "px-2 py-0.5 rounded border text-[10px] font-extrabold uppercase shrink-0 mt-1" }, cat.label), /* @__PURE__ */ React.createElement(
      "textarea",
      {
        value: blk.text,
        rows: Math.max(1, Math.ceil(blk.text.length / 100)),
        onChange: (e) => {
          const val = e.target.value;
          setNewLessonData((prev) => ({
            ...prev,
            sections: prev.sections.map((s, si) => si !== sIdx ? s : {
              ...s,
              blocks: s.blocks.map((b, bi) => bi !== bIdx ? b : { ...b, text: val })
            })
          }));
        },
        style: { color: cat && colorEnabled ? cat.textCol : "#1e293b" },
        className: "w-full p-2 bg-white border border-slate-200 rounded-xl text-xs font-medium outline-none focus:ring-2 focus:ring-purple-400"
      }
    ));
  }))))), /* @__PURE__ */ React.createElement("div", { className: "sticky bottom-4 bg-slate-900 text-white rounded-2xl p-4 shadow-2xl flex flex-col sm:flex-row items-center justify-between gap-4 z-40" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-sm font-bold" }, "\u{1F4E5} ", mode === "existing" ? `\u0110\xE3 duy\u1EC7t ${activeCount} / ${items.length} \u0111o\u1EA1n ch\xE8n` : `Gi\xE1o \xE1n ${subject} L\u1EDBp ${grade} (${periods} ti\u1EBFt)`), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-slate-400 mt-0.5" }, "Xu\u1EA5t file Word (.docx) chu\u1EA9n \u0111\u1ECBnh d\u1EA1ng, gi\u1EEF nguy\xEAn m\xE0u ph\xE2n bi\u1EC7t.")), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-3 w-full sm:w-auto" }, /* @__PURE__ */ React.createElement("button", { onClick: () => {
    setStep(0);
    setDocInfo(null);
    setItems([]);
    setNewLessonData(null);
  }, className: "px-4 py-3 bg-white/10 hover:bg-white/20 text-white font-bold text-xs rounded-xl" }, "\u{1F504} ", mode === "existing" ? "Gi\xE1o \xE1n kh\xE1c" : "So\u1EA1n b\xE0i kh\xE1c"), /* @__PURE__ */ React.createElement("button", { onClick: handleDownloadWord, disabled: downloading || mode === "existing" && activeCount === 0, className: "flex-1 sm:flex-none px-6 py-3 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-700 text-white font-black text-sm rounded-xl shadow-lg" }, downloading ? "\u0110ang t\u1EA1o file..." : "\u{1F4E5} T\u1EA3i Word (.docx)"))))), loading && /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 bg-slate-900/70 backdrop-blur-sm z-50 flex items-center justify-center p-4" }, /* @__PURE__ */ React.createElement("div", { className: "bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl text-center" }, /* @__PURE__ */ React.createElement("div", { className: "w-16 h-16 rounded-2xl bg-purple-100 text-purple-600 flex items-center justify-center text-3xl mx-auto mb-4 animate-bounce" }, "\u2728"), /* @__PURE__ */ React.createElement("h3", { className: "text-lg font-black text-slate-900 mb-1" }, "AI \u0110ang X\u1EED L\xFD T\xEDch H\u1EE3p"), /* @__PURE__ */ React.createElement("p", { className: "text-xs text-slate-500 mb-6 font-medium" }, progressStatus), /* @__PURE__ */ React.createElement("div", { className: "w-full bg-slate-100 h-3 rounded-full overflow-hidden mb-3" }, /* @__PURE__ */ React.createElement("div", { className: "bg-gradient-to-r from-purple-600 to-indigo-600 h-full rounded-full transition-all duration-300", style: { width: `${progress}%` } })), /* @__PURE__ */ React.createElement("div", { className: "flex justify-between text-xs text-slate-400 font-bold" }, /* @__PURE__ */ React.createElement("span", null, "\u0110ang x\u1EED l\xFD..."), /* @__PURE__ */ React.createElement("span", null, progress, "%")))));
}
var rootEl = document.getElementById("root");
if (rootEl) {
  const root = ReactDOM.createRoot(rootEl);
  root.render(React.createElement(App));
}
export {
  App
};
