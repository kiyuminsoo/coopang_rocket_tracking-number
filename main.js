/* 모든 로직은 브라우저에서 동작합니다. */

const pdfInput = document.getElementById("pdfInput");
const excelInput = document.getElementById("excelInput");
const parseBtn = document.getElementById("parseBtn");
const copyBtn = document.getElementById("copyBtn");
const showFcToggle = document.getElementById("showFcToggle");
const statusEl = document.getElementById("status");
const errorBox = document.getElementById("errorBox");
const resultSection = document.getElementById("resultSection");
const resultText = document.getElementById("resultText");
const excelOptions = document.getElementById("excelOptions");
const sheetSelect = document.getElementById("sheetSelect");
const headerRowInput = document.getElementById("headerRowInput");
const columnSelectField = document.getElementById("columnSelectField");
const columnSelect = document.getElementById("columnSelect");
const ctColumnSelectField = document.getElementById("ctColumnSelectField");
const ctColumnSelect = document.getElementById("ctColumnSelect");

const PDF_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const FC_HEADER_KEYWORDS = ["물류센터", "센터", "FC", "받는 사람", "수취인", "배송처", "납품처"];
const CT_HEADER_KEYWORDS = ["C/T NO", "C/T NO.", "CT NO", "CTNO", "CT NO.", "C/T", "CT", "C/TNO"];

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;
}

function normalizeWhitespace(value) {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeFc(value) {
  let text = normalizeWhitespace(String(value ?? ""));
  if (!text) return "";
  text = text.replace(/\(\d+\)\s*$/g, "").trim();
  text = text.replace(/\s*FC\s*$/i, "").trim();
  return text;
}

function normalizeHeader(value) {
  return normalizeWhitespace(String(value ?? ""))
    .toLowerCase()
    .replace(/[\s\.\-_/]/g, "");
}

function setStatus(message) {
  statusEl.textContent = message;
}

function showErrors(messages) {
  if (!messages.length) {
    errorBox.hidden = true;
    errorBox.innerHTML = "";
    return;
  }
  errorBox.hidden = false;
  errorBox.innerHTML = messages.map((msg) => `• ${msg}`).join("<br>");
}

function resetOutput() {
  resultSection.hidden = true;
  resultText.value = "";
  copyBtn.disabled = true;
}

const excelState = {
  workbook: null,
  sheetNames: [],
  selectedSheet: "",
  headerRowIndex: 0,
  headerColIndex: -1,
  ctHeaderColIndex: -1,
  rows: []
};

function clearExcelState() {
  excelState.workbook = null;
  excelState.sheetNames = [];
  excelState.selectedSheet = "";
  excelState.headerRowIndex = 0;
  excelState.headerColIndex = -1;
  excelState.ctHeaderColIndex = -1;
  excelState.rows = [];
  sheetSelect.innerHTML = "";
  columnSelect.innerHTML = "";
  ctColumnSelect.innerHTML = "";
  excelOptions.hidden = true;
  columnSelectField.hidden = true;
  ctColumnSelectField.hidden = true;
}

function requireFiles() {
  if (!pdfInput.files?.length || !excelInput.files?.length) {
    setStatus("PDF와 엑셀 파일을 모두 선택해 주세요.");
    return false;
  }
  return true;
}

function buildLinesFromTextItems(items) {
  let buffer = "";
  items.forEach((item) => {
    const text = typeof item.str === "string" ? item.str : "";
    buffer += text;
    buffer += item.hasEOL ? "\n" : " ";
  });
  return buffer.replace(/[ \t]+\n/g, "\n").replace(/\n{2,}/g, "\n").trim();
}

function pickDefaultSheet(sheetNames) {
  const matches = sheetNames.filter((name) => /<패킹>|패킹|packing|pacing/i.test(name));
  if (matches.length === 1) return matches[0];
  return "";
}

function getSheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
}

function detectHeaderColumnIndex(rows, headerRowIndex, keywords) {
  const headerRow = rows[headerRowIndex] || [];
  const matches = [];
  headerRow.forEach((cell, index) => {
    const value = normalizeHeader(cell);
    if (!value) return;
    if (keywords.some((keyword) => value.includes(normalizeHeader(keyword)))) {
      matches.push(index);
    }
  });
  if (matches.length === 1) {
    return matches[0];
  }
  return -1;
}

function columnLabel(index) {
  let label = "";
  let n = index + 1;
  while (n > 0) {
    const mod = (n - 1) % 26;
    label = String.fromCharCode(65 + mod) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function populateColumnSelect(selectEl, rows, headerRowIndex) {
  const headerRow = rows[headerRowIndex] || [];
  selectEl.innerHTML = "";
  headerRow.forEach((cell, index) => {
    const value = normalizeWhitespace(String(cell ?? ""));
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = `${columnLabel(index)}: ${value || "빈칸"}`;
    selectEl.appendChild(option);
  });
}

function updateColumnSelection() {
  if (!excelState.rows.length) return;
  const headerRowIndex = excelState.headerRowIndex;
  const detectedFc = detectHeaderColumnIndex(excelState.rows, headerRowIndex, FC_HEADER_KEYWORDS);
  const detectedCt = detectHeaderColumnIndex(excelState.rows, headerRowIndex, CT_HEADER_KEYWORDS);
  populateColumnSelect(columnSelect, excelState.rows, headerRowIndex);
  populateColumnSelect(ctColumnSelect, excelState.rows, headerRowIndex);

  if (detectedFc >= 0) {
    excelState.headerColIndex = detectedFc;
    columnSelect.value = String(detectedFc);
    columnSelectField.hidden = true;
  } else {
    excelState.headerColIndex = -1;
    columnSelectField.hidden = false;
  }

  if (detectedCt >= 0) {
    excelState.ctHeaderColIndex = detectedCt;
    ctColumnSelect.value = String(detectedCt);
    ctColumnSelectField.hidden = true;
  } else {
    excelState.ctHeaderColIndex = -1;
    ctColumnSelectField.hidden = false;
  }
}

function extractFcFromLines(lines) {
  const indices = [];
  const candidates = [];
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    if (!line.includes("받는 사람")) continue;
    const match = line.match(/받는 사람\s*:\s*(.*)$/);
    if (match) {
      indices.push(i);
      const inline = normalizeFc(match[1] || "");
      if (inline) {
        candidates.push(inline);
      } else {
        const nextLine = lines[i + 1] ?? "";
        const nextValue = normalizeFc(nextLine);
        if (nextValue) candidates.push(nextValue);
      }
    }
  }
  if (indices.length !== 1) {
    return { fc: null, count: indices.length };
  }
  const fc = candidates[0] ?? "";
  return { fc: fc || null, count: indices.length };
}

async function parsePdf(file) {
  const buffer = await file.arrayBuffer();
  const doc = await window.pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];

  for (let i = 1; i <= doc.numPages; i += 1) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    const pageText = buildLinesFromTextItems(content.items);
    pages.push({ pageNo: i, text: pageText });
  }

  return pages;
}

function extractRecordsFromPdf(pages) {
  const issues = [];
  const records = [];
  const mrbRegex = /MR[A-Z0-9]*-[0-9]{3}/g;

  pages.forEach(({ pageNo, text }) => {
    const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);
    const fcInfo = extractFcFromLines(lines);
    const mrbMatches = text.match(mrbRegex) ?? [];

    if (fcInfo.count !== 1) {
      issues.push(`페이지 ${pageNo}: FC는 페이지당 1개여야 합니다. (검출: ${fcInfo.count}개)`);
    }

    if (mrbMatches.length !== 1) {
      issues.push(`페이지 ${pageNo}: MRB는 페이지당 1개여야 합니다. (검출: ${mrbMatches.length}개)`);
    }

    if (fcInfo.count !== 1 || mrbMatches.length !== 1) {
      return;
    }

    if (!fcInfo.fc) {
      issues.push(`페이지 ${pageNo}: FC가 비어 있습니다.`);
      return;
    }

    const mrb = mrbMatches[0];
    const boxNoRaw = mrb.split("-").pop();
    const boxNo = boxNoRaw ? Number.parseInt(boxNoRaw, 10) : NaN;

    if (!Number.isFinite(boxNo)) {
      issues.push(`페이지 ${pageNo}: MRB 박스 번호를 해석할 수 없습니다.`);
      return;
    }

    records.push({
      pageNo,
      fc: fcInfo.fc,
      mrb,
      boxNo
    });
  });

  return { records, issues };
}

function loadWorkbookFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("엑셀 파일을 읽을 수 없습니다."));
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        const workbook = window.XLSX.read(data, { type: "array" });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function initializeExcelState(workbook) {
  excelState.workbook = workbook;
  excelState.sheetNames = workbook.SheetNames || [];
  sheetSelect.innerHTML = "";
  const defaultSheet = pickDefaultSheet(excelState.sheetNames);
  if (!defaultSheet) {
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "시트를 선택하세요";
    sheetSelect.appendChild(placeholder);
  }
  excelState.sheetNames.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    sheetSelect.appendChild(option);
  });

  excelState.selectedSheet = defaultSheet || "";
  sheetSelect.value = excelState.selectedSheet;
  excelOptions.hidden = false;
  updateRowsForSelectedSheet();
}

function updateRowsForSelectedSheet() {
  if (!excelState.workbook || !excelState.selectedSheet) {
    excelState.rows = [];
    excelState.headerColIndex = -1;
    excelState.ctHeaderColIndex = -1;
    columnSelectField.hidden = true;
    ctColumnSelectField.hidden = true;
    return;
  }
  excelState.rows = getSheetRows(excelState.workbook, excelState.selectedSheet);
  const headerRowValue = Number.parseInt(headerRowInput.value, 10);
  excelState.headerRowIndex = Number.isFinite(headerRowValue) && headerRowValue > 0 ? headerRowValue - 1 : 0;
  updateColumnSelection();
}

function parseBoxNo(rawValue) {
  if (rawValue === null || rawValue === undefined) return NaN;
  if (typeof rawValue === "number") {
    return Number.isFinite(rawValue) ? Math.trunc(rawValue) : NaN;
  }
  const text = String(rawValue).trim();
  if (!text) return NaN;
  const tailMatch = text.match(/(\d{1,3})\s*$/);
  if (tailMatch) {
    return Number.parseInt(tailMatch[1], 10);
  }
  const anyMatch = text.match(/(\d+)/g);
  if (anyMatch && anyMatch.length) {
    return Number.parseInt(anyMatch[anyMatch.length - 1], 10);
  }
  return NaN;
}

function buildRowOrderFromState() {
  if (!excelState.workbook) {
    return { rows: [], errors: ["엑셀 파일을 다시 선택해 주세요."] };
  }
  if (!excelState.selectedSheet) {
    return { rows: [], errors: ["엑셀 시트를 선택해 주세요."] };
  }
  if (excelState.headerColIndex < 0) {
    return { rows: [], errors: ["물류센터 컬럼을 선택해 주세요."] };
  }
  if (excelState.ctHeaderColIndex < 0) {
    return { rows: [], errors: ["C/T NO. 컬럼을 선택해 주세요."] };
  }
  if (!excelState.rows.length) {
    return { rows: [], errors: ["선택한 시트에서 데이터를 찾지 못했습니다."] };
  }

  const rows = [];
  const seen = new Set();
  for (let r = excelState.headerRowIndex + 1; r < excelState.rows.length; r += 1) {
    const row = excelState.rows[r] || [];
    const rawFc = String(row[excelState.headerColIndex] ?? "").trim();
    const rawCt = row[excelState.ctHeaderColIndex];
    if (!rawFc && (rawCt === null || rawCt === undefined || String(rawCt).trim() === "")) {
      continue;
    }
    const normalized = normalizeFc(rawFc);
    if (!normalized) continue;
    const boxNo = parseBoxNo(rawCt);
    if (!Number.isFinite(boxNo)) {
      return {
        rows: [],
        errors: [`${r + 1}행: C/T NO. 값을 박스번호로 해석할 수 없습니다.`]
      };
    }
    const key = `${normalized}:${boxNo}`;
    if (seen.has(key)) {
      return {
        rows: [],
        errors: [`${r + 1}행: 동일한 물류센터/박스번호 조합이 중복되었습니다.`]
      };
    }
    seen.add(key);
    rows.push({
      rowIndex: r,
      fcRaw: normalizeWhitespace(rawFc),
      fcNormalized: normalized,
      boxNo
    });
  }

  return { rows, errors: [] };
}

function buildOutput(rows, records, showFc) {
  const errors = [];
  const byKey = new Map();

  records.forEach((record) => {
    const key = `${record.fc}:${record.boxNo}`;
    if (byKey.has(key)) {
      const prev = byKey.get(key);
      errors.push(
        `FC ${record.fc}: 박스번호 ${record.boxNo}가 중복되었습니다. (페이지 ${prev.pageNo}, ${record.pageNo})`
      );
    } else {
      byKey.set(key, record);
    }
  });

  if (errors.length) {
    return { output: "", errors };
  }

  const lines = [];
  let lastFc = null;
  rows.forEach((row) => {
    const key = `${row.fcNormalized}:${row.boxNo}`;
    const record = byKey.get(key);
    if (!record) {
      errors.push(`엑셀 ${row.rowIndex + 1}행: ${row.fcRaw} / 박스번호 ${row.boxNo}가 PDF에 없습니다.`);
      return;
    }
    if (showFc && lastFc !== row.fcRaw) {
      lines.push(`FC: ${row.fcRaw}`);
      lastFc = row.fcRaw;
    }
    lines.push(record.mrb);
  });

  if (errors.length) {
    return { output: "", errors };
  }

  return { output: lines.join("\n"), errors: [] };
}

async function handleParse() {
  showErrors([]);
  resetOutput();

  if (!requireFiles()) return;
  if (!window.pdfjsLib || !window.XLSX) {
    showErrors(["필요한 라이브러리를 불러오지 못했습니다. 새로고침 후 다시 시도해 주세요."]);
    setStatus("검증 실패: 라이브러리 로드 오류");
    return;
  }

  try {
    setStatus("파일을 분석 중입니다...");
    const [{ rows, errors }, pages] = await Promise.all([
      Promise.resolve(buildRowOrderFromState()),
      parsePdf(pdfInput.files[0])
    ]);
    if (errors.length) {
      showErrors(errors);
      setStatus("검증 실패: 엑셀 입력 오류");
      return;
    }
    if (!rows.length) {
      showErrors(["엑셀에서 물류센터/C/T NO. 목록을 찾지 못했습니다."]);
      setStatus("검증 실패: 엑셀 입력 오류");
      return;
    }

    const { records, issues } = extractRecordsFromPdf(pages);
    if (issues.length) {
      showErrors(issues);
      setStatus("검증 실패: PDF에서 문제가 발견되었습니다.");
      return;
    }

    const { output, errors } = buildOutput(rows, records, showFcToggle.checked);
    if (errors.length) {
      showErrors(errors);
      setStatus("검증 실패: FC/박스번호 오류가 있습니다.");
      return;
    }

    resultText.value = output;
    resultSection.hidden = false;
    copyBtn.disabled = !output;
    setStatus("완료되었습니다. 복사 버튼으로 결과를 복사하세요.");
  } catch (err) {
    showErrors([err instanceof Error ? err.message : "처리 중 오류가 발생했습니다."]);
    setStatus("검증 실패: 입력을 확인해 주세요.");
  }
}

async function handleCopy() {
  if (!resultText.value) return;
  await navigator.clipboard.writeText(resultText.value);
  setStatus("복사가 완료되었습니다.");
}

parseBtn.addEventListener("click", handleParse);
copyBtn.addEventListener("click", handleCopy);

excelInput.addEventListener("change", async () => {
  resetOutput();
  showErrors([]);
  if (!excelInput.files?.length) {
    clearExcelState();
    return;
  }
  if (!window.XLSX) {
    showErrors(["엑셀 라이브러리를 불러오지 못했습니다. 새로고침 후 다시 시도해 주세요."]);
    return;
  }
  try {
    setStatus("엑셀 파일을 읽는 중입니다...");
    const workbook = await loadWorkbookFromFile(excelInput.files[0]);
    initializeExcelState(workbook);
  if (!excelState.selectedSheet) {
    setStatus("엑셀 시트를 선택해 주세요.");
    return;
  }
  if (excelState.headerColIndex < 0 || excelState.ctHeaderColIndex < 0) {
    setStatus("물류센터/C/T NO. 컬럼을 선택해 주세요.");
    return;
  }
  setStatus("엑셀 시트와 헤더 정보를 확인해 주세요.");
  } catch (err) {
    clearExcelState();
    showErrors([err instanceof Error ? err.message : "엑셀 파일을 읽는 중 오류가 발생했습니다."]);
  }
});

sheetSelect.addEventListener("change", () => {
  excelState.selectedSheet = sheetSelect.value;
  updateRowsForSelectedSheet();
  if (!excelState.selectedSheet) {
    setStatus("엑셀 시트를 선택해 주세요.");
    return;
  }
  if (excelState.headerColIndex < 0 || excelState.ctHeaderColIndex < 0) {
    setStatus("물류센터/C/T NO. 컬럼을 선택해 주세요.");
  }
});

headerRowInput.addEventListener("change", () => {
  updateRowsForSelectedSheet();
  if (excelState.headerColIndex < 0 || excelState.ctHeaderColIndex < 0) {
    setStatus("물류센터/C/T NO. 컬럼을 선택해 주세요.");
  }
});

columnSelect.addEventListener("change", () => {
  excelState.headerColIndex = Number.parseInt(columnSelect.value, 10);
  if (excelState.headerColIndex >= 0) {
    setStatus("엑셀 시트와 헤더 정보를 확인해 주세요.");
  }
});

ctColumnSelect.addEventListener("change", () => {
  excelState.ctHeaderColIndex = Number.parseInt(ctColumnSelect.value, 10);
  if (excelState.ctHeaderColIndex >= 0) {
    setStatus("엑셀 시트와 헤더 정보를 확인해 주세요.");
  }
});
