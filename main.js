/* 모든 로직은 브라우저에서 동작합니다. */

const pdfInput = document.getElementById("pdfInput");
const excelInput = document.getElementById("excelInput");
const parseBtn = document.getElementById("parseBtn");
const copyBtn = document.getElementById("copyBtn");
const statusEl = document.getElementById("status");
const errorBox = document.getElementById("errorBox");
let warningBox = document.getElementById("warningBox");
const resultSection = document.getElementById("resultSection");
const resultText = document.getElementById("resultText");

const PDF_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const PACKING_FC_COL_INDEX = 2; // C열
const PACKING_CT_COL_INDEX = 4; // E열

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;
}

function normalizeWhitespace(value) {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeFc(value) {
  let text = String(value ?? "");
  if (!text) return "";
  text = text.replace(/\s+/g, "");
  text = text.replace(/대표번호.*$/g, "");
  const match = text.match(/^(.*?FC)/);
  if (match) text = match[1];
  text = text.replace(/\(\d+\)FC$/g, "");
  text = text.replace(/FC$/g, "");
  return text.trim();
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

function ensureWarningBox() {
  if (warningBox) return;
  warningBox = document.createElement("div");
  warningBox.id = "warningBox";
  warningBox.style.marginTop = "8px";
  warningBox.style.padding = "10px 12px";
  warningBox.style.borderRadius = "6px";
  warningBox.style.background = "#fff4cc";
  warningBox.style.border = "1px solid #f1c232";
  warningBox.style.color = "#7f6000";
  warningBox.style.fontSize = "14px";
  warningBox.hidden = true;
  if (errorBox && errorBox.parentNode) {
    errorBox.parentNode.insertBefore(warningBox, errorBox.nextSibling);
  } else {
    document.body.appendChild(warningBox);
  }
}

function showWarnings(messages) {
  ensureWarningBox();
  if (!messages.length) {
    warningBox.hidden = true;
    warningBox.innerHTML = "";
    return;
  }
  warningBox.hidden = false;
  warningBox.innerHTML = messages.map((msg) => `• ${msg}`).join("<br>");
}

function resetOutput() {
  resultSection.hidden = true;
  resultText.value = "";
  copyBtn.disabled = true;
}

const excelState = {
  workbook: null,
  sheetNames: []
};

function clearExcelState() {
  excelState.workbook = null;
  excelState.sheetNames = [];
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

function buildRecipientHint(compactText) {
  const keyword = "받는사람";
  const index = compactText.indexOf(keyword);
  if (index < 0) {
    return "받는사람 포함: 아니오";
  }
  const start = Math.max(0, index - 20);
  const end = Math.min(compactText.length, index + keyword.length + 20);
  const snippet = compactText.slice(start, end);
  return `받는사람 포함: 예 (스니펫: ${snippet})`;
}

async function parsePdf(file) {
  const buffer = await file.arrayBuffer();
  const doc = await window.pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];

  for (let i = 1; i <= doc.numPages; i += 1) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    const pageText = buildLinesFromTextItems(content.items);
    pages.push({ pageNo: i, text: pageText, items: content.items });
  }

  return pages;
}

function extractPdfMap(pages) {
  const issues = [];
  const map = new Map();
  const mrbRegex = /MRB\d+-\d{3}/g;
  const fcRegex = /[-–]?받는사람[:：](.*?FC)/g;

  pages.forEach(({ pageNo, text }) => {
    const raw = text;
    const compact = raw.replace(/\s+/g, "");
    const fcMatches = Array.from(compact.matchAll(fcRegex));
    const mrbMatches = (compact.match(mrbRegex) ?? []);

    if (fcMatches.length !== 1) {
      const reason = fcMatches.length === 0 ? "FC 검출 0개" : "FC 검출 다중";
      issues.push(`페이지 ${pageNo}: ${reason}. ${buildRecipientHint(compact)}`);
    }

    if (mrbMatches.length !== 1) {
      issues.push(`페이지 ${pageNo}: MRB는 페이지당 1개여야 합니다. (검출: ${mrbMatches.length}개)`);
    }

    if (fcMatches.length !== 1 || mrbMatches.length !== 1) {
      return;
    }

    const fcRaw = fcMatches[0]?.[1] ?? "";
    const fcNormalized = normalizeFc(fcRaw);
    if (!fcNormalized) {
      issues.push(`페이지 ${pageNo}: FC가 비어 있습니다.`);
      return;
    }

    const mrb = mrbMatches[0];
    const ctRaw = mrb.split("-").pop();
    const ct = ctRaw ? Number.parseInt(ctRaw, 10) : NaN;

    if (!Number.isFinite(ct)) {
      issues.push(`페이지 ${pageNo}: MRB CT 번호를 해석할 수 없습니다.`);
      return;
    }

    const key = `${fcNormalized}|${ct}`;
    if (!key || key.startsWith("|")) {
      issues.push(`페이지 ${pageNo}: FC를 해석할 수 없습니다.`);
      return;
    }

    if (map.has(key)) {
      issues.push(`페이지 ${pageNo}: PDF에서 중복된 FC/CT (${key})가 발견되었습니다.`);
      return;
    }

    map.set(key, mrb);
  });

  return { map, issues };
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
}

function parseCtValue(rawValue) {
  if (rawValue === null || rawValue === undefined) return NaN;
  if (typeof rawValue === "number") {
    return Number.isFinite(rawValue) ? Math.trunc(rawValue) : NaN;
  }
  const text = String(rawValue).trim();
  if (!text) return NaN;
  const match = text.match(/(\d+)/g);
  if (!match || !match.length) return NaN;
  return Number.parseInt(match[match.length - 1], 10);
}

function buildRowOrderFromState() {
  if (!excelState.workbook) {
    return { rows: [], errors: ["엑셀 파일을 다시 선택해 주세요."] };
  }
  if (!excelState.sheetNames.length) {
    return { rows: [], errors: ["엑셀에서 시트를 찾지 못했습니다."] };
  }

  const selectedSheet = pickDefaultSheet(excelState.sheetNames);
  if (!selectedSheet) {
    return { rows: [], errors: ["<패킹> 시트를 찾지 못했습니다. 시트명을 확인해 주세요."] };
  }

  const sheetRows = getSheetRows(excelState.workbook, selectedSheet);
  if (!sheetRows.length) {
    return { rows: [], errors: ["<패킹> 시트에서 데이터를 찾지 못했습니다."] };
  }

  const headerRowIndex = 0;
  if (sheetRows.length <= headerRowIndex + 1) {
    return { rows: [], errors: ["<패킹> 시트에서 데이터 행을 찾지 못했습니다."] };
  }

  const resultRows = [];
  for (let r = headerRowIndex + 1; r < sheetRows.length; r += 1) {
    const row = sheetRows[r] || [];
    const rawFc = String(row[PACKING_FC_COL_INDEX] ?? "").trim();
    const rawCt = row[PACKING_CT_COL_INDEX];
    if (!rawFc && (rawCt === null || rawCt === undefined || String(rawCt).trim() === "")) {
      resultRows.push({
        rowIndex: r,
        fcRaw: "",
        fcNormalized: null,
        ct: null
      });
      continue;
    }
    const normalized = normalizeFc(rawFc);
    const ct = parseCtValue(rawCt);
    if (!normalized || !Number.isFinite(ct)) {
      resultRows.push({
        rowIndex: r,
        fcRaw: normalizeWhitespace(rawFc),
        fcNormalized: normalized || null,
        ct: Number.isFinite(ct) ? ct : null
      });
      continue;
    }
    resultRows.push({
      rowIndex: r,
      fcRaw: normalizeWhitespace(rawFc),
      fcNormalized: normalized,
      ct
    });
  }

  return { rows: resultRows, errors: [] };
}

function buildOutput(rows, pdfMap) {
  const lines = [];
  const issues = [];
  const skips = [];

  rows.forEach((row) => {
    const displayRow = row.rowIndex + 1;
    if (!row.fcNormalized) {
      issues.push(`행 ${displayRow}: 센터명이 비어 있습니다.`);
      return;
    }
    if (!Number.isFinite(row.ct)) {
      issues.push(`행 ${displayRow}: C/T NO 값이 올바르지 않습니다. (센터: ${row.fcRaw || "-"})`);
      return;
    }
    const key = `${row.fcNormalized}|${row.ct}`;
    const mrb = pdfMap.get(key);
    if (!mrb) {
      skips.push({
        rowNo: displayRow,
        fcRaw: row.fcRaw || "-",
        ct: row.ct
      });
      return;
    }
    lines.push(mrb);
  });

  return { outputLines: lines, issues, skips };
}

async function handleParse() {
  showErrors([]);
  showWarnings([]);
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

    const { map: pdfMap, issues: pdfIssues } = extractPdfMap(pages);
    if (pdfIssues.length) {
      showErrors(pdfIssues);
      setStatus("검증 실패: PDF에서 문제가 발견되었습니다.");
      return;
    }
    if (pdfMap.size === 0) {
      showErrors(["PDF에서 운송장 번호를 찾지 못했습니다. PDF 형식을 확인해 주세요."]);
      setStatus("검증 실패: PDF 인식 오류");
      return;
    }

    const { outputLines, issues: outputIssues, skips } = buildOutput(rows, pdfMap);
    if (outputIssues.length) {
      showErrors(outputIssues);
      setStatus("검증 실패: 엑셀/PDF 매칭 오류");
      return;
    }

    if (skips.length) {
      const preview = skips.slice(0, 20).map((item) => {
        return `행 ${item.rowNo}: ${item.fcRaw}, CT ${item.ct}`;
      });
      const extra = skips.length > 20 ? `외 ${skips.length - 20}건` : "";
      const warningMessages = [
        `PDF에 없어 건너뜀: ${skips.length}건`,
        ...preview,
        ...(extra ? [extra] : [])
      ];
      showWarnings(warningMessages);
    }

    if (outputLines.length === 0) {
      resultText.value = "";
      resultSection.hidden = true;
      copyBtn.disabled = true;
      setStatus("출력할 MRB가 없습니다(PDF에 해당 센터/CT 없음).");
      return;
    }

    const output = outputLines.join("\n");
    resultText.value = output;
    resultSection.hidden = false;
    copyBtn.disabled = false;
    setStatus(`완료되었습니다. 매칭 ${outputLines.length}건`);
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
  showWarnings([]);
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
    setStatus("<패킹> 시트를 자동으로 찾습니다. 분석 버튼을 눌러주세요.");
  } catch (err) {
    clearExcelState();
    showErrors([err instanceof Error ? err.message : "엑셀 파일을 읽는 중 오류가 발생했습니다."]);
  }
});
