/* 모든 로직은 브라우저에서 동작합니다. */

const pdfInput = document.getElementById("pdfInput");
const excelInput = document.getElementById("excelInput");
const parseBtn = document.getElementById("parseBtn");
const copyBtn = document.getElementById("copyBtn");
const statusEl = document.getElementById("status");
const errorBox = document.getElementById("errorBox");
const resultSection = document.getElementById("resultSection");
const resultText = document.getElementById("resultText");

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

function normalizeComparable(value) {
  return normalizeWhitespace(String(value ?? ""))
    .toLowerCase()
    .replace(/fc/g, "")
    .replace(/[^a-z0-9가-힣]/gi, "");
}

function extractFcFromLines(lines, fcCandidates) {
  const indices = [];
  const candidates = [];
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const normalizedLine = normalizeComparable(line);
    if (!normalizedLine.includes("받는사람")) continue;
    indices.push(i);
    const inlineMatch = line.match(/받는\s*사람\s*[:：]?\s*(.*)$/);
    const inlineValue = inlineMatch ? normalizeFc(inlineMatch[1] || "") : "";
    if (inlineValue) {
      candidates.push(inlineValue);
      continue;
    }
    let nextValue = "";
    for (let j = i + 1; j < lines.length; j += 1) {
      const nextLine = lines[j];
      if (!nextLine) continue;
      nextValue = normalizeFc(nextLine);
      if (nextValue) break;
    }
    if (nextValue) candidates.push(nextValue);
  }
  if (indices.length !== 1) {
    // fallback: try to match FC names from 엑셀 목록
    if (fcCandidates && fcCandidates.length) {
      const found = new Set();
      lines.forEach((line) => {
        const normalizedLine = normalizeComparable(line);
        fcCandidates.forEach((candidate) => {
          if (normalizedLine.includes(candidate.key)) {
            found.add(candidate.normalized);
          }
        });
      });
      if (found.size === 1) {
        return { fc: Array.from(found)[0], count: 1 };
      }
    }
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

function normalizeMrb(value) {
  return String(value ?? "").replace(/\s+/g, "");
}

function extractRecordsFromPdf(pages, fcCandidates) {
  const issues = [];
  const records = [];
  const mrbRegex = /MR[A-Z0-9\s]{0,20}-\s*[0-9]{3}/g;

  pages.forEach(({ pageNo, text }) => {
    const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);
    const fcInfo = extractFcFromLines(lines, fcCandidates);
    const mrbMatches = (text.match(mrbRegex) ?? []).map(normalizeMrb);

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

  let headerRowIndex = -1;
  let headerColIndex = -1;
  let ctHeaderColIndex = -1;
  const scanLimit = Math.min(sheetRows.length, 10);
  for (let i = 0; i < scanLimit; i += 1) {
    const fcIdx = detectHeaderColumnIndex(sheetRows, i, FC_HEADER_KEYWORDS);
    const ctIdx = detectHeaderColumnIndex(sheetRows, i, CT_HEADER_KEYWORDS);
    if (fcIdx >= 0 && ctIdx >= 0) {
      headerRowIndex = i;
      headerColIndex = fcIdx;
      ctHeaderColIndex = ctIdx;
      break;
    }
  }

  if (headerRowIndex < 0 || headerColIndex < 0) {
    return { rows: [], errors: ["<패킹> 시트에서 물류센터 컬럼을 찾지 못했습니다."] };
  }
  if (ctHeaderColIndex < 0) {
    return { rows: [], errors: ["<패킹> 시트에서 C/T NO. 컬럼을 찾지 못했습니다."] };
  }

  const resultRows = [];
  for (let r = headerRowIndex + 1; r < sheetRows.length; r += 1) {
    const row = sheetRows[r] || [];
    const rawFc = String(row[headerColIndex] ?? "").trim();
    const rawCt = row[ctHeaderColIndex];
    if (!rawFc && (rawCt === null || rawCt === undefined || String(rawCt).trim() === "")) {
      resultRows.push({
        rowIndex: r,
        fcRaw: "",
        fcNormalized: null,
        boxNo: null
      });
      continue;
    }
    const normalized = normalizeFc(rawFc);
    const boxNo = parseBoxNo(rawCt);
    if (!normalized || !Number.isFinite(boxNo)) {
      resultRows.push({
        rowIndex: r,
        fcRaw: normalizeWhitespace(rawFc),
        fcNormalized: normalized || null,
        boxNo: Number.isFinite(boxNo) ? boxNo : null
      });
      continue;
    }
    resultRows.push({
      rowIndex: r,
      fcRaw: normalizeWhitespace(rawFc),
      fcNormalized: normalized,
      boxNo
    });
  }

  return { rows: resultRows, errors: [] };
}

function buildOutput(rows, records) {
  const byKey = new Map();

  records.forEach((record) => {
    const key = `${record.fc}:${record.boxNo}`;
    if (!byKey.has(key)) {
      byKey.set(key, record);
    }
  });

  const lines = [];
  let matchCount = 0;
  rows.forEach((row) => {
    if (row.fcNormalized && Number.isFinite(row.boxNo)) {
      const key = `${row.fcNormalized}:${row.boxNo}`;
      const record = byKey.get(key);
      if (record) {
        matchCount += 1;
        lines.push(record.mrb);
      } else {
        lines.push("");
      }
    } else {
      lines.push("");
    }
  });

  const sampleExcelKeys = rows
    .filter((row) => row.fcNormalized && Number.isFinite(row.boxNo))
    .slice(0, 3)
    .map((row) => `${row.fcNormalized}:${row.boxNo}`);
  const samplePdfKeys = Array.from(byKey.keys()).slice(0, 3);

  return {
    output: lines.join("\n"),
    matchCount,
    sampleExcelKeys,
    samplePdfKeys
  };
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

    const fcCandidates = Array.from(
      rows.reduce((map, row) => {
        if (!map.has(row.fcNormalized)) {
          map.set(row.fcNormalized, {
            normalized: row.fcNormalized,
            key: normalizeComparable(row.fcNormalized)
          });
        }
        return map;
      }, new Map()).values()
    );

    const { records, issues } = extractRecordsFromPdf(pages, fcCandidates);
    if (issues.length) {
      showErrors(issues);
      setStatus("검증 실패: PDF에서 문제가 발견되었습니다.");
      return;
    }

    const { output, matchCount, sampleExcelKeys, samplePdfKeys } = buildOutput(rows, records);
    if (records.length === 0) {
      showErrors(["PDF에서 운송장 번호를 찾지 못했습니다. PDF 형식을 확인해 주세요."]);
      setStatus("검증 실패: PDF 인식 오류");
      return;
    }
    if (matchCount === 0) {
      showErrors([
        "엑셀과 PDF가 매칭되지 않았습니다.",
        `엑셀 예시 키: ${sampleExcelKeys.join(", ") || "없음"}`,
        `PDF 예시 키: ${samplePdfKeys.join(", ") || "없음"}`
      ]);
      setStatus("검증 실패: 매칭 오류");
      return;
    }

    resultText.value = output;
    resultSection.hidden = false;
    copyBtn.disabled = !output;
    setStatus(`완료되었습니다. 매칭 ${matchCount}건`);
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
    setStatus("<패킹> 시트를 자동으로 찾습니다. 분석 버튼을 눌러주세요.");
  } catch (err) {
    clearExcelState();
    showErrors([err instanceof Error ? err.message : "엑셀 파일을 읽는 중 오류가 발생했습니다."]);
  }
});
