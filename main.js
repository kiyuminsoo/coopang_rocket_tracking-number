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
  if (ui.copyVerifyBtn) ui.copyVerifyBtn.disabled = true;
  if (ui.statsSection) ui.statsSection.hidden = true;
  if (ui.matchedBody) ui.matchedBody.innerHTML = "";
  if (ui.skipSummaryBody) ui.skipSummaryBody.innerHTML = "";
  if (ui.skipDetailsBody) ui.skipDetailsBody.innerHTML = "";
  if (ui.unusedList) ui.unusedList.innerHTML = "";
  if (ui.unusedBox) ui.unusedBox.hidden = true;
  latestResult = null;
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
  const matched = [];
  const usedMrbs = new Set();

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
    const ctFromMrb = Number.parseInt(String(mrb).split("-").pop() || "", 10);
    const ctMatches = Number.isFinite(ctFromMrb) && ctFromMrb === row.ct;
    if (!ctMatches) {
      issues.push(`행 ${displayRow}: CT와 MRB 끝 3자리가 일치하지 않습니다. (CT: ${row.ct}, MRB: ${mrb})`);
      return;
    }
    lines.push(mrb);
    matched.push({
      rowNo: displayRow,
      fcRaw: row.fcRaw || "-",
      ct: row.ct,
      mrb,
      ctMatches: true
    });
    usedMrbs.add(mrb);
  });

  return { outputLines: lines, issues, skips, matched, usedMrbs };
}

const ui = {
  statsSection: null,
  statsValues: {},
  matchedSection: null,
  matchedBody: null,
  skipSection: null,
  skipSummaryBody: null,
  skipDetailsBody: null,
  unusedBox: null,
  unusedList: null,
  copyVerifyBtn: null,
  resultLabel: null,
  arrivalInput: null,
  arrivalHint: null
};

function ensureStatsSection() {
  if (ui.statsSection) return;
  const inputCard = parseBtn.closest(".card");
  const statsSection = document.createElement("section");
  statsSection.className = "card";
  statsSection.id = "statsSection";
  const title = document.createElement("h2");
  title.textContent = "요약";
  const grid = document.createElement("div");
  grid.style.display = "grid";
  grid.style.gridTemplateColumns = "repeat(auto-fit, minmax(180px, 1fr))";
  grid.style.gap = "10px";
  grid.style.marginTop = "10px";

  const items = [
    { key: "totalExcelRows", label: "엑셀 대상 행 수" },
    { key: "totalPdfRecords", label: "PDF 페이지/레코드 수" },
    { key: "matchedCount", label: "매칭 성공" },
    { key: "skippedCount", label: "PLT 출고(밀크런 송장 미발행)로 제외" },
    { key: "fatalErrorCount", label: "치명적 오류" },
    { key: "unusedPdfCount", label: "PDF 미사용 MRB" }
  ];

  items.forEach((item) => {
    const box = document.createElement("div");
    box.style.border = "1px solid #e1e1e1";
    box.style.borderRadius = "6px";
    box.style.padding = "10px 12px";
    box.style.background = "#fff";
    const label = document.createElement("div");
    label.textContent = item.label;
    label.style.fontSize = "12px";
    label.style.color = "#555";
    const value = document.createElement("div");
    value.textContent = "-";
    value.style.fontSize = "18px";
    value.style.fontWeight = "600";
    value.style.marginTop = "4px";
    box.appendChild(label);
    box.appendChild(value);
    grid.appendChild(box);
    ui.statsValues[item.key] = value;
  });

  statsSection.appendChild(title);
  statsSection.appendChild(grid);
  statsSection.hidden = true;

  if (inputCard && inputCard.parentNode) {
    inputCard.parentNode.insertBefore(statsSection, resultSection);
  } else {
    document.body.appendChild(statsSection);
  }
  ui.statsSection = statsSection;
}

function setStats(stats) {
  ensureStatsSection();
  ui.statsSection.hidden = false;
  const map = {
    totalExcelRows: stats.totalExcelRows ?? "-",
    totalPdfRecords: stats.totalPdfPages !== undefined
      ? `${stats.totalPdfPages} / ${stats.totalPdfRecords ?? 0}`
      : "-",
    matchedCount: stats.matchedCount ?? "-",
    skippedCount: stats.skippedCount ?? "-",
    fatalErrorCount: stats.fatalErrorCount ?? "-",
    unusedPdfCount: stats.unusedPdfCount ?? "-"
  };
  Object.entries(map).forEach(([key, value]) => {
    if (ui.statsValues[key]) ui.statsValues[key].textContent = String(value);
  });
}

function ensureResultLayout() {
  if (ui.matchedSection) return;
  const resultTitle = resultSection.querySelector("h2");
  if (resultTitle) resultTitle.textContent = "결과";

  const matchedSection = document.createElement("div");
  const matchedTitle = document.createElement("h3");
  matchedTitle.textContent = "매칭 성공";
  matchedTitle.style.marginTop = "18px";
  const matchedWrap = document.createElement("div");
  matchedWrap.style.maxHeight = "280px";
  matchedWrap.style.overflow = "auto";
  matchedWrap.style.border = "1px solid #e1e1e1";
  matchedWrap.style.borderRadius = "6px";
  matchedWrap.style.marginTop = "8px";

  const matchedTable = document.createElement("table");
  matchedTable.style.width = "100%";
  matchedTable.style.borderCollapse = "collapse";
  const matchedHead = document.createElement("thead");
  const headRow = document.createElement("tr");
  ["엑셀 행번호", "센터(엑셀 원문)", "CT", "MRB", "CT==MRB끝3자리"].forEach((text) => {
    const th = document.createElement("th");
    th.textContent = text;
    th.style.position = "sticky";
    th.style.top = "0";
    th.style.background = "#fff";
    th.style.borderBottom = "1px solid #ddd";
    th.style.padding = "8px 6px";
    th.style.textAlign = "left";
    th.style.fontSize = "12px";
    headRow.appendChild(th);
  });
  matchedHead.appendChild(headRow);
  const matchedBody = document.createElement("tbody");
  matchedTable.appendChild(matchedHead);
  matchedTable.appendChild(matchedBody);
  matchedWrap.appendChild(matchedTable);
  matchedSection.appendChild(matchedTitle);
  matchedSection.appendChild(matchedWrap);

  const skipSection = document.createElement("div");
  const skipTitle = document.createElement("h3");
  skipTitle.textContent = "스킵";
  skipTitle.style.marginTop = "18px";
  const skipSummaryTitle = document.createElement("div");
  skipSummaryTitle.textContent = "센터별 요약";
  skipSummaryTitle.style.fontSize = "13px";
  skipSummaryTitle.style.marginTop = "8px";
  const skipSummaryWrap = document.createElement("div");
  skipSummaryWrap.style.border = "1px solid #e1e1e1";
  skipSummaryWrap.style.borderRadius = "6px";
  skipSummaryWrap.style.marginTop = "6px";
  skipSummaryWrap.style.overflow = "auto";
  const skipSummaryTable = document.createElement("table");
  skipSummaryTable.style.width = "100%";
  skipSummaryTable.style.borderCollapse = "collapse";
  const skipSummaryHead = document.createElement("thead");
  const skipSummaryHeadRow = document.createElement("tr");
  ["센터", "스킵 건수", "CT 분포"].forEach((text) => {
    const th = document.createElement("th");
    th.textContent = text;
    th.style.position = "sticky";
    th.style.top = "0";
    th.style.background = "#fff";
    th.style.borderBottom = "1px solid #ddd";
    th.style.padding = "8px 6px";
    th.style.textAlign = "left";
    th.style.fontSize = "12px";
    skipSummaryHeadRow.appendChild(th);
  });
  skipSummaryHead.appendChild(skipSummaryHeadRow);
  const skipSummaryBody = document.createElement("tbody");
  skipSummaryTable.appendChild(skipSummaryHead);
  skipSummaryTable.appendChild(skipSummaryBody);
  skipSummaryWrap.appendChild(skipSummaryTable);

  const skipDetails = document.createElement("details");
  skipDetails.style.marginTop = "10px";
  const skipSummary = document.createElement("summary");
  skipSummary.textContent = "자세히 보기";
  skipSummary.style.cursor = "pointer";
  skipSummary.style.fontSize = "13px";
  skipDetails.appendChild(skipSummary);
  const skipDetailsWrap = document.createElement("div");
  skipDetailsWrap.style.maxHeight = "240px";
  skipDetailsWrap.style.overflow = "auto";
  skipDetailsWrap.style.border = "1px solid #e1e1e1";
  skipDetailsWrap.style.borderRadius = "6px";
  skipDetailsWrap.style.marginTop = "8px";
  const skipDetailsTable = document.createElement("table");
  skipDetailsTable.style.width = "100%";
  skipDetailsTable.style.borderCollapse = "collapse";
  const skipDetailsHead = document.createElement("thead");
  const skipDetailsHeadRow = document.createElement("tr");
  ["엑셀 행번호", "센터", "CT", "사유"].forEach((text) => {
    const th = document.createElement("th");
    th.textContent = text;
    th.style.position = "sticky";
    th.style.top = "0";
    th.style.background = "#fff";
    th.style.borderBottom = "1px solid #ddd";
    th.style.padding = "8px 6px";
    th.style.textAlign = "left";
    th.style.fontSize = "12px";
    skipDetailsHeadRow.appendChild(th);
  });
  skipDetailsHead.appendChild(skipDetailsHeadRow);
  const skipDetailsBody = document.createElement("tbody");
  skipDetailsTable.appendChild(skipDetailsHead);
  skipDetailsTable.appendChild(skipDetailsBody);
  skipDetailsWrap.appendChild(skipDetailsTable);
  skipDetails.appendChild(skipDetailsWrap);

  skipSection.appendChild(skipTitle);
  skipSection.appendChild(skipSummaryTitle);
  skipSection.appendChild(skipSummaryWrap);
  skipSection.appendChild(skipDetails);

  const resultLabel = document.createElement("div");
  resultLabel.textContent = "엑셀 붙여넣기용(MRB만)";
  resultLabel.style.fontSize = "13px";
  resultLabel.style.marginTop = "18px";
  resultLabel.style.marginBottom = "6px";

  resultSection.insertBefore(matchedSection, resultText);
  resultSection.insertBefore(skipSection, resultText);
  resultSection.insertBefore(resultLabel, resultText);

  ui.matchedSection = matchedSection;
  ui.matchedBody = matchedBody;
  ui.skipSection = skipSection;
  ui.skipSummaryBody = skipSummaryBody;
  ui.skipDetailsBody = skipDetailsBody;
  ui.resultLabel = resultLabel;
}

function ensureUnusedBox() {
  if (ui.unusedBox) return;
  ensureWarningBox();
  ui.unusedBox = warningBox;
  const list = document.createElement("div");
  list.style.marginTop = "6px";
  list.style.fontSize = "13px";
  ui.unusedList = list;
  ui.unusedBox.appendChild(list);
}

function updateUnusedBox(unusedMrbs) {
  ensureUnusedBox();
  if (!unusedMrbs.length) {
    ui.unusedBox.hidden = true;
    ui.unusedList.innerHTML = "";
    return;
  }
  ui.unusedBox.hidden = false;
  ui.unusedBox.innerHTML = "PDF MRB 중 엑셀에 사용되지 않은 MRB가 있습니다.";
  ui.unusedBox.appendChild(ui.unusedList);
  const preview = unusedMrbs.slice(0, 20);
  const extra = unusedMrbs.length > 20 ? `외 ${unusedMrbs.length - 20}건` : "";
  const lines = [preview.join(", "), extra].filter(Boolean).join("<br>");
  ui.unusedList.innerHTML = lines;
}

function updateMatchedTable(matched) {
  ensureResultLayout();
  ui.matchedBody.innerHTML = "";
  matched.forEach((item) => {
    const row = document.createElement("tr");
    [item.rowNo, item.fcRaw, item.ct, item.mrb, item.ctMatches ? "OK" : "FAIL"].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = String(value);
      td.style.padding = "6px";
      td.style.borderBottom = "1px solid #f0f0f0";
      td.style.fontSize = "13px";
      row.appendChild(td);
    });
    ui.matchedBody.appendChild(row);
  });
}

function buildSkipSummary(skips) {
  const map = new Map();
  skips.forEach((item) => {
    const key = item.fcRaw || "-";
    if (!map.has(key)) {
      map.set(key, { count: 0, ctCounts: new Map(), min: null, max: null });
    }
    const entry = map.get(key);
    entry.count += 1;
    const ctValue = Number.isFinite(item.ct) ? item.ct : null;
    if (ctValue !== null) {
      entry.min = entry.min === null ? ctValue : Math.min(entry.min, ctValue);
      entry.max = entry.max === null ? ctValue : Math.max(entry.max, ctValue);
      entry.ctCounts.set(ctValue, (entry.ctCounts.get(ctValue) || 0) + 1);
    }
  });
  return Array.from(map.entries()).map(([center, entry]) => {
    const topCts = Array.from(entry.ctCounts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([ct]) => ct);
    const range = entry.min !== null ? `${entry.min}~${entry.max}` : "-";
    const topLabel = topCts.length ? ` (상위: ${topCts.join(", ")})` : "";
    return {
      center,
      count: entry.count,
      dist: `${range}${topLabel}`
    };
  });
}

function updateSkipSection(skips) {
  ensureResultLayout();
  ui.skipSummaryBody.innerHTML = "";
  ui.skipDetailsBody.innerHTML = "";
  const summary = buildSkipSummary(skips);
  summary.forEach((item) => {
    const row = document.createElement("tr");
    [item.center, item.count, item.dist].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = String(value);
      td.style.padding = "6px";
      td.style.borderBottom = "1px solid #f0f0f0";
      td.style.fontSize = "13px";
      row.appendChild(td);
    });
    ui.skipSummaryBody.appendChild(row);
  });
  skips.forEach((item) => {
    const row = document.createElement("tr");
    [item.rowNo, item.fcRaw, item.ct, "PDF에 없음"].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = String(value);
      td.style.padding = "6px";
      td.style.borderBottom = "1px solid #f0f0f0";
      td.style.fontSize = "13px";
      row.appendChild(td);
    });
    ui.skipDetailsBody.appendChild(row);
  });
}

function ensureCopyButtons() {
  if (ui.copyVerifyBtn) return;
  copyBtn.textContent = "MRB만 복사";
  const actions = copyBtn.parentNode;
  const verifyBtn = document.createElement("button");
  verifyBtn.id = "copyVerifyBtn";
  verifyBtn.textContent = "검증용 복사";
  verifyBtn.className = "secondary";
  verifyBtn.disabled = true;
  if (actions) {
    actions.appendChild(verifyBtn);
  }
  verifyBtn.addEventListener("click", handleVerifyCopy);
  ui.copyVerifyBtn = verifyBtn;
}

let latestResult = null;

function ensureArrivalDateInput() {
  if (ui.arrivalInput) return;
  const inputCard = parseBtn.closest(".card");
  if (!inputCard) return;
  const field = document.createElement("label");
  field.className = "field";
  const label = document.createElement("span");
  label.textContent = "입고예정일자 (필수)";
  const input = document.createElement("input");
  input.type = "date";
  input.id = "arrivalDate";
  const hint = document.createElement("small");
  hint.textContent = "입고예정일자를 먼저 선택한 후 PDF/엑셀 파일을 업로드하세요.";
  hint.style.display = "block";
  hint.style.marginTop = "6px";
  hint.style.color = "#555";
  field.appendChild(label);
  field.appendChild(input);
  field.appendChild(hint);
  const pdfField = pdfInput.closest(".field");
  if (pdfField && pdfField.parentNode) {
    pdfField.parentNode.insertBefore(field, pdfField);
  } else {
    inputCard.appendChild(field);
  }
  ui.arrivalInput = input;
  ui.arrivalHint = hint;
}

function extractArrivalDateDetails(pages) {
  const dateToPages = new Map();
  const regex = /입고예정일자[:：]?(\d{4}[./-]\d{2}[./-]\d{2})/g;
  pages.forEach(({ text, pageNo }) => {
    const compact = String(text || "").replace(/\s+/g, "");
    let match;
    while ((match = regex.exec(compact)) !== null) {
      const raw = match[1];
      const normalized = raw.replace(/[./]/g, "-");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) continue;
      if (!dateToPages.has(normalized)) {
        dateToPages.set(normalized, new Set());
      }
      dateToPages.get(normalized).add(pageNo);
    }
  });
  return { dateToPages, dates: new Set(dateToPages.keys()) };
}

function formatPageRanges(pages) {
  const sorted = Array.from(pages).sort((a, b) => a - b);
  if (!sorted.length) return "-";
  const ranges = [];
  let start = sorted[0];
  let prev = sorted[0];
  for (let i = 1; i < sorted.length; i += 1) {
    const current = sorted[i];
    if (current === prev + 1) {
      prev = current;
      continue;
    }
    if (start === prev) {
      ranges.push(`${start}페이지`);
    } else {
      ranges.push(`${start}-${prev}페이지`);
    }
    start = current;
    prev = current;
  }
  if (start === prev) {
    ranges.push(`${start}페이지`);
  } else {
    ranges.push(`${start}-${prev}페이지`);
  }
  return ranges.join(", ");
}

function updateInputAvailability() {
  const hasDate = !!(ui.arrivalInput && ui.arrivalInput.value);
  pdfInput.disabled = !hasDate;
  excelInput.disabled = !hasDate;
  const hasFiles = !!(pdfInput.files?.length && excelInput.files?.length);
  parseBtn.disabled = !(hasDate && hasFiles);
  copyBtn.disabled = true;
  if (ui.copyVerifyBtn) ui.copyVerifyBtn.disabled = true;
}

async function handleParse() {
  showErrors([]);
  showWarnings([]);
  resetOutput();
  ensureCopyButtons();
  ensureArrivalDateInput();
  updateInputAvailability();

  const selectedArrivalDate = ui.arrivalInput ? ui.arrivalInput.value : "";
  if (!selectedArrivalDate) {
    showErrors(["입고예정일자를 선택해야 분석할 수 있습니다."]);
    setStatus("검증 실패: 입고예정일자 미선택");
    return;
  }
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
    setStats({
      totalExcelRows: rows.length,
      totalPdfPages: pages.length,
      totalPdfRecords: 0,
      matchedCount: 0,
      skippedCount: 0,
      fatalErrorCount: 0,
      unusedPdfCount: 0
    });
    if (errors.length) {
      showErrors(errors);
      setStats({
        totalExcelRows: rows.length,
        totalPdfPages: pages.length,
        totalPdfRecords: 0,
        matchedCount: 0,
        skippedCount: 0,
        fatalErrorCount: errors.length,
        unusedPdfCount: 0
      });
      setStatus("검증 실패: 엑셀 입력 오류");
      return;
    }
    if (!rows.length) {
      showErrors(["엑셀에서 물류센터/C/T NO. 목록을 찾지 못했습니다."]);
      setStats({
        totalExcelRows: 0,
        totalPdfPages: pages.length,
        totalPdfRecords: 0,
        matchedCount: 0,
        skippedCount: 0,
        fatalErrorCount: 1,
        unusedPdfCount: 0
      });
      setStatus("검증 실패: 엑셀 입력 오류");
      return;
    }

    const { dateToPages, dates } = extractArrivalDateDetails(pages);
    if (dates.size === 0) {
      showErrors(["PDF에서 입고예정일자를 찾지 못했습니다. 올바른 밀크런 운송장 PDF인지 확인하세요."]);
      setStats({
        totalExcelRows: rows.length,
        totalPdfPages: pages.length,
        totalPdfRecords: 0,
        matchedCount: 0,
        skippedCount: 0,
        fatalErrorCount: 1,
        unusedPdfCount: 0
      });
      setStatus("검증 실패: 입고예정일자 확인");
      return;
    }
    const hasMismatch = dates.size > 1 || !dates.has(selectedArrivalDate);
    if (hasMismatch) {
      const summary = Array.from(dateToPages.entries()).map(([date, pagesSet]) => {
        const pageInfo = formatPageRanges(pagesSet);
        return `PDF 날짜 ${date}: ${pageInfo}`;
      });
      showErrors([
        "입고예정일 불일치로 분석 중단",
        `선택한 날짜: ${selectedArrivalDate}`,
        ...summary
      ]);
      setStats({
        totalExcelRows: rows.length,
        totalPdfPages: pages.length,
        totalPdfRecords: 0,
        matchedCount: 0,
        skippedCount: 0,
        fatalErrorCount: 1,
        unusedPdfCount: 0
      });
      setStatus("검증 실패: 입고예정일 불일치");
      return;
    }

    const { map: pdfMap, issues: pdfIssues } = extractPdfMap(pages);
    setStats({
      totalExcelRows: rows.length,
      totalPdfPages: pages.length,
      totalPdfRecords: pdfMap.size,
      matchedCount: 0,
      skippedCount: 0,
      fatalErrorCount: 0,
      unusedPdfCount: 0
    });
    if (pdfIssues.length) {
      showErrors(pdfIssues);
      setStats({
        totalExcelRows: rows.length,
        totalPdfPages: pages.length,
        totalPdfRecords: pdfMap.size,
        matchedCount: 0,
        skippedCount: 0,
        fatalErrorCount: pdfIssues.length,
        unusedPdfCount: 0
      });
      setStatus("검증 실패: PDF에서 문제가 발견되었습니다.");
      return;
    }
    if (pdfMap.size === 0) {
      showErrors(["PDF에서 운송장 번호를 찾지 못했습니다. PDF 형식을 확인해 주세요."]);
      setStats({
        totalExcelRows: rows.length,
        totalPdfPages: pages.length,
        totalPdfRecords: 0,
        matchedCount: 0,
        skippedCount: 0,
        fatalErrorCount: 1,
        unusedPdfCount: 0
      });
      setStatus("검증 실패: PDF 인식 오류");
      return;
    }

    const { outputLines, issues: outputIssues, skips, matched, usedMrbs } = buildOutput(rows, pdfMap);
    if (outputIssues.length) {
      showErrors(outputIssues);
      setStats({
        totalExcelRows: rows.length,
        totalPdfPages: pages.length,
        totalPdfRecords: pdfMap.size,
        matchedCount: 0,
        skippedCount: 0,
        fatalErrorCount: outputIssues.length,
        unusedPdfCount: 0
      });
      setStatus("검증 실패: 엑셀/PDF 매칭 오류");
      return;
    }

    const unusedPdf = Array.from(pdfMap.values()).filter((mrb) => !usedMrbs.has(mrb));
    updateUnusedBox(unusedPdf);
    updateMatchedTable(matched);
    updateSkipSection(skips);
    setStats({
      totalExcelRows: rows.length,
      totalPdfPages: pages.length,
      totalPdfRecords: pdfMap.size,
      matchedCount: outputLines.length,
      skippedCount: skips.length,
      fatalErrorCount: 0,
      unusedPdfCount: unusedPdf.length
    });

    if (outputLines.length === 0) {
      resultText.value = "";
      resultSection.hidden = false;
      copyBtn.disabled = true;
      if (ui.copyVerifyBtn) ui.copyVerifyBtn.disabled = true;
      setStatus("출력할 MRB가 없습니다(PDF에 해당 센터/CT 없음).");
      return;
    }

    const output = outputLines.join("\n");
    resultText.value = output;
    resultSection.hidden = false;
    copyBtn.disabled = false;
    if (ui.copyVerifyBtn) ui.copyVerifyBtn.disabled = false;
    latestResult = { matched };
    setStatus(`완료되었습니다. 매칭 ${outputLines.length}건`);
  } catch (err) {
    showErrors([err instanceof Error ? err.message : "처리 중 오류가 발생했습니다."]);
    setStats({
      totalExcelRows: 0,
      totalPdfPages: 0,
      totalPdfRecords: 0,
      matchedCount: 0,
      skippedCount: 0,
      fatalErrorCount: 1,
      unusedPdfCount: 0
    });
    setStatus("검증 실패: 입력을 확인해 주세요.");
  }
}

async function handleCopy() {
  if (!resultText.value) return;
  await navigator.clipboard.writeText(resultText.value);
  setStatus("복사가 완료되었습니다.");
}

async function handleVerifyCopy() {
  if (!latestResult || !latestResult.matched?.length) return;
  const lines = latestResult.matched.map((item) => {
    return `${item.fcRaw}\t${item.ct}\t${item.mrb}`;
  });
  await navigator.clipboard.writeText(lines.join("\n"));
  setStatus("검증용 복사가 완료되었습니다.");
}

parseBtn.addEventListener("click", handleParse);
copyBtn.addEventListener("click", handleCopy);

ensureArrivalDateInput();
updateInputAvailability();

if (ui.arrivalInput) {
  ui.arrivalInput.addEventListener("change", () => {
    if (!ui.arrivalInput.value) {
      pdfInput.value = "";
      excelInput.value = "";
      clearExcelState();
      resetOutput();
    }
    updateInputAvailability();
  });
}

excelInput.addEventListener("change", async () => {
  resetOutput();
  showErrors([]);
  showWarnings([]);
  updateInputAvailability();
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

pdfInput.addEventListener("change", () => {
  resetOutput();
  showErrors([]);
  showWarnings([]);
  updateInputAvailability();
});
