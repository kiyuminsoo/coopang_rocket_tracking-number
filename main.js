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

const PDF_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;
}

function normalizeWhitespace(value) {
  return value.replace(/\s+/g, " ").trim();
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

function extractFcFromLines(lines) {
  const indices = [];
  for (let i = 0; i < lines.length; i += 1) {
    if (lines[i].includes("받는 사람:")) {
      indices.push(i);
    }
  }
  if (indices.length !== 1) {
    return { fc: null, count: indices.length };
  }
  const idx = indices[0];
  const nextLine = lines[idx + 1] ?? "";
  const fc = normalizeWhitespace(nextLine);
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
  const mrbRegex = /MRB[0-9]+-[0-9]{3}/g;

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

function buildFcOrderFromExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("엑셀 파일을 읽을 수 없습니다."));
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        const workbook = window.XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets["<패킹> F_260205"];
        if (!sheet) {
          reject(new Error("엑셀 시트 '<패킹> F_260205'을 찾을 수 없습니다."));
          return;
        }
        const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
        let headerRowIndex = -1;
        let headerColIndex = -1;

        for (let r = 0; r < rows.length; r += 1) {
          const row = rows[r] || [];
          for (let c = 0; c < row.length; c += 1) {
            const value = String(row[c] ?? "").trim();
            if (!value) continue;
            if (/물류센터|센터|FC/i.test(value)) {
              headerRowIndex = r;
              headerColIndex = c;
              break;
            }
          }
          if (headerRowIndex !== -1) break;
        }

        if (headerRowIndex === -1) {
          reject(new Error("엑셀에서 '물류센터/센터/FC' 헤더를 찾지 못했습니다."));
          return;
        }

        const fcOrder = [];
        const seen = new Set();

        for (let r = headerRowIndex + 1; r < rows.length; r += 1) {
          const value = String((rows[r] || [])[headerColIndex] ?? "").trim();
          if (!value) continue;
          const fc = normalizeWhitespace(value);
          if (!seen.has(fc)) {
            seen.add(fc);
            fcOrder.push(fc);
          }
        }

        resolve(fcOrder);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function buildOutput(fcOrder, records, showFc) {
  const errors = [];
  const byFc = new Map();

  records.forEach((record) => {
    const list = byFc.get(record.fc) ?? [];
    list.push(record);
    byFc.set(record.fc, list);
  });

  // 동일 FC에서 동일 boxNo 중복 방지
  byFc.forEach((list, fc) => {
    const seenBox = new Map();
    list.forEach((record) => {
      if (seenBox.has(record.boxNo)) {
        const prev = seenBox.get(record.boxNo);
        errors.push(
          `FC ${fc}: 박스번호 ${record.boxNo}가 중복되었습니다. (페이지 ${prev.pageNo}, ${record.pageNo})`
        );
      } else {
        seenBox.set(record.boxNo, record);
      }
    });
  });

  fcOrder.forEach((fc) => {
    if (!byFc.has(fc)) {
      errors.push(`엑셀에 있는 FC ${fc}가 PDF에서 발견되지 않았습니다.`);
    }
  });

  if (errors.length) {
    return { output: "", errors };
  }

  const lines = [];
  fcOrder.forEach((fc) => {
    const list = byFc.get(fc) || [];
    const sorted = [...list].sort((a, b) => a.boxNo - b.boxNo);
    if (showFc) {
      lines.push(`FC: ${fc}`);
    }
    sorted.forEach((record) => {
      lines.push(record.mrb);
    });
  });

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
    const [fcOrder, pages] = await Promise.all([
      buildFcOrderFromExcel(excelInput.files[0]),
      parsePdf(pdfInput.files[0])
    ]);
    if (!fcOrder.length) {
      showErrors(["엑셀에서 물류센터 목록을 찾지 못했습니다."]);
      setStatus("검증 실패: 엑셀 입력 오류");
      return;
    }

    const { records, issues } = extractRecordsFromPdf(pages);
    if (issues.length) {
      showErrors(issues);
      setStatus("검증 실패: PDF에서 문제가 발견되었습니다.");
      return;
    }

    const { output, errors } = buildOutput(fcOrder, records, showFcToggle.checked);
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
