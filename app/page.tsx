"use client";

import { useMemo, useState } from "react";
import { z } from "zod";

const ResultRowSchema = z.object({
  fc: z.string().min(1),
  mrbs: z.array(z.string().min(1))
});

const PageRecordSchema = z.object({
  pageNo: z.number().int().positive(),
  fc: z.string().min(1),
  mrb: z.string().min(1),
  seq: z.string().regex(/^\d{3}$/)
});

const SuccessResponseSchema = z.object({
  data: z.array(ResultRowSchema),
  records: z.array(PageRecordSchema)
});

const ErrorIssueSchema = z.object({
  pageNo: z.number().int().positive().nullable(),
  code: z.string().min(1),
  message: z.string().min(1)
});

const ErrorResponseSchema = z.object({
  error: z.string().min(1),
  issues: z.array(ErrorIssueSchema)
});

type ResultRow = z.infer<typeof ResultRowSchema>;
type ValidationIssue = z.infer<typeof ErrorIssueSchema>;

function escapeForSeparated(value: string, separator: string) {
  const needsQuote = value.includes("\n") || value.includes("\r") || value.includes("\t") || value.includes(separator) || value.includes("\"");
  const escaped = value.replace(/"/g, '""');
  return needsQuote ? `"${escaped}"` : escaped;
}

export default function Home() {
  const [results, setResults] = useState<ResultRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [issues, setIssues] = useState<ValidationIssue[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);

  const hasResults = results.length > 0;

  const tsvText = useMemo(() => {
    if (!hasResults) return "";
    return results
      .map((row) => {
        const mrbText = row.mrbs.join("\n");
        return `${escapeForSeparated(row.fc, "\t")}\t${escapeForSeparated(mrbText, "\t")}`;
      })
      .join("\n");
  }, [results, hasResults]);

  const csvText = useMemo(() => {
    if (!hasResults) return "";
    const header = "FC,운송장번호";
    const body = results
      .map((row) => {
        const mrbText = row.mrbs.join("\n");
        return `${escapeForSeparated(row.fc, ",")},${escapeForSeparated(mrbText, ",")}`;
      })
      .join("\n");
    return `${header}\n${body}`;
  }, [results, hasResults]);

  async function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setLoading(true);
    setError(null);
    setResults([]);
    setIssues([]);

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/parse", {
        method: "POST",
        body: formData
      });

      const payload = await response.json().catch(() => null);

      if (!response.ok) {
        const parsedError = ErrorResponseSchema.safeParse(payload);
        if (parsedError.success) {
          setError(parsedError.data.error);
          setIssues(parsedError.data.issues);
          return;
        }
        throw new Error("서버 오류가 발생했습니다.");
      }

      const parsed = SuccessResponseSchema.safeParse(payload);
      if (!parsed.success) {
        throw new Error("응답 형식이 올바르지 않습니다.");
      }
      setResults(parsed.data.data);
      setIssues([]);
    } catch (err) {
      const message = err instanceof Error ? err.message : "처리 중 오류가 발생했습니다.";
      setError(message);
      setIssues([]);
    } finally {
      setLoading(false);
    }
  }

  async function handleCopy() {
    if (!hasResults) return;
    await navigator.clipboard.writeText(tsvText);
  }

  function handleDownloadCsv() {
    if (!hasResults) return;
    const blob = new Blob([csvText], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "milkrun-results.csv";
    link.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main>
      <header>
        <h1>밀크런 운송장 PDF 파서</h1>
        <p className="subtitle">PDF 업로드 → FC별 MRB 운송장번호 추출</p>
      </header>

      <section className="card">
        <input
          type="file"
          accept="application/pdf"
          onChange={handleFileChange}
        />
        <div className="notice">
          {loading && "PDF를 분석 중입니다..."}
          {!loading && fileName && `선택된 파일: ${fileName}`}
          {!loading && !fileName && "밀크런 운송장 PDF를 업로드하세요."}
        </div>
        {error && (
          <div className="error-block">
            <div className="error">{error}</div>
            {issues.length > 0 && (
              <div className="error-list">
                {issues.map((issue, index) => (
                  <div className="error-item" key={`${issue.code}-${issue.pageNo ?? "all"}-${index}`}>
                    <strong>{issue.pageNo ? `페이지 ${issue.pageNo}` : "전체"}</strong>: {issue.message}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        <div className="actions">
          <button onClick={handleCopy} disabled={!hasResults}>
            복사
          </button>
          <button className="secondary" onClick={handleDownloadCsv} disabled={!hasResults}>
            CSV 다운로드
          </button>
        </div>
      </section>

      {hasResults && (
        <section className="card" style={{ marginTop: 24 }}>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>FC</th>
                  <th>운송장번호</th>
                </tr>
              </thead>
              <tbody>
                {results.map((row) => (
                  <tr key={row.fc}>
                    <td className="fc">{row.fc}</td>
                    <td className="mrb">{row.mrbs.join("\n")}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="notice">운송장번호가 여러 개면 줄바꿈으로 표시됩니다.</div>
        </section>
      )}
    </main>
  );
}
