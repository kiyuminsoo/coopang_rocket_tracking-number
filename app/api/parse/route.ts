import pdfParse from "pdf-parse";
import { NextResponse } from "next/server";
import { z } from "zod";

export const runtime = "nodejs";

const fcRegex = /받는 사람:\s*\n([^\n\r]+)/g;
const mrbRegex = /MRB[0-9]+-[0-9]{3}/g;
const seqRegex = /밀크런 운송장\s*(\d{3})/g;

const PageRecordSchema = z.object({
  pageNo: z.number().int().positive(),
  fc: z.string().min(1),
  mrb: z.string().min(1),
  seq: z.string().regex(/^\d{3}$/)
});

const ResultRowSchema = z.object({
  fc: z.string().min(1),
  mrbs: z.array(z.string().min(1))
});

const ErrorIssueSchema = z.object({
  pageNo: z.number().int().positive().nullable(),
  code: z.string().min(1),
  message: z.string().min(1)
});

const SuccessResponseSchema = z.object({
  data: z.array(ResultRowSchema),
  records: z.array(PageRecordSchema)
});

const ErrorResponseSchema = z.object({
  error: z.string().min(1),
  issues: z.array(ErrorIssueSchema)
});

type PageRecord = z.infer<typeof PageRecordSchema>;

function normalizeWhitespace(value: string) {
  return value.replace(/\s+/g, " ").trim();
}

function buildPageText(items: any[]) {
  let buffer = "";
  for (const item of items) {
    const text = typeof item.str === "string" ? item.str : "";
    buffer += text;
    buffer += item.hasEOL ? "\n" : " ";
  }
  return buffer.replace(/[ \t]+\n/g, "\n").replace(/\n{2,}/g, "\n").trim();
}

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: "파일이 없습니다." }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const pages: string[] = [];

    await pdfParse(buffer, {
      pagerender: async (pageData) => {
        const textContent = await pageData.getTextContent();
        const pageText = buildPageText(textContent.items);
        pages.push(pageText);
        return pageText;
      }
    });

    const issues: z.infer<typeof ErrorIssueSchema>[] = [];
    const records: PageRecord[] = [];
    const mrbToFc = new Map<string, { fc: string; pageNo: number }>();

    pages.forEach((pageText, index) => {
      const pageNo = index + 1;
      const fcMatches = Array.from(pageText.matchAll(fcRegex));
      const mrbMatches = pageText.match(mrbRegex) ?? [];
      const seqMatches = Array.from(pageText.matchAll(seqRegex));

      if (fcMatches.length !== 1) {
        issues.push({
          pageNo,
          code: "FC_COUNT",
          message: `FC는 페이지당 1개여야 합니다. (검출: ${fcMatches.length}개)`
        });
      }

      if (mrbMatches.length !== 1) {
        issues.push({
          pageNo,
          code: "MRB_COUNT",
          message: `MRB는 페이지당 1개여야 합니다. (검출: ${mrbMatches.length}개)`
        });
      }

      if (seqMatches.length !== 1) {
        issues.push({
          pageNo,
          code: "SEQ_COUNT",
          message: `SEQ는 '밀크런 운송장 001'에서 1개만 추출되어야 합니다. (검출: ${seqMatches.length}개)`
        });
      }

      if (fcMatches.length !== 1 || mrbMatches.length !== 1 || seqMatches.length !== 1) {
        return;
      }

      const fcLine = normalizeWhitespace(fcMatches[0][1]);
      if (!fcLine) {
        issues.push({
          pageNo,
          code: "FC_EMPTY",
          message: "FC가 비어 있습니다."
        });
        return;
      }

      const mrb = mrbMatches[0];
      const seq = seqMatches[0][1];

      if (!mrb.endsWith(`-${seq}`)) {
        issues.push({
          pageNo,
          code: "SEQ_MRB_MISMATCH",
          message: `MRB(${mrb})의 끝 3자리와 SEQ(${seq})가 일치하지 않습니다.`
        });
        return;
      }

      const existing = mrbToFc.get(mrb);
      if (existing && existing.fc !== fcLine) {
        issues.push({
          pageNo,
          code: "MRB_FC_CONFLICT",
          message: `MRB(${mrb})가 다른 FC에 중복 등장합니다. (기존: ${existing.fc}, 페이지 ${existing.pageNo})`
        });
        return;
      }
      if (!existing) {
        mrbToFc.set(mrb, { fc: fcLine, pageNo });
      }

      records.push({
        pageNo,
        fc: fcLine,
        mrb,
        seq
      });
    });

    if (records.length !== pages.length) {
      issues.push({
        pageNo: null,
        code: "PAGE_RECORD_MISMATCH",
        message: `전체 페이지(${pages.length})와 레코드(${records.length}) 수가 일치하지 않습니다.`
      });
    }

    const recordsByFc = new Map<string, PageRecord[]>();
    for (const record of records) {
      const list = recordsByFc.get(record.fc) ?? [];
      list.push(record);
      recordsByFc.set(record.fc, list);
    }

    for (const [fc, list] of recordsByFc.entries()) {
      const sorted = [...list].sort((a, b) => Number(a.seq) - Number(b.seq));
      for (let i = 1; i < sorted.length; i += 1) {
        const prev = sorted[i - 1];
        const current = sorted[i];
        if (Number(current.seq) !== Number(prev.seq) + 1) {
          issues.push({
            pageNo: current.pageNo,
            code: "SEQ_GAP",
            message: `FC(${fc})의 SEQ가 연속적이지 않습니다. (이전: ${prev.seq} @페이지 ${prev.pageNo}, 현재: ${current.seq} @페이지 ${current.pageNo})`
          });
          break;
        }
      }
    }

    if (issues.length > 0) {
      const payload = ErrorResponseSchema.parse({
        error: "검증 실패",
        issues
      });
      return NextResponse.json(payload, { status: 422 });
    }

    const data = Array.from(recordsByFc.entries()).map(([fc, list]) => {
      const sorted = [...list].sort((a, b) => Number(a.seq) - Number(b.seq));
      return {
        fc,
        mrbs: sorted.map((record) => record.mrb)
      };
    });

    const payload = SuccessResponseSchema.parse({ data, records });
    return NextResponse.json(payload);
  } catch (error) {
    console.error(error);
    return NextResponse.json({ error: "처리 중 오류가 발생했습니다." }, { status: 500 });
  }
}
