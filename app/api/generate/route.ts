/**
 * POST /api/generate
 * body: { inspectionId: string }
 *
 * 1) Sheets에서 점검 + 지적사항 조회
 * 2) Gemini로 텍스트 요약
 * 3) 템플릿 PPTX/xlsx 다운로드 → 데이터 채움
 * 4) Drive 업로드 → 현장 조치링크용 토큰 업데이트
 * 5) 다운로드 URL 반환
 */
import { NextRequest, NextResponse } from "next/server";
import {
  getInspections,
  getFindingsByInspection,
  updateInspectionUrls,
} from "@/lib/sheets";
import { downloadFileAsBuffer, uploadToDrive } from "@/lib/drive";
import { generateRegularPptx, generateSapaPptx, Finding } from "@/lib/pptx";
import { generateRegularXlsx } from "@/lib/xlsx";
import { summarizeForCell, CELL_LIMITS, detectRiskType } from "@/lib/gemini";
import { randomUUID } from "crypto";

const REGULAR_PPTX_ID = process.env.REGULAR_PPTX_TEMPLATE_ID!;
const REGULAR_XLSX_ID = process.env.REGULAR_XLSX_TEMPLATE_ID!;
const SAPA_PPTX_ID   = process.env.SAPA_PPTX_TEMPLATE_ID!;
const RESULT_FOLDER  = process.env.DRIVE_RESULT_FOLDER_ID!;

export async function POST(req: NextRequest) {
  try {
    const { inspectionId } = await req.json();
    if (!inspectionId) return NextResponse.json({ error: "inspectionId required" }, { status: 400 });

    // 1. 데이터 조회
    const allInspections = await getInspections();
    const insp = allInspections.find((i: any) => i.id === inspectionId);
    if (!insp) return NextResponse.json({ error: "inspection not found" }, { status: 404 });

    const rawFindings = await getFindingsByInspection(inspectionId);

    // 2. Gemini 요약
    const findings: Finding[] = await Promise.all(
      rawFindings.map(async (f: any, idx: number) => {
        const content = await summarizeForCell(f.content, CELL_LIMITS.별첨_내용);
        const actionRequest = await summarizeForCell(f.actionRequest ?? "", CELL_LIMITS.조치요구);
        return {
          seq: idx + 1,
          grade: f.grade === "위험" ? "위험" : "미흡",
          riskType: detectRiskType(f.content),
          content,
          actionRequest,
        };
      })
    );

    const token = randomUUID();
    const timestamp = new Date().toISOString().slice(0, 10);
    const baseName = `${insp.siteName}_${timestamp}`;

    let pptxUrl = "";
    let xlsxUrl = "";

    if (insp.type === "정기평가") {
      // PPTX
      const pptxTpl = await downloadFileAsBuffer(REGULAR_PPTX_ID);
      const pptxBuf = await generateRegularPptx(pptxTpl, {
        siteName: insp.siteName,
        constructionPeriod: insp.constructionPeriod ?? "",
        amount: insp.amount ?? "",
        progress: insp.progress ?? "",
        mainWork: insp.mainWork ?? "",
        inspectionStart: insp.inspectionStart ?? "",
        inspectionEnd: insp.inspectionEnd ?? "",
        inspectors: insp.inspectors ?? "",
        docScore: insp.docScore ?? "",
        fieldScore: insp.fieldScore ?? "",
        findings,
      });
      const pptxResult = await uploadToDrive(
        `${baseName}_정기평가.pptx`,
        pptxBuf,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        RESULT_FOLDER
      );
      pptxUrl = pptxResult.webViewLink;

      // xlsx
      const xlsxTpl = await downloadFileAsBuffer(REGULAR_XLSX_ID);
      const xlsxBuf = await generateRegularXlsx(xlsxTpl, {
        siteName: insp.siteName,
        amount: insp.amount ?? "",
        constructionPeriod: insp.constructionPeriod ?? "",
        inspectionPeriod: `${insp.inspectionStart} ~ ${insp.inspectionEnd}`,
        siteManager: insp.siteManager ?? "",
        safetyManager: insp.safetyManager ?? "",
        docScore: insp.docScore ?? "",
        fieldScore: insp.fieldScore ?? "",
        deduction: insp.deduction ?? "0",
        totalScore: insp.totalScore ?? "",
        docFindings: rawFindings.filter((f: any) => f.section === "서류"),
        fieldFindings: rawFindings.filter((f: any) => f.section === "현장"),
      });
      const xlsxResult = await uploadToDrive(
        `${baseName}_정기평가.xlsx`,
        xlsxBuf,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        RESULT_FOLDER
      );
      xlsxUrl = xlsxResult.webViewLink;

    } else if (insp.type === "준수평가") {
      const pptxTpl = await downloadFileAsBuffer(SAPA_PPTX_ID);
      const pptxBuf = await generateSapaPptx(pptxTpl, {
        siteName: insp.siteName,
        inspectionDate: insp.inspectionStart ?? "",
        inspectors: insp.inspectors ?? "",
        overallResult: insp.overallResult ?? "미흡",
        findings: rawFindings.map((f: any, i: number) => ({
          seq: i + 1,
          itemName: f.itemName ?? "",
          detailContent: f.content ?? "",
          actionRequest: f.actionRequest ?? "",
          result: f.result ?? "미흡",
        })),
      });
      const pptxResult = await uploadToDrive(
        `${baseName}_준수평가.pptx`,
        pptxBuf,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        RESULT_FOLDER
      );
      pptxUrl = pptxResult.webViewLink;
    }

    // 4. 토큰·URL Sheets 업데이트
    await updateInspectionUrls(inspectionId, token, pptxUrl, xlsxUrl);

    const actionLink = `${process.env.NEXT_PUBLIC_BASE_URL}/action/${token}`;
    return NextResponse.json({ pptxUrl, xlsxUrl, actionLink });

  } catch (e: any) {
    console.error("[generate] error", e);
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
