/**
 * POST /api/generate
 * body: { inspectionId: string, matchOverrides?: Record<string, number> }
 *
 * matchOverrides: { [findingId]: 체크리스트 row 번호 } — /inspection/[id]/review 화맩에서
 * 사용자가 AI 매징을 사전 확인·수정한 경우 전도되어, 실제 생성 시에는 다시 AI
 * 호다하지 않고 사용자가 지정한 행을 그대로 쓴다.
 *
 * 1) Sheets에서 점검 + 지적사항 조회
 * 2) Gemini로 텍스트 요약 (현장부문 지적사항만 → 별첨 슬라이드 대상)
 * 3) 템플릿 PPTX/xlsx 다운로드 → 데이타 채움
 *    (xlsx 평가결과 시트의 총점/점수 수식은 체크리스트 입력에 따라 자동 계산되목로 직접 쓰지 않아.
 *    보정계수(O13)도 동일하게 하위 3건(P15/P16/P17)를 전도해 수식이 자동 계산하게 한다.)
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
    const { inspectionId, matchOverrides } = await req.json();
    if (!inspectionId) return NextResponse.json({ error: "inspectionId required" }, { status: 400 });
    const overrides: Record<string, number> = matchOverrides ?? {};

    // 1. 데이타 조회
    const allInspections = await getInspections();
    const insp = allInspections.find((i: any) => i.id === inspectionId);
    if (!insp) return NextResponse.json({ error: "inspection not found" }, { status: 404 });

    const rawFindings = await getFindingsByInspection(inspectionId);
    const fieldRawFindings = rawFindings.filter((f: any) => f.section === "현장");
    const docRawFindings = rawFindings.filter((f: any) => f.section === "서류");

    // 2. Gemini 요약 — 현장부문 지적사항만 별첨 슬라이드로 생성·요약
    const findings: Finding[] = await Promise.all(
      fieldRawFindings.map(async (f: any, idx: number) => {
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
      // PPTX (별첨 슬라이드는 현장부문 지적사항만 대상)
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

      // xlsx — findings(이미 요약된 텍스트)를 재사용해 PPTX와 100% 동일한 내용이 들어가게 하고,
      // review 화맩에서 사용자가 확정한 matchOverrides는 그대로 적용.
      const xlsxTpl = await downloadFileAsBuffer(REGULAR_XLSX_ID);
      const xlsxBuf = await generateRegularXlsx(xlsxTpl, {
        siteName: insp.siteName,
        inspectionPeriod: `${insp.inspectionStart} ~ ${insp.inspectionEnd}`,
        inspectors: insp.inspectors ?? "",
        mainWork: insp.mainWork ?? "",
        amount: insp.amount ?? "",
        constructionPeriod: insp.constructionPeriod ?? "",
        managerInfo: `${insp.siteManager ?? ""} ${insp.safetyManager ?? ""}`.trim(),
        monthlyProgressFactor: insp.monthlyProgressFactor ?? "1",
        riskWorkFactor: insp.riskWorkFactor ?? "1",
        helperOperationFactor: insp.helperOperationFactor ?? "1",
        fieldFindings: findings.map((f, i) => ({
          content: f.content,
          grade: f.grade,
          matchedRow: overrides[fieldRawFindings[i]?.id],
        })),
        docFindings: docRawFindings.map((f: any) => ({
          content: f.content,
          matchedRow: overrides[f.id],
        })),
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
