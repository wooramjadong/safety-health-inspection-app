/**
 * POST /api/match-preview
 * body: { inspectionId: string }
 *
 * 지적사항도 PPTX/xlsx를 생성하기 전, AI(Gemini)가 체크리스트 어느 항목에
 * 매징했는지 보여좼고 사용자가 확인/수정할 수 있게 해주는 추계 API.
 */
import { NextRequest, NextResponse } from "next/server";
import { getInspections, getFindingsByInspection } from "@/lib/sheets";
import { summarizeForCell, CELL_LIMITS, matchChecklistRow } from "@/lib/gemini";
import { FIELD_CHECKLIST, DOC_CHECKLIST } from "@/lib/checklist-data";

export async function POST(req: NextRequest) {
  try {
    const { inspectionId } = await req.json();
    if (!inspectionId) return NextResponse.json({ error: "inspectionId required" }, { status: 400 });

    const allInspections = await getInspections();
    const insp = allInspections.find((i: any) => i.id === inspectionId);
    if (!insp) return NextResponse.json({ error: "inspection not found" }, { status: 404 });

    if (insp.type !== "정기평가") {
      // 준수평가는 체크리스트 매징이 없이 분롵 없이 보여도 돔다
      return NextResponse.json({ fieldMatches: [], docMatches: [], fieldChecklist: [], docChecklist: [] });
    }

    const rawFindings = await getFindingsByInspection(inspectionId);
    const fieldRaw = rawFindings.filter((f: any) => f.section === "현장");
    const docRaw = rawFindings.filter((f: any) => f.section === "서류");

    const fieldMatches = await Promise.all(
      fieldRaw.map(async (f: any) => {
        const content = await summarizeForCell(f.content, CELL_LIMITS.별첨_내용);
        const row = await matchChecklistRow(content, FIELD_CHECKLIST);
        const matched = FIELD_CHECKLIST.find((c) => c.row === row);
        return {
          findingId: f.id,
          content,
          grade: f.grade ?? "미흡",
          matchedRow: row,
          matchedGroup: matched?.group ?? "",
          matchedItem: matched?.item ?? "",
        };
      })
    );

    const docMatches = await Promise.all(
      docRaw.map(async (f: any) => {
        const row = await matchChecklistRow(f.content, DOC_CHECKLIST);
        const matched = DOC_CHECKLIST.find((c) => c.row === row);
        return {
          findingId: f.id,
          content: f.content,
          matchedRow: row,
          matchedGroup: matched?.group ?? "",
          matchedItem: matched?.item ?? "",
        };
      })
    );

    return NextResponse.json({
      fieldMatches,
      docMatches,
      fieldChecklist: FIELD_CHECKLIST,
      docChecklist: DOC_CHECKLIST,
    });
  } catch (e: any) {
    console.error("[match-preview] error", e);
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
