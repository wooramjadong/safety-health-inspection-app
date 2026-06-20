/**
 * POST /api/action-submit
 * body: { token: string }
 * → 조치 완료 표시 (Sheets 상태 업데이트)
 * → 최종본 PPTX 생성 (Phase 2에서 사진 삽입 포함)
 */
import { NextRequest, NextResponse } from "next/server";
import { getInspectionByToken, updateInspectionUrls } from "@/lib/sheets";

export async function POST(req: NextRequest) {
  try {
    const { token } = await req.json();
    if (!token) return NextResponse.json({ error: "token required" }, { status: 400 });

    const insp = await getInspectionByToken(token);
    if (!insp) return NextResponse.json({ error: "invalid token" }, { status: 404 });

    // Phase 1: 상태만 완료로 표시
    // Phase 2에서 최종본 PPTX(사진 포함) 자동 생성 추가 예정
    await updateInspectionUrls(
      insp.id,
      token,
      insp.pptxUrl ?? "",
      insp.xlsxUrl ?? "",
      "조치완료"
    );

    return NextResponse.json({ success: true });
  } catch (e: any) {
    console.error("[action-submit] error", e);
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
