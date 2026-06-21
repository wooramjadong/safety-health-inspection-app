/**
 * POST /api/import-historical/confirm
 * body: {
 *   type: "정기평가" | "준수평가",
 *   inspectionStart: string,  // 점검일자 (파섬이 목 정확히 뇑아내지 못하따 경우 사용자가 지정)
 *   data: ImportedRegularInspection | ImportedSapaInspection (사용자가 확인/수정한 후)
 * }
 *
 * 이목일롸이르 확인된 데이타를 점검(Inspection) + 지적사항(Finding) 로 Sheets DB에 저장한다.
 * 상태는 "점검완료"대신 "과거점검(임포트)"로 구분해 대시대시보드에서 식별되게 한다.
 */
import { NextRequest, NextResponse } from "next/server";
import { createInspection, createFinding } from "@/lib/sheets";

export async function POST(req: NextRequest) {
  try {
    const { type, inspectionStart, data } = await req.json();
    if (!type || !data) return NextResponse.json({ error: "type, data required" }, { status: 400 });

    if (type === "정기평가") {
      const inspectionId = await createInspection({
        type: "정기평가",
        siteName: data.siteName,
        inspectionStart: inspectionStart || data.inspectionPeriod?.split("~")[0]?.trim() || "",
        inspectionEnd: data.inspectionPeriod?.split("~")[1]?.trim() || "",
        inspectors: data.inspectors,
        constructionPeriod: data.constructionPeriod,
        amount: data.amount,
        siteManager: data.managerInfo,
        docScore: data.docTotalScore,
        fieldScore: data.fieldSectionScore,
        status: "과거점검(임포트)",
      });

      let count = 0;
      for (const f of data.fieldFindings ?? []) {
        await createFinding({
          inspectionId, section: "현장", grade: f.grade,
          content: f.content, itemName: `${f.group} / ${f.item}`,
        });
        count++;
      }
      for (const f of data.docFindings ?? []) {
        await createFinding({
          inspectionId, section: "서류",
          content: f.content, itemName: `${f.group} / ${f.item}`,
        });
        count++;
      }

      return NextResponse.json({ inspectionId, importedFindings: count });

    } else if (type === "준수평가") {
      const inspectionId = await createInspection({
        type: "준수평가",
        siteName: data.siteName,
        inspectionStart: inspectionStart || data.inspectionDate || "",
        inspectors: data.inspectors,
        status: "과거점검(임포트)",
      });

      let count = 0;
      for (const f of data.findings ?? []) {
        if (!f.detail) continue;
        await createFinding({
          inspectionId, section: "서류",
          content: f.detail, actionRequest: f.action, itemName: f.itemName,
        });
        count++;
      }

      return NextResponse.json({ inspectionId, importedFindings: count });
    }

    return NextResponse.json({ error: "알 수 없는 type" }, { status: 400 });
  } catch (e: any) {
    console.error("[import-historical/confirm] error", e);
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
