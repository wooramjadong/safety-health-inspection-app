/**
 * POST /api/import-historical
 * multipart/form-data: { file: File, type: "정기평가" | "준수평가" }
 *
 * 과거 점검 파일을 파싱해 보여주기만 한다 (DB 저장은 confirm 엔뛜에서).
 * 정기평가: xlsx 파일을 보내야 한다 (현장부문/서류부문 시트에 이미 정리된 데이타 활용)
 * 준수평가: pptx 파일을 보내야 한다
 */
import { NextRequest, NextResponse } from "next/server";
import { extractRegularFromXlsx, extractSapaFromPptx } from "@/lib/import";

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();
    const file = form.get("file") as File | null;
    const type = form.get("type") as string | null;

    if (!file || !type) {
      return NextResponse.json({ error: "file, type required" }, { status: 400 });
    }

    const buf = Buffer.from(await file.arrayBuffer());

    if (type === "정기평가") {
      if (!file.name.endsWith(".xlsx")) {
        return NextResponse.json(
          { error: "정기평가는 xlsx 파일을 업로드해주세요 (현장부문/서류부문 시트에 이미 정리된 지적사항이 있습니다)" },
          { status: 400 }
        );
      }
      const result = await extractRegularFromXlsx(buf);
      return NextResponse.json({ type, data: result });
    } else if (type === "준수평가") {
      if (!file.name.endsWith(".pptx")) {
        return NextResponse.json({ error: "준수평가는 pptx 파일을 업로드해주세요" }, { status: 400 });
      }
      const result = await extractSapaFromPptx(buf);
      return NextResponse.json({ type, data: result });
    }

    return NextResponse.json({ error: "type은 정기평가 또는 준수평가여야 합니다" }, { status: 400 });
  } catch (e: any) {
    console.error("[import-historical] error", e);
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
