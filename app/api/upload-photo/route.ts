/**
 * POST /api/upload-photo
 * multipart/form-data: { file: File, token: string, findingId: string }
 * → Drive 03_조치사진/{token}/ 에 업로드 → URL 반환
 */
import { NextRequest, NextResponse } from "next/server";
import { uploadToDrive } from "@/lib/drive";
import { getInspectionByToken, createAction } from "@/lib/sheets";

const PHOTO_FOLDER = process.env.DRIVE_ACTION_PHOTO_FOLDER_ID!;

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();
    const file = form.get("file") as File | null;
    const token = form.get("token") as string | null;
    const findingId = form.get("findingId") as string | null;
    const comment = (form.get("comment") as string) ?? "";

    if (!file || !token || !findingId) {
      return NextResponse.json({ error: "file, token, findingId required" }, { status: 400 });
    }

    // 토큰으로 점검 조회
    const insp = await getInspectionByToken(token);
    if (!insp) return NextResponse.json({ error: "invalid token" }, { status: 404 });

    const bytes = await file.arrayBuffer();
    const buf = Buffer.from(bytes);
    const timestamp = Date.now();
    const fileName = `${token}_${findingId}_${timestamp}_${file.name}`;

    const result = await uploadToDrive(fileName, buf, file.type, PHOTO_FOLDER);

    // 조치 기록 저장
    await createAction({
      inspectionId: insp.id,
      findingId,
      photoUrl: result.webViewLink,
      comment,
      uploadedAt: new Date().toISOString(),
    });

    return NextResponse.json({ url: result.webViewLink });
  } catch (e: any) {
    console.error("[upload-photo] error", e);
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
