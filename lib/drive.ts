import { google } from "googleapis";
import { getGoogleAuth } from "./google";
import { Readable } from "stream";

async function getDrive() {
  const auth = getGoogleAuth();
  const client = await auth.getClient();
  return google.drive({ version: "v3", auth: client as never });
}

/** Drive 파일을 Buffer로 다운로드 */
export async function downloadFileAsBuffer(fileId: string): Promise<Buffer> {
  const drive = await getDrive();
  const res = await drive.files.get(
    { fileId, alt: "media" },
    { responseType: "arraybuffer" }
  );
  return Buffer.from(res.data as ArrayBuffer);
}

/** Buffer를 Drive 폴더에 업로드, 공유 링크 반환 */
export async function uploadToDrive(
  fileName: string,
  buffer: Buffer,
  mimeType: string,
  folderId: string
): Promise<{ id: string; webViewLink: string }> {
  const drive = await getDrive();
  const res = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType, body: Readable.from(buffer) },
    fields: "id,webViewLink",
  });
  // 링크 공개 (뷰어)
  await drive.permissions.create({
    fileId: res.data.id!,
    requestBody: { role: "reader", type: "anyone" },
  });
  return { id: res.data.id!, webViewLink: res.data.webViewLink! };
}

/** 파일 메타데이터만 조회 */
export async function getFileMeta(fileId: string) {
  const drive = await getDrive();
  const res = await drive.files.get({ fileId, fields: "id,name,webViewLink" });
  return res.data;
}
