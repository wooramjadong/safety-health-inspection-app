import { randomUUID } from "crypto";

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

export async function createInspection(input: { type: string; siteName: string; inspectors: string }) {
  const inspectionId = `INS-${Date.now()}`;
  const token = randomUUID();
  // TODO: Google Sheets append 연동
  return { ok: true, sheetId: SHEET_ID, inspectionId, actionToken: token, ...input };
}

export async function saveActionSubmit(input: { token: string; actionText: string }) {
  // TODO: Drive 이미지 업로드 + 조치관리 시트 append 연동
  return { ok: true, status: "submitted", ...input };
}
