import { google } from "googleapis";
import { getGoogleAuth } from "./google";
import { randomUUID } from "crypto";

const SHEET_ID = process.env.GOOGLE_SHEET_ID!;

async function getSheets() {
  const auth = getGoogleAuth();
  const client = await auth.getClient();
  return google.sheets({ version: "v4", auth: client as never });
}

// ── 공통 ──────────────────────────────────────────────
async function appendRow(sheetName: string, values: (string | number)[]) {
  const s = await getSheets();
  await s.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [values] },
  });
}

async function getRows(sheetName: string): Promise<string[][]> {
  const s = await getSheets();
  const res = await s.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A2:ZZ`,
  });
  return (res.data.values as string[][]) ?? [];
}

async function updateRow(sheetName: string, rowIndex: number, values: (string | number)[]) {
  const s = await getSheets();
  const row = rowIndex + 2; // 1행=헤더, 2행=첫 데이터
  await s.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A${row}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [values] },
  });
}

// ── 현장 ──────────────────────────────────────────────
export type SiteInput = {
  siteName: string; constructionPeriod: string; amount: string;
  progress: string; siteManager: string; safetyManager: string;
  siteEmail: string;
};

export async function createSite(input: SiteInput) {
  const siteId = `SITE-${Date.now()}`;
  const now = new Date().toISOString();
  await appendRow("현장목록", [
    siteId, input.siteName, input.constructionPeriod, input.amount,
    input.progress, input.siteManager, input.safetyManager,
    input.siteEmail, "운영중", now, now,
  ]);
  return { siteId };
}

export async function getSites() {
  const rows = await getRows("현장목록");
  return rows.map(r => ({
    siteId: r[0], siteName: r[1], constructionPeriod: r[2],
    amount: r[3], progress: r[4], siteManager: r[5],
    safetyManager: r[6], siteEmail: r[7], status: r[8],
  }));
}

// ── 점검 ──────────────────────────────────────────────
export type InspectionInput = {
  type: "정기안전보건평가" | "중처법준수평가";
  siteId: string; siteName: string;
  startDate: string; endDate: string;
  inspectors: string; mainWork: string;
};

export async function createInspection(input: InspectionInput) {
  const inspectionId = `INS-${Date.now()}`;
  const token = randomUUID();
  const now = new Date().toISOString();
  await appendRow("점검목록", [
    inspectionId, input.type, input.siteId, input.siteName,
    input.startDate, input.endDate, input.inspectors, input.mainWork,
    "", "", "", "", "", "진행중", token, "", "", "", now, now,
  ]);
  return { inspectionId, token };
}

export async function getInspections() {
  const rows = await getRows("점검목록");
  return rows.map(r => ({
    inspectionId: r[0], type: r[1], siteId: r[2], siteName: r[3],
    startDate: r[4], endDate: r[5], inspectors: r[6], mainWork: r[7],
    docScore: r[8], fieldScore: r[9], deduction: r[10],
    totalScore: r[12], status: r[13], token: r[14],
    pptUrl: r[15], xlsxUrl: r[16], createdAt: r[18],
  }));
}

export async function getInspectionByToken(token: string) {
  const rows = await getRows("점검목록");
  const idx = rows.findIndex(r => r[14] === token);
  if (idx === -1) return null;
  const r = rows[idx];
  return {
    rowIndex: idx, inspectionId: r[0], type: r[1], siteName: r[3],
    startDate: r[4], endDate: r[5], inspectors: r[6],
    status: r[13], token: r[14],
  };
}

export async function updateInspectionUrls(
  inspectionId: string,
  urls: { pptUrl?: string; xlsxUrl?: string }
) {
  const rows = await getRows("점검목록");
  const idx = rows.findIndex(r => r[0] === inspectionId);
  if (idx === -1) return;
  const r = rows[idx];
  if (urls.pptUrl) r[15] = urls.pptUrl;
  if (urls.xlsxUrl) r[16] = urls.xlsxUrl;
  r[19] = new Date().toISOString();
  await updateRow("점검목록", idx, r);
}

// ── 지적사항 ──────────────────────────────────────────
export type FindingInput = {
  inspectionId: string; seq: number;
  grade: "위험" | "미흡";
  riskType: string; original: string; summary: string;
  confirmed: string; deduction: number;
  actionRequest: string; photoUrl: string;
  charCount: number;
};

export async function createFinding(input: FindingInput) {
  const findingId = `FND-${Date.now()}-${input.seq}`;
  const now = new Date().toISOString();
  await appendRow("지적사항", [
    findingId, input.inspectionId, input.seq, input.grade,
    input.riskType, input.original, input.summary, input.confirmed,
    input.deduction, input.actionRequest, input.photoUrl,
    input.charCount <= 35 ? "OK" : "요약됨", input.charCount,
    "조치대기", now, now,
  ]);
  return { findingId };
}

export async function getFindingsByInspection(inspectionId: string) {
  const rows = await getRows("지적사항");
  return rows
    .filter(r => r[1] === inspectionId)
    .map(r => ({
      findingId: r[0], inspectionId: r[1], seq: Number(r[2]),
      grade: r[3] as "위험" | "미흡", riskType: r[4],
      original: r[5], summary: r[6], confirmed: r[7],
      deduction: Number(r[8]), actionRequest: r[9],
      photoUrl: r[10], status: r[13],
    }));
}

// ── 조치관리 ──────────────────────────────────────────
export async function createAction(input: {
  findingId: string; inspectionId: string;
  actionContent: string; photoUrl: string; submitter: string;
}) {
  const actionId = `ACT-${Date.now()}`;
  const now = new Date().toISOString();
  await appendRow("조치관리", [
    actionId, input.findingId, input.inspectionId,
    input.actionContent, input.photoUrl, input.submitter,
    now, "검토중", "", "", "",
  ]);
  return { actionId };
}
