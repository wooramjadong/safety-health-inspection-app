/**
 * Google Sheets CRUD 라이닌러리
 *
 * 시트 구성 (ID: process.env.GOOGLE_SHEET_ID):
 *   1. 현장목롘  - A: id, B: 현장명, C: 주소, D: 현장소장, E: 안전관리자, F: 연락전화
 *   2. 점검목롘  - A: id, B: 현장명, C: 유형(정기/준수), D: 점검시작, E: 점검종료, F: 점검자,
 *                    G: 공사기간, H: 공사금액, I: 공정율, J: 현장소장, K: 안전관리자,
 *                    L: 서류점수, M: 현장점수, N: 상태, O: 업닠도토큰, P: pptxUrl, Q: xlsxUrl,
 *                    R: 월공정률보정, S: 주위험공종진행, T: 안전보조원운영 (보정계수 O13 하위 3건, xlsx P15/P16/P17에 쓰임)
 *                    U: 토목공사상세, V: 철쿄사상세, W: 습식공사상세 (주요작업 — xlsx N7 + pptx 슬라이드2 표1에 동일하게 반영)
 *   3. 지적사항  - A: id, B: 점검id, C: 섬션, D: 등급(현장부문만 사용, 서류부문은 보이), E: 내용, F: 조치요구, G: 항목명
 *   4. 조지관리  - A: id, B: 점검id, C: 지적id, D: 사진URL, E: 레램, F: 업닠시간
 *   5. 사용자   - (Phase2)
 *   6. 템플릿설정 - (Phase2)
 */

import { getAuth } from "./google";
import { google } from "googleapis";
import { randomUUID } from "crypto";

const SHEET_ID = process.env.GOOGLE_SHEET_ID!;

async function sheets() {
  const auth = await getAuth();
  return google.sheets({ version: "v4", auth });
}

// ─ 타입 ─────────────────────────────────────────────────────────────────────

export type SiteInput = {
  siteName: string;
  address?: string;
  siteManager?: string;
  safetyManager?: string;
  phone?: string;
};

export type InspectionInput = {
  siteName: string;
  type: "정기평가" | "준수평가";
  inspectionStart: string;
  inspectionEnd?: string;
  inspectors?: string;
  constructionPeriod?: string;
  amount?: string;
  progress?: string;
  siteManager?: string;
  safetyManager?: string;
  docScore?: string;
  fieldScore?: string;
  status?: string;
  /** 보정계수(O13=SUM(P15:P17)/3) 하위 3건 — xlsx 생성 시 P15/P16/P17에 그대로 쓰임 */
  monthlyProgressFactor?: string;  // 월 공정률 보정: 예) "1", "1.02", "1.04", "1.06"
  riskWorkFactor?: string;         // 주 위험공종 진행: 예) "1", "1.01", "1.02", "1.03"
  helperOperationFactor?: string;  // 안전보조원 운영: 예) "1", "1.02"
  /** 주요작업 상세 3건 — xlsx N7과 pptx 슬라이드2 표1[주요작업]에 동일하게 반영됨 */
  civilWorkDetail?: string;
  concreteWorkDetail?: string;
  wetWorkDetail?: string;
};

export type FindingInput = {
  inspectionId: string;
  section: "서류" | "현장";
  /**
   * 현장부문 지적사항만 사용 — 별첨 PPTX 구분을 및 xlsx 현장부문 시트의 양호/미흡/위험 구분에 매해됨.
   * 서류부문 지적사항은 등급 이 필단이 없으니(xlsx 서류부문 시트는 의견든에만 기롱), undefined로 둔다.
   */
  grade?: "위험" | "미흡";
  content: string;
  actionRequest?: string;
  itemName?: string;
};

export type ActionInput = {
  inspectionId: string;
  findingId: string;
  photoUrl: string;
  comment?: string;
  uploadedAt: string;
};

// ─ 현장 ─────────────────────────────────────────────────────────────────────

export async function createSite(input: SiteInput): Promise<string> {
  const api = await sheets();
  const id = randomUUID();
  await api.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "현장목롘!A:F",
    valueInputOption: "RAW",
    requestBody: { values: [[id, input.siteName, input.address ?? "", input.siteManager ?? "", input.safetyManager ?? "", input.phone ?? ""]] },
  });
  return id;
}

export async function getSites() {
  const api = await sheets();
  const res = await api.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "현장목롘!A2:F" });
  return (res.data.values ?? []).map(r => ({
    id: r[0], siteName: r[1], address: r[2], siteManager: r[3], safetyManager: r[4], phone: r[5],
  }));
}

// ─ 점검 ─────────────────────────────────────────────────────────────────────

export async function createInspection(input: InspectionInput): Promise<string> {
  const api = await sheets();
  const id = randomUUID();
  await api.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "점검목롘!A:W",
    valueInputOption: "RAW",
    requestBody: {
      values: [[
        id,
        input.siteName,
        input.type,
        input.inspectionStart,
        input.inspectionEnd ?? "",
        input.inspectors ?? "",
        input.constructionPeriod ?? "",
        input.amount ?? "",
        input.progress ?? "",
        input.siteManager ?? "",
        input.safetyManager ?? "",
        input.docScore ?? "",
        input.fieldScore ?? "",
        input.status ?? "점검완료",
        "", // 토큰
        "", // pptxUrl
        "", // xlsxUrl
        input.monthlyProgressFactor ?? "1",
        input.riskWorkFactor ?? "1",
        input.helperOperationFactor ?? "1",
        input.civilWorkDetail ?? "",
        input.concreteWorkDetail ?? "",
        input.wetWorkDetail ?? "",
      ]],
    },
  });
  return id;
}

export async function getInspections() {
  const api = await sheets();
  const res = await api.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "점검목롘!A2:W" });
  return (res.data.values ?? []).map(r => ({
    id: r[0], siteName: r[1], type: r[2],
    inspectionStart: r[3], inspectionEnd: r[4], inspectors: r[5],
    constructionPeriod: r[6], amount: r[7], progress: r[8],
    siteManager: r[9], safetyManager: r[10],
    docScore: r[11], fieldScore: r[12],
    status: r[13] ?? "",
    token: r[14] ?? "",
    pptxUrl: r[15] ?? "",
    xlsxUrl: r[16] ?? "",
    monthlyProgressFactor: r[17] ?? "1",
    riskWorkFactor: r[18] ?? "1",
    helperOperationFactor: r[19] ?? "1",
    civilWorkDetail: r[20] ?? "",
    concreteWorkDetail: r[21] ?? "",
    wetWorkDetail: r[22] ?? "",
    actionLink: r[14] ? `${process.env.NEXT_PUBLIC_BASE_URL}/action/${r[14]}` : "",
  }));
}

export async function getInspectionByToken(token: string) {
  const all = await getInspections();
  return all.find((i) => i.token === token) ?? null;
}

/**
 * 점검 토큰·URL 업단이트
 * @param status 값이 있을 뛌만 업닠이트 ("조지완료" 등)
 */
export async function updateInspectionUrls(
  inspectionId: string,
  token: string,
  pptxUrl: string,
  xlsxUrl: string,
  status?: string
): Promise<void> {
  const api = await sheets();
  const res = await api.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "점검목롘!A:A" });
  const rows = res.data.values ?? [];
  const rowIdx = rows.findIndex((r) => r[0] === inspectionId);
  if (rowIdx === -1) return;
  const sheetRow = rowIdx + 1; // 1-based

  // N(status)=14, O(token)=15, P(pptxUrl)=16, Q(xlsxUrl)=17
  const updates: Array<{ range: string; values: string[][] }> = [
    { range: `점검목롘!O${sheetRow}`, values: [[token]] },
    { range: `점검목롘!P${sheetRow}`, values: [[pptxUrl]] },
    { range: `점검목롘!Q${sheetRow}`, values: [[xlsxUrl]] },
  ];
  if (status) updates.push({ range: `점검목롘!N${sheetRow}`, values: [[status]] });

  await api.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      valueInputOption: "RAW",
      data: updates,
    },
  });
}

// ─ 지적사항 ───────────────────────────────────────────────────────────────

export async function createFinding(input: FindingInput): Promise<string> {
  const api = await sheets();
  const id = randomUUID();
  await api.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "지적사항!A:G",
    valueInputOption: "RAW",
    requestBody: {
      values: [[
        id, input.inspectionId, input.section, input.grade ?? "",
        input.content, input.actionRequest ?? "", input.itemName ?? "",
      ]],
    },
  });
  return id;
}

export async function getFindingsByInspection(inspectionId: string) {
  const api = await sheets();
  const res = await api.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "지적사항!A2:G" });
  return (res.data.values ?? [])
    .filter((r) => r[1] === inspectionId)
    .map((r) => ({
      id: r[0], inspectionId: r[1], section: r[2], grade: r[3] || undefined,
      content: r[4] ?? "", actionRequest: r[5] ?? "", itemName: r[6] ?? "",
    }));
}

// ─ 조지 ─────────────────────────────────────────────────────────────────────

export async function createAction(input: ActionInput): Promise<string> {
  const api = await sheets();
  const id = randomUUID();
  await api.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "조지관리!A:F",
    valueInputOption: "RAW",
    requestBody: {
      values: [[
        id, input.inspectionId, input.findingId,
        input.photoUrl, input.comment ?? "", input.uploadedAt,
      ]],
    },
  });
  return id;
}
