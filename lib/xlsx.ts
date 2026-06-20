/**
 * 정기평가 xlsx 자동 생성 (JSZip 기반 XML 직접 조작)
 *
 * 시트 구조 (제주삼다수 템플릿 기준, 분석 완료):
 *   1. "평가 결과"        — B7=현장명 F7=점검기간 J7=점검자 N7=주요작업
 *                          C9=공사금액 F9=공사기간 J9=현장소장/안전관리자
 *                          D13=서류부문총점 F13=서류점수 I13=현장점수 L13=가감점 O13=보정계수
 *   2. "안전보건 서류부문" — 체크리스트 47항목(DOC_CHECKLIST). H{row}=의견(보조)
 *   3. "안전보건 현장부문" — 체크리스트 158항목(FIELD_CHECKLIST). D양호/E미흡/F위험=1or2, G{row}=발건사항
 *   4. "Sheet3"            — 가감점 보정표 (고정값, 미사용)
 *
 * 지적사항 텍스트는 PPTX 별첨 슬라이드와 100% 동일해야 함 — Gemini가 가장 유사한
 * 체크리스트 항목을 찾아 해당 행에 등곁/내용을 기록한다 (matchChecklistRow).
 */

import JSZip from "jszip";
import { FIELD_CHECKLIST, DOC_CHECKLIST } from "./checklist-data";
import { matchChecklistRow } from "./gemini";

export type FieldFindingInput = { content: string; grade: "위험" | "미흡" };
export type DocFindingInput = { content: string };

export type RegularXlsxInput = {
  siteName: string;
  inspectionPeriod: string;     // F7 (예: "2026-06-10 ~ 06-11")
  inspectors: string;           // J7
  mainWork: string;             // N7
  amount: string;               // C9 (숫자 또는 "47,316,083,290")
  constructionPeriod: string;   // F9
  managerInfo: string;          // J9 (현장소장/안전관리자 이맄)
  docTotalScore: string;        // D13 (서류부문 총점, 100점 만점)
  docSectionScore: string;      // F13 (서류부문 점수, 50점 만점)
  fieldSectionScore: string;    // I13 (현장부문 점수, 50점 만점)
  deduction: string;            // L13 (가감점)
  correctionFactor: string;     // O13 (보정계수)
  fieldFindings: FieldFindingInput[]; // 안전보건 현장부문 지적사항 (PPTX 슬라이드4와 동일 텍스트)
  docFindings: DocFindingInput[];     // 안전보건 서류부문 지적사항 (PPTX 슬라이드3과 동일 텍스트)
};

// ── 셀 참조 유틸 ─────────────────────────────────────────────────────────────────────

function colToNum(col: string): number {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function parseRef(ref: string): { col: string; row: number } {
  const m = ref.match(/^([A-Z]+)(\d+)$/)!;
  return { col: m[1], row: parseInt(m[2], 10) };
}

function escXml(s: string): string {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ── sharedStrings 조작 ────────────────────────────────────────────────────

function addSharedString(sharedXml: string, value: string): { xml: string; idx: number } {
  const count = [...sharedXml.matchAll(/<si>/g)].length;
  const newSi = `<si><t xml:space="preserve">${escXml(value)}</t></si>`;
  let xml = sharedXml.replace("</sst>", `${newSi}</sst>`);
  xml = xml.replace(/(<sst[^>]*\bcount=")(\d+)(")/, (_, p, n, s) => `${p}${parseInt(n) + 1}${s}`);
  xml = xml.replace(/(<sst[^>]*\buniqueCount=")(\d+)(")/, (_, p, n, s) => `${p}${parseInt(n) + 1}${s}`);
  return { xml, idx: count };
}

// ── 행 안에서 셀 삽입 위아(열 순서) 찾기 ──────────────────────────

function insertCellInOrder(rowContent: string, ref: string, newCellXml: string): string {
  const target = colToNum(parseRef(ref).col);
  const cellRe = /<c r="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
  let m: RegExpExecArray | null;
  while ((m = cellRe.exec(rowContent)) !== null) {
    if (colToNum(m[1]) > target) {
      return rowContent.slice(0, m.index) + newCellXml + rowContent.slice(m.index);
    }
  }
  return rowContent + newCellXml;
}

function getRowBlock(sheetXml: string, rowNum: number): { full: string; open: string; content: string; close: string } | null {
  const re = new RegExp(`(<row[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
  const m = sheetXml.match(re);
  if (!m) return null;
  return { full: m[0], open: m[1], content: m[2], close: m[3] };
}

/** 셀에 공유문자열(텍스트) 기록. 기존 셀 있으면 style 보존하며 교체, 없으면 행 안에 순서대로 삽입 */
function setCellSharedString(sheetXml: string, ref: string, ssIdx: number): string {
  const { row } = parseRef(ref);
  const rb = getRowBlock(sheetXml, row);
  if (!rb) return sheetXml; // 해당 행이 템플릿에 없으으면 스킵

  const cellRe = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`);
  let newContent: string;
  if (cellRe.test(rb.content)) {
    newContent = rb.content.replace(cellRe, (_match, attrs: string) => {
      const styleM = attrs.match(/\ss="(\d+)"/);
      const styleAttr = styleM ? ` s="${styleM[1]}"` : "";
      return `<c r="${ref}"${styleAttr} t="s"><v>${ssIdx}</v></c>`;
    });
  } else {
    newContent = insertCellInOrder(rb.content, ref, `<c r="${ref}" t="s"><v>${ssIdx}</v></c>`);
  }
  return sheetXml.replace(rb.full, `${rb.open}${newContent}${rb.close}`);
}

/** 셀에 숫자값 기록 */
function setCellNumber(sheetXml: string, ref: string, value: number): string {
  const { row } = parseRef(ref);
  const rb = getRowBlock(sheetXml, row);
  if (!rb) return sheetXml;

  const cellRe = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`);
  let newContent: string;
  if (cellRe.test(rb.content)) {
    newContent = rb.content.replace(cellRe, (_match, attrs: string) => {
      const styleM = attrs.match(/\ss="(\d+)"/);
      const styleAttr = styleM ? ` s="${styleM[1]}"` : "";
      return `<c r="${ref}"${styleAttr}><v>${value}</v></c>`;
    });
  } else {
    newContent = insertCellInOrder(rb.content, ref, `<c r="${ref}"><v>${value}</v></c>`);
  }
  return sheetXml.replace(rb.full, `${rb.open}${newContent}${rb.close}`);
}

// ── 시트 이렬 → 파일 경로 매핵 ──────────────────────────────────

async function buildSheetPathMap(zip: JSZip): Promise<Record<string, string>> {
  const wbXml = await zip.file("xl/workbook.xml")!.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");

  const relMap: Record<string, string> = {};
  for (const m of relsXml.matchAll(/<Relationship Id="(rId\d+)"[^>]+Target="([^"]+)"/g)) {
    relMap[m[1]] = m[2];
  }

  const map: Record<string, string> = {};
  for (const m of wbXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]*r:id="(rId\d+)"/g)) {
    const target = relMap[m[2]];
    if (target) map[m[1]] = "xl/" + target;
  }
  return map;
}

// ── PUBLIC: 정기평가 xlsx 생성 ─────────────────────────────────────

export async function generateRegularXlsx(
  templateBuffer: Buffer,
  data: RegularXlsxInput
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetPath = await buildSheetPathMap(zip);

  let ssXml = await zip.file("xl/sharedStrings.xml")!.async("string");
  function writeText(sheetXml: string, ref: string, text: string): string {
    const r = addSharedString(ssXml, text);
    ssXml = r.xml;
    return setCellSharedString(sheetXml, ref, r.idx);
  }

  // ── 1. 평가 결과 시트 ──
  const s1Path = sheetPath["평가 결과"];
  if (s1Path && zip.file(s1Path)) {
    let s1 = await zip.file(s1Path)!.async("string");
    s1 = writeText(s1, "B7", data.siteName);
    s1 = writeText(s1, "F7", data.inspectionPeriod);
    s1 = writeText(s1, "J7", data.inspectors);
    s1 = writeText(s1, "N7", data.mainWork);
    const amountNum = parseFloat(String(data.amount).replace(/[^0-9.]/g, "")) || 0;
    s1 = setCellNumber(s1, "C9", amountNum);
    s1 = writeText(s1, "F9", data.constructionPeriod);
    s1 = writeText(s1, "J9", data.managerInfo);
    s1 = setCellNumber(s1, "D13", parseFloat(data.docTotalScore) || 0);
    s1 = setCellNumber(s1, "F13", parseFloat(data.docSectionScore) || 0);
    s1 = setCellNumber(s1, "I13", parseFloat(data.fieldSectionScore) || 0);
    s1 = setCellNumber(s1, "L13", parseFloat(data.deduction) || 0);
    s1 = setCellNumber(s1, "O13", parseFloat(data.correctionFactor) || 1);
    zip.file(s1Path, s1);
  }

  // ── 2. 안전보건 현장부문 시트 — AI 매징 후 등걡/내용 기록 ──
  const s3Path = sheetPath["안전보건 현장부문"];
  if (s3Path && zip.file(s3Path)) {
    let s3 = await zip.file(s3Path)!.async("string");
    for (const f of data.fieldFindings) {
      const row = await matchChecklistRow(f.content, FIELD_CHECKLIST);
      if (!row) continue;
      const col = f.grade === "위험" ? "F" : "E";
      const val = f.grade === "위험" ? 2 : 1;
      s3 = setCellNumber(s3, `${col}${row}`, val);
      s3 = writeText(s3, `G${row}`, f.content);
    }
    zip.file(s3Path, s3);
  }

  // ── 3. 안전보건 서류부문 시트 — AI 매징 후 의견란 기록 ──
  const s2Path = sheetPath["안전보건 서류부문"];
  if (s2Path && zip.file(s2Path)) {
    let s2 = await zip.file(s2Path)!.async("string");
    for (const f of data.docFindings) {
      const row = await matchChecklistRow(f.content, DOC_CHECKLIST);
      if (!row) continue;
      s2 = writeText(s2, `H${row}`, f.content);
    }
    zip.file(s2Path, s2);
  }

  zip.file("xl/sharedStrings.xml", ssXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
