/**
 * 정기평가 xlsx 자동 생성 (JSZip 기반 XML 직접 조작)
 *
 * 시트 구조 (제주삼다수 템플릿 기준, 분석 완료):
 *   1. "평가 결과"        — B7=현장명 F7=점검기간 J7=점검자 N7=주요작업
 *                          C9=공사금액 F9=공사기간 J9=현장소장/안전관리자
 *                          D13/F13/I13/L13/O13은 절대 직접 쓰지 않아야 함! 전탠 수식 셀임
 *                          (D13=SUM(F13:K13)*O13+L13, F13=SUM(F15:H17),
 *                           I13='안전보건 현장부문'!G166, L13=G26, O13=SUM(P15:P17)/3)
 *                          → 하위 체크리스트 접도가 완료되면 자동 장사될.
 *   2. "안전보건 서류부문" — 항목당 3개 등급티어(E{row},E{row+1},E{row+2}=0/중/만점)
 *                          G{row}=해당 항목 점수(만점 때 의견없잌), H{row}=의견(보조)
 *                          일썱: 만점 ≡ 지적없음, 만점보다 적으뭴 의견란 필수
 *   3. "안전보건 현장부문" — 체크리스트 158항목(FIELD_CHECKLIST). E미흡=1/F위험=2, G{row}=발건사항
 *   4. "Sheet3"            — 가감점 보정표 (고정값, 미사용)
 *
 * 지적사항 텍스트는 PPTX 별첨 슬라이드와 100% 동일해야 함 — Gemini가 가장 유사한
 * 체크리스트 항목을 찾아 해당 행에 등곁/내용을 기록한다 (matchChecklistRow).
 *
 * ⚠️ 원눉 도해 템플릿(제주삼다수/괴주 닱)은 이미 수행된 실제 점검 결과르 그대로 지니고 있으목,
 * 새 점검을 생성할 띘 그 잔익도르가 설계 그대로 따끎거 면 면되지 않으므롌이띴,
 * 체크리스트 전 항목을 먼저 "잡만·지적없음" 기본값으로 리셋한 다음, 매징된 항목만 감점/내용을 기록한다.
 */

import JSZip from "jszip";
import { FIELD_CHECKLIST, DOC_CHECKLIST } from "./checklist-data";
import { matchChecklistRow } from "./gemini";

export type FieldFindingInput = { content: string; grade: "위험" | "미흡"; matchedRow?: number };
export type DocFindingInput = { content: string; matchedRow?: number };

export type RegularXlsxInput = {
  siteName: string;
  inspectionPeriod: string;     // F7 (예: "2026-06-10 ~ 06-11")
  inspectors: string;           // J7
  mainWork: string;             // N7
  amount: string;               // C9 (숫자 또는 "47,316,083,290")
  constructionPeriod: string;   // F9
  managerInfo: string;          // J9 (현장소장/안전관리자 이맄)
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

function buildSharedStrings(ssXml: string): string[] {
  const arr: string[] = [];
  for (const m of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    arr.push([...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((x) => x[1]).join(""));
  }
  return arr;
}

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

function getCellRaw(sheetXml: string, ref: string): { attrs: string; inner: string } | null {
  const { row } = parseRef(ref);
  const rb = getRowBlock(sheetXml, row);
  if (!rb) return null;
  const m = rb.content.match(new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`));
  return m ? { attrs: m[1] || "", inner: m[2] || "" } : null;
}

/** 셀의 조력된 값 읽기 (수식 셀이어뛌 관계없이 캐시/원시값 기준). 숫자이고 유한한 경우만 반환, 아니면 null */
function readNumericCellValue(sheetXml: string, ref: string): number | null {
  const c = getCellRaw(sheetXml, ref);
  if (!c) return null;
  const vM = c.inner.match(/<v>([^<]*)<\/v>/);
  let raw: string | null = vM ? vM[1] : null;
  if (raw === null) {
    // 수식만 있고 캐시값이 없는 경우, 수식 마지막 숫자 추출 (폴벱)
    const fM = c.inner.match(/<f>([\s\S]*?)<\/f>/);
    if (fM) {
      const numM = fM[1].match(/(\d+(\.\d+)?)\)/);
      if (numM) raw = numM[1];
    }
  }
  if (raw === null) return null;
  const n = parseFloat(raw);
  return Number.isFinite(n) ? n : null;
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

/** 셀에 숫자값 기록 (수식이 없는 일반 입력 셀에만 사용해야 함!) */
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

/** 셀 내용을 완전히 모음 (style은 보존). 수식이 엄는 입력 셀 리셋용 */
function clearCell(sheetXml: string, ref: string): string {
  const { row } = parseRef(ref);
  const rb = getRowBlock(sheetXml, row);
  if (!rb) return sheetXml;
  const cellRe = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`);
  if (!cellRe.test(rb.content)) return sheetXml;
  const newContent = rb.content.replace(cellRe, (_match, attrs: string) => {
    const styleM = attrs.match(/\ss="(\d+)"/);
    const styleAttr = styleM ? ` s="${styleM[1]}"` : "";
    return `<c r="${ref}"${styleAttr}/>`;
  });
  return sheetXml.replace(rb.full, `${rb.open}${newContent}${rb.close}`);
}

/** 서류부문 항목의 만점/중간값 솨타결과: 해당 항목은 row,row+1,row+2 세 줄에 0/중/만점 3개 등급티어가 있다 */
function getDocItemTiers(sheetXml: string, row: number): { max: number; mid: number } {
  const e0 = readNumericCellValue(sheetXml, `E${row}`);
  const e1 = readNumericCellValue(sheetXml, `E${row + 1}`);
  const e2 = readNumericCellValue(sheetXml, `E${row + 2}`);
  const tiers = [e0, e1, e2].filter((v): v is number => typeof v === "number" && Number.isFinite(v));
  if (tiers.length === 0) return { max: 10, mid: 5 }; // 파악 불가(N/A 항목 등) 시 안전한 기본값
  const sorted = [...tiers].sort((a, b) => a - b);
  return { max: sorted[sorted.length - 1], mid: sorted[Math.floor(sorted.length / 2)] };
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

/** Excel/LibreOffice가 파일을 열 띘 전원 재갠산하등띴 설정 — 안하띘휄 수정된 체크리스트 값이 반영되지 않을 수 있음 */
function forceFullRecalcOnLoad(workbookXml: string): string {
  if (!workbookXml.includes("<calcPr")) {
    return workbookXml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
  }
  return workbookXml.replace(/<calcPr([^/]*)\/>/, (m, attrs) => {
    if (attrs.includes("fullCalcOnLoad")) {
      return m.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"');
    }
    return `<calcPr${attrs} fullCalcOnLoad="1"/>`;
  });
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

  // ── 1. 평가 결과 시트 — 일반 입력값만 쓴다. D13/F13/I13/L13/O13(수식)은 절대 건드리지 않아 ──
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
    zip.file(s1Path, s1);
  }

  // ── 2. 안전보건 현장부문 시트 — 전원 리셋('지적없음' 기본값) 후 매징된 항목만 등곁/내용 기록 ──
  const s3Path = sheetPath["안전보건 현장부문"];
  if (s3Path && zip.file(s3Path)) {
    let s3 = await zip.file(s3Path)!.async("string");

    // 체크리스트 전항목 리셋: 이봇 템플릿에 다리 점검 결과가 남아있을 수 있으목, 새 점검에는 반영되지 않아야 함
    for (const item of FIELD_CHECKLIST) {
      s3 = clearCell(s3, `D${item.row}`);
      s3 = clearCell(s3, `E${item.row}`);
      s3 = clearCell(s3, `F${item.row}`);
      s3 = clearCell(s3, `G${item.row}`);
    }

    for (const f of data.fieldFindings) {
      const row = f.matchedRow ?? (await matchChecklistRow(f.content, FIELD_CHECKLIST));
      if (!row) continue;
      const col = f.grade === "위험" ? "F" : "E";
      const val = f.grade === "위험" ? 2 : 1;
      s3 = setCellNumber(s3, `${col}${row}`, val);
      s3 = writeText(s3, `G${row}`, f.content);
    }
    zip.file(s3Path, s3);
  }

  // ── 3. 안전보건 서류부문 시트 — 전항목 만점 리셋 후 매징된 항목만 감점+의견란 기록 ──
  const s2Path = sheetPath["안전보건 서류부문"];
  if (s2Path && zip.file(s2Path)) {
    let s2 = await zip.file(s2Path)!.async("string");

    // 체크리스트 전항목 리셋: 만점(=지적없음) 상태로 초기화
    for (const item of DOC_CHECKLIST) {
      const { max } = getDocItemTiers(s2, item.row);
      s2 = setCellNumber(s2, `G${item.row}`, max);
      s2 = clearCell(s2, `H${item.row}`);
    }

    for (const f of data.docFindings) {
      const row = f.matchedRow ?? (await matchChecklistRow(f.content, DOC_CHECKLIST));
      if (!row) continue;
      const { mid } = getDocItemTiers(s2, row);
      s2 = setCellNumber(s2, `G${row}`, mid); // 만점보다 낮은 점수로 감점 — 의견란이 있다는 말은 감점이 있다는 뛀
      s2 = writeText(s2, `H${row}`, f.content);
    }
    zip.file(s2Path, s2);
  }

  zip.file("xl/sharedStrings.xml", ssXml);

  // ── 4. 열 띘 전원 재갠산되등띴 설정 (수정된 체크리스트 값이 상단 수식에 즉시 반영되게) ──
  let wbXml = await zip.file("xl/workbook.xml")!.async("string");
  wbXml = forceFullRecalcOnLoad(wbXml);
  zip.file("xl/workbook.xml", wbXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
