/**
 * 정기평가 xlsx 자동 생성 (JSZip 기반 XML 직접 조작)
 *
 * 시트 구조 (제주삼다수 템플릿 기준, 분석 완료):
 *   1. "평가 결과"        — B7=현장명 F7=점검기간 J7=점검자 N7=주요작업
 *                          C9=공사금액 F9=공사기간 J9=현장소장/안전관리자
 *                          D13/F13/I13/L13/O13은 절대 직접 쓰지 않아야 함! 전탠 수식 셀임.
 *                          D13=SUM(F13:K13)*O13+L13, F13=SUM(F15:H17) (F15=P21/10 등),
 *                          P21/P22/P23='안전보건 서류부문'!H6/H50/H87 를 참조,
 *                          I13='안전보건 현장부문'!G166, L13=G26, O13=SUM(P15:P17)/3
 *   2. "안전보건 서류부문" — 항목당 3개 등급티어(E{row},E{row+1},E{row+2}=0/중/만점)
 *                          G{row}=해당 항목 점수, H{row}=의견(보조)
 *                          G6/G50/G87=소개 합산, H6/H50/H87=백분율(이게 평가결과!P21~23로 전도될)
 *                          일보: 만점 ≡ 지적없음, 만점보다 적으본이 적의견란 필수
 *                          주의: 일부 항목은 row+2 셀이 "N/A"이어서 H6/H50/H87 분부에서 제외되면—
 *                          이런 항목은 점수 0륐 처리해야 함 (아니면 분자에만 점수가 추가되어 100% 초과)
 *   3. "안전보건 현장부문" — 체크리스트 158항목(FIELD_CHECKLIST). E미흡=1/F위험=2, G{row}=발건사항
 *   4. "Sheet3"            — 가감점 보정표 (고정값, 미사용)
 *
 * ⚠️ 원본 템플릿(제주삼다수 등)은 이미 수행된 실제 점검 결과를 그대로 지니고 있어서,
 * 새 점검 생성 시에는 체크리스트 전항목을 만점/지적없음 기본값으로 먼저 리셋한 후, 매징된
 * 항목만 감점/내용을 기록한다.
 *
 * 중요: 다수 수식의 총점이 설정될 따 계산될 있는 수식 셀(G6,H6,P21,F13,D13 등)의 캠시된 <v>
 * 값이 업데이트되지 않을 수 있어 엔맄 계산(LibreOffice/Excel)가 잡에지 악을 수다— 따라서
 * 마지막에 모든 시트의 쯄식 셀 캠시된 값을 제거해 열 띘 잡에 장제로 도롭다 (stripFormulaCache).
 */

import JSZip from "jszip";
import { FIELD_CHECKLIST, DOC_CHECKLIST } from "./checklist-data";
import { matchChecklistRow } from "./gemini";

export type FieldFindingInput = { content: string; grade: "위험" | "미흡"; matchedRow?: number };
export type DocFindingInput = { content: string; matchedRow?: number };

export type RegularXlsxInput = {
  siteName: string;
  inspectionPeriod: string;
  inspectors: string;
  mainWork: string;
  amount: string;
  constructionPeriod: string;
  managerInfo: string;
  fieldFindings: FieldFindingInput[];
  docFindings: DocFindingInput[];
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

function addSharedString(sharedXml: string, value: string): { xml: string; idx: number } {
  const count = [...sharedXml.matchAll(/<si>/g)].length;
  const newSi = `<si><t xml:space="preserve">${escXml(value)}</t></si>`;
  let xml = sharedXml.replace("</sst>", `${newSi}</sst>`);
  xml = xml.replace(/(<sst[^>]*\bcount=")(\d+)(")/, (_, p, n, s) => `${p}${parseInt(n) + 1}${s}`);
  xml = xml.replace(/(<sst[^>]*\buniqueCount=")(\d+)(")/, (_, p, n, s) => `${p}${parseInt(n) + 1}${s}`);
  return { xml, idx: count };
}

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

/** 셀의 캐시/원시 <v> 값을 수자로 읽음. 수식의 캠시값이 엄고 N/A릅 경우 null */
function readNumericCellValue(sheetXml: string, ref: string): number | null {
  const c = getCellRaw(sheetXml, ref);
  if (!c) return null;
  const vM = c.inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return null;
  const n = parseFloat(vM[1]);
  return Number.isFinite(n) ? n : null;
}

/** 해당 셀의 캐시된 <v> 결개개 정확히 "N/A" 문자열인지 확인 (수식 소스코돜 자이귶 IF(...,"N/A",..) 자체에도
 * "N/A" 문자열이 활장되어 있으목로, 절대 <f> 다이 아니라 <v> 결개개만 확인해야 함 */
function isNACell(sheetXml: string, ref: string): boolean {
  const c = getCellRaw(sheetXml, ref);
  if (!c) return false;
  const vM = c.inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return false;
  return /^N\/A$/i.test(vM[1].trim());
}

function setCellSharedString(sheetXml: string, ref: string, ssIdx: number): string {
  const { row } = parseRef(ref);
  const rb = getRowBlock(sheetXml, row);
  if (!rb) return sheetXml;

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

/**
 * 서류부문 항목의 등급티어 조회: row, row+1, row+2 세 줄에 0/중/만점 등급이 있다.
 * row+2(만점 참조셀)가 "N/A"닌 동일 이 항목은 H6/H50/H87 분부에서 제외되르니롌,
 * 점수를 0으로 둔 익에래 이 구자도롭해야 한다 (아니면 백분윤이 100%를 초과함).
 */
function getDocItemTiers(sheetXml: string, row: number): { max: number; mid: number; notApplicable: boolean } {
  if (isNACell(sheetXml, `E${row + 2}`)) return { max: 0, mid: 0, notApplicable: true };
  const e0 = readNumericCellValue(sheetXml, `E${row}`);
  const e1 = readNumericCellValue(sheetXml, `E${row + 1}`);
  const e2 = readNumericCellValue(sheetXml, `E${row + 2}`);
  const tiers = [e0, e1, e2].filter((v): v is number => typeof v === "number" && Number.isFinite(v));
  if (tiers.length === 0) return { max: 10, mid: 5, notApplicable: false };
  const sorted = [...tiers].sort((a, b) => a - b);
  return { max: sorted[sorted.length - 1], mid: sorted[Math.floor(sorted.length / 2)], notApplicable: false };
}

/** 해당 시트의 모돈 수식 셀에서 캐시된 <v> 결개를 제거 (수식은 유지) — 열 띘 잡에 강제 장제로 재계산되게 함 */
function stripFormulaCache(sheetXml: string): string {
  return sheetXml.replace(/(<c[^>]*>)(<f>[\s\S]*?<\/f>)<v>[^<]*<\/v>(<\/c>)/g, "$1$2$3");
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
  let s1 = "";
  if (s1Path && zip.file(s1Path)) {
    s1 = await zip.file(s1Path)!.async("string");
    s1 = writeText(s1, "B7", data.siteName);
    s1 = writeText(s1, "F7", data.inspectionPeriod);
    s1 = writeText(s1, "J7", data.inspectors);
    s1 = writeText(s1, "N7", data.mainWork);
    const amountNum = parseFloat(String(data.amount).replace(/[^0-9.]/g, "")) || 0;
    s1 = setCellNumber(s1, "C9", amountNum);
    s1 = writeText(s1, "F9", data.constructionPeriod);
    s1 = writeText(s1, "J9", data.managerInfo);
  }

  // ── 2. 안전보건 현장부문 시트 — 전원 리셋('지적없음' 기본값) 후 매징된 항목만 등곁/내용 기록 ──
  const s3Path = sheetPath["안전보건 현장부문"];
  let s3 = "";
  if (s3Path && zip.file(s3Path)) {
    s3 = await zip.file(s3Path)!.async("string");
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
  }

  // ── 3. 안전보건 서류부문 시트 — 전항목 만점 리셋 후 매징된 항목만 감점+의견란 기록 ──
  const s2Path = sheetPath["안전보건 서류부문"];
  let s2 = "";
  if (s2Path && zip.file(s2Path)) {
    s2 = await zip.file(s2Path)!.async("string");

    // 원반 리셋: N/A 항목은 0점, 다르고 항목은 만점과 함껸 의견란 초기화
    for (const item of DOC_CHECKLIST) {
      const { max, notApplicable } = getDocItemTiers(s2, item.row);
      s2 = setCellNumber(s2, `G${item.row}`, notApplicable ? 0 : max);
      s2 = clearCell(s2, `H${item.row}`);
    }

    for (const f of data.docFindings) {
      const row = f.matchedRow ?? (await matchChecklistRow(f.content, DOC_CHECKLIST));
      if (!row) continue;
      const { mid, notApplicable } = getDocItemTiers(s2, row);
      s2 = setCellNumber(s2, `G${row}`, notApplicable ? 0 : mid);
      s2 = writeText(s2, `H${row}`, f.content);
    }
  }

  zip.file("xl/sharedStrings.xml", ssXml);

  // ── 4. 쯄식 셀 캠시 제거(모든 시트) + 재계산 핬랈강 설정 — 소개/총점이 강제로 새롬감삼되게 함 ──
  for (const sheetName of Object.keys(sheetPath)) {
    const sp = sheetPath[sheetName];
    let xml = sp === s1Path ? s1 : sp === s3Path ? s3 : sp === s2Path ? s2 : await zip.file(sp)!.async("string");
    xml = stripFormulaCache(xml);
    zip.file(sp, xml);
  }

  let wbXml = await zip.file("xl/workbook.xml")!.async("string");
  wbXml = forceFullRecalcOnLoad(wbXml);
  zip.file("xl/workbook.xml", wbXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
