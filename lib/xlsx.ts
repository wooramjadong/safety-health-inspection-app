/**
 * 정기평가 xlsx 자동 생성 (JSZip 기반 XML 직접 조작)
 *
 * 시트 구조 (제주삼다수 템플릿 기준):
 *   1. "평가 결과" — B7=현장명 F7=점검기간 J7=점검자 N7=주요작업, C9=공사금액 F9=공사기간 J9=담당자.
 *      D13/F13/I13/L13/O13은 절대 직접 쓰지 않아야 함(전탠 수식 셀). 하위 체크리스트가 채워지막 자동 장사.
 *   2. "안전보건 서류부문" — 항목당 3개 등급티어(E{row},E{row+1},E{row+2}=0/중/만점). G{row}=항목점수, H{row}=의견.
 *      G6/G50/G87=소계 합산, H6/H50/H87=백분율(평가결과!P21~23로 전도). 만점=지적없음, 감점=의견필수.
 *      H열은 감제 좌 좌우원의 도니(20.33), 3행 병합(65.7pt), 9pt. 다수 지적사항이 한 항목에 모여리면
 *      \n로 줄을 나눠 wrapText로 표시하고, 총자가 수가 츠첨(45자) 넘으도니 Gemini로 축쇕한다 (폰트 크기는 변경 안 함).
 *   3. "안전보건 현장부문" — 체크리스트 158항목. E미흡=1/F위험=2, G{row}=발건사항 (이롱 shrinkToFit 원본 스타일 유지).
 *   4. "Sheet3" — 가감점 보정표 (고정값, 미사용).
 *
 * ⚠️ 원본 템플릿은 이미 수행된 실제 점검 결과를 그대로 지니고 있어서, 새 점검 생성 시에는 체크리스트
 * 전항목을 만점/지적없음 기본값으로 적재리셋한 후, 매징된 항목만 감점/내용을 기록한다.
 *
 * 주의(3가지 핵시 수정):
 * 1) 소계/총점의 수식 셀(G6,H6,P21,F13,D13 등)은 캠시된 <v> 값이 업댌이트되지
 *    않을 수 있어, 열 래 잡에 반영안다일 수 있다 → 모든 수식 셀 캠시를 제거해 강제 장제(stripFormulaCache).
 * 2) 서류부문의 H6/H50/H87 원본 수식은 엑셀 전용 "SUM((A1,A2,...))" 유니울 문법이른데, 구귀시트
 *    파서가 읽지 모해 #ERROR가 난다 → SUM(A1,A2,...) 호환 문법으로 자동 변환(fixGoogleSheetsCompat).
 * 3) 다수 지적사항이 같은 체크리스트 항목에 매징되어도 쓰지 않고 합쳐서 쓴다(combineFindingTexts).
 *    서류부문 H열은 추가돔띴 wrapText 스타일을 부여해 줄바꿈이 동작하게 한다(ensureWrapTextStyle).
 */

import JSZip from "jszip";
import { FIELD_CHECKLIST, DOC_CHECKLIST } from "./checklist-data";
import { matchChecklistRow, combineFindingTexts, DOC_OPINION_TOTAL_CHARS } from "./gemini";

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

/** 셀의 캐시/원시 <v> 값을 숫자로 읽어온다. 없거나 N/A면 null */
function readNumericCellValue(sheetXml: string, ref: string): number | null {
  const c = getCellRaw(sheetXml, ref);
  if (!c) return null;
  const vM = c.inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return null;
  const n = parseFloat(vM[1]);
  return Number.isFinite(n) ? n : null;
}

/** 셀의 캐시된 <v> 값이 정확히 "N/A"인지 확인. 수식 소스(<f>) 자이에는 IF(...,"N/A",..)와 같이
 * "N/A" 문자열이 포함되어 있을 수 있으뫀롌이띴, 반드시 <v> 결과값만 확인해야 한다 */
function isNACell(sheetXml: string, ref: string): boolean {
  const c = getCellRaw(sheetXml, ref);
  if (!c) return false;
  const vM = c.inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return false;
  return /^N\/A$/i.test(vM[1].trim());
}

function getCellStyleIndex(sheetXml: string, ref: string): number {
  const c = getCellRaw(sheetXml, ref);
  if (!c) return 0;
  const m = c.attrs.match(/\ss="(\d+)"/);
  return m ? parseInt(m[1], 10) : 0;
}

function setCellStyleIndex(sheetXml: string, ref: string, newStyleIdx: number): string {
  const { row } = parseRef(ref);
  const rb = getRowBlock(sheetXml, row);
  if (!rb) return sheetXml;
  const cellRe = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`);
  if (!cellRe.test(rb.content)) return sheetXml;
  const newContent = rb.content.replace(cellRe, (_match, attrs: string, inner?: string) => {
    const restAttrs = attrs.replace(/\ss="\d+"/, "");
    const newAttrs = ` s="${newStyleIdx}"${restAttrs}`;
    return inner !== undefined ? `<c r="${ref}"${newAttrs}>${inner}</c>` : `<c r="${ref}"${newAttrs}/>`;
  });
  return sheetXml.replace(rb.full, `${rb.open}${newContent}${rb.close}`);
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
 * row+2(만점 참조셀)가 "N/A"면 이 항목은 H6/H50/H87 백분율 분부에서 제외되뱀이부디,
 * 점수를 0으로 둔뛌 이 구조롬해야 한다 (아니면 백분율이 100%를 초과함).
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

/** 해당 시트의 모든 수식 셀에서 캠시된 <v> 결과를 제거 (수식은 유지) — 열 래 강제 장제로 재강산되게 함 */
function stripFormulaCache(sheetXml: string): string {
  return sheetXml.replace(/(<c[^>]*>)(<f>[\s\S]*?<\/f>)<v>[^<]*<\/v>(<\/c>)/g, "$1$2$3");
}

/**
 * 엑셀 전용 SUM((A1,A2,...)) 유니울 문법을 구귀시트 호환 SUM(A1,A2,...)로 바꿜다.
 */
function fixGoogleSheetsCompat(sheetXml: string): string {
  return sheetXml.replace(/SUM\(\(((?:[A-Z]+\d+,?)+)\)\)/g, "SUM($1)");
}

/**
 * styles.xml의 cellXfs에서 baseIndex 스타일을 기반으로 wrapText=1이 추가된 새 스타일을 만들어 말한다.
 * 이보 폰트·타례롔·채우면 육지하며 alignment에 wrapText만 추가한다. 결과는 캐시해 중복 생성을 피한다.
 */
function ensureWrapTextStyle(
  stylesXml: string,
  baseIndex: number,
  cache: Map<number, number>
): { xml: string; index: number } {
  if (cache.has(baseIndex)) return { xml: stylesXml, index: cache.get(baseIndex)! };

  const cellXfsMatch = stylesXml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  if (!cellXfsMatch) return { xml: stylesXml, index: baseIndex };

  const count = parseInt(cellXfsMatch[1], 10);
  const xfsContent = cellXfsMatch[2];
  const xfEntries = xfsContent.match(/<xf\b[^>]*?\/>|<xf\b[^>]*?>[\s\S]*?<\/xf>/g) || [];
  if (baseIndex >= xfEntries.length) return { xml: stylesXml, index: baseIndex };

  const baseXf = xfEntries[baseIndex];
  if (/wrapText="1"/.test(baseXf)) {
    cache.set(baseIndex, baseIndex);
    return { xml: stylesXml, index: baseIndex };
  }

  let newXf: string;
  if (/<alignment\b[^>]*\/>/.test(baseXf)) {
    newXf = baseXf.replace(/<alignment\b([^>]*)\/>/, (_m, attrs) => `<alignment${attrs} wrapText="1"/>`);
  } else if (/<alignment\b[^>]*>/.test(baseXf)) {
    newXf = baseXf.replace(/<alignment\b([^>]*)>/, (_m, attrs) => `<alignment${attrs} wrapText="1">`);
  } else if (baseXf.endsWith("/>")) {
    newXf = baseXf.slice(0, -2) + ' applyAlignment="1"><alignment wrapText="1"/></xf>';
  } else {
    newXf = baseXf.replace(/<\/xf>$/, '<alignment wrapText="1"/></xf>');
  }

  const newIndex = xfEntries.length;
  const newXfsContent = xfsContent + newXf;
  const newCellXfs = `<cellXfs count="${count + 1}">${newXfsContent}</cellXfs>`;
  const newStylesXml = stylesXml.replace(cellXfsMatch[0], newCellXfs);
  cache.set(baseIndex, newIndex);
  return { xml: newStylesXml, index: newIndex };
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

  // ── 2. 안전보건 현장부문 시트 — 전원 리셋('지적없음' 기본값) 후, 같은 항목에 매징된 지적사항은 합쳐서 기록 ──
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

    const fieldGroups = new Map<number, { grades: string[]; texts: string[] }>();
    for (const f of data.fieldFindings) {
      const row = f.matchedRow ?? (await matchChecklistRow(f.content, FIELD_CHECKLIST));
      if (!row) continue;
      if (!fieldGroups.has(row)) fieldGroups.set(row, { grades: [], texts: [] });
      const g = fieldGroups.get(row)!;
      g.grades.push(f.grade);
      g.texts.push(f.content);
    }

    for (const [row, { grades, texts }] of fieldGroups) {
      const grade = grades.includes("위험") ? "위험" : "미흡";
      const col = grade === "위험" ? "F" : "E";
      const val = grade === "위험" ? 2 : 1;
      s3 = setCellNumber(s3, `${col}${row}`, val);
      const combined = await combineFindingTexts(texts, 60); // 현장부문 G열(돈 35.89, 바다)는 여유가 끴해서 대떨 좀더 너과롱다
      s3 = writeText(s3, `G${row}`, combined);
    }
  }

  // ── 3. 안전보건 서류부문 시트 — 전항목 만점 리셋 후, 같은 항목에 매징된 지적사항은 합쳐서 감점+의견란 기록 ──
  const s2Path = sheetPath["안전보건 서류부문"];
  let s2 = "";
  let stylesXml = await zip.file("xl/styles.xml")!.async("string");
  const wrapStyleCache = new Map<number, number>();

  if (s2Path && zip.file(s2Path)) {
    s2 = await zip.file(s2Path)!.async("string");

    for (const item of DOC_CHECKLIST) {
      const { max, notApplicable } = getDocItemTiers(s2, item.row);
      s2 = setCellNumber(s2, `G${item.row}`, notApplicable ? 0 : max);
      s2 = clearCell(s2, `H${item.row}`);
    }

    const docGroups = new Map<number, string[]>();
    for (const f of data.docFindings) {
      const row = f.matchedRow ?? (await matchChecklistRow(f.content, DOC_CHECKLIST));
      if (!row) continue;
      if (!docGroups.has(row)) docGroups.set(row, []);
      docGroups.get(row)!.push(f.content);
    }

    for (const [row, texts] of docGroups) {
      const { mid, notApplicable } = getDocItemTiers(s2, row);
      s2 = setCellNumber(s2, `G${row}`, notApplicable ? 0 : mid);

      const combined = await combineFindingTexts(texts, DOC_OPINION_TOTAL_CHARS);

      // 줄바꿈(wrapText)이 적용된 스타일인지 확인하고, 아니면 해당 셀의 기존 스타일을 원본으로 새 스타일 생성
      const curStyleIdx = getCellStyleIndex(s2, `H${row}`);
      const { xml: newStylesXml, index: wrappedIdx } = ensureWrapTextStyle(stylesXml, curStyleIdx, wrapStyleCache);
      stylesXml = newStylesXml;
      if (wrappedIdx !== curStyleIdx) {
        s2 = setCellStyleIndex(s2, `H${row}`, wrappedIdx);
      }

      s2 = writeText(s2, `H${row}`, combined);
    }
  }

  zip.file("xl/sharedStrings.xml", ssXml);
  zip.file("xl/styles.xml", stylesXml);

  // ── 4. 구귀시트 호환 문법 보정 + 수식 셀 캠시 제거(모든 시트) + 재강산 한구설정 ──
  for (const sheetName of Object.keys(sheetPath)) {
    const sp = sheetPath[sheetName];
    let xml = sp === s1Path ? s1 : sp === s3Path ? s3 : sp === s2Path ? s2 : await zip.file(sp)!.async("string");
    xml = fixGoogleSheetsCompat(xml);
    xml = stripFormulaCache(xml);
    zip.file(sp, xml);
  }

  let wbXml = await zip.file("xl/workbook.xml")!.async("string");
  wbXml = forceFullRecalcOnLoad(wbXml);
  zip.file("xl/workbook.xml", wbXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
