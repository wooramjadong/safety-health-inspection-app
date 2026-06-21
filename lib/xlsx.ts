/**
 * 정기평가 xlsx 자동 생성 (JSZip 기반 XML 직접 조작)
 *
 * 중요: generateRegularXlsx는 { buffer, scores } 를 반환한다. scores는 xlsx의
 * 수식(D13/F13/I13/O13)과 수학적으로 동일하게 JS로 계산된 값이고, PPTX 슬라이드의
 * 점수표와 일썱시키는 용도로 쓴다 (LibreOffice 재계산 검증 결과 소수점까지 일썱).
 *
 * 점수 공식 (원본 수식을 그대로 분석해 재현):
 *   현장부문점수 = 50 - 구 매징된 체크리스트 항목의 감점 합 (행 단위, 중복 매징도 1번만 차감) — 미흡=-1, 위험=-2
 *   서류부문점수 = 현장소장%/10 + 관리감독자%/5 + 안전관리자%/5
 *   보정계수 = (월공정률보정+주위험공종진행+안전보조원운영)/3
 *   총점 = (서류부문점수+현장부문점수)*보정계수+가감점(현재 UI 없으목 항상 0)
 *
 * ⚠️ 현장부문 지적사항(G열)은 PPTX 별첨 슬라이드와 텍스트가 글자 하나까지 동일해야 하므로,
 * 여러 지적사항이 같은 항목에 매징되어도 절대 추가로 축쇕하지 않고 줄바꿈만으로 합쇔다
 * (원본 G열 스타일이 shrinkToFit이니다도 그래도 넘지면 Excel이 자주 구도롱다. 폰트는 다조지 않으르, 원본 템플릿의
 * 기반 동작이다. 단 서류부문(H열)뚘 PPTX에 대응하는 항목이 없으몤롱 조그 계속 축쇕한다).
 *
 * 시트 구조 (제주삼다수 템플릿 기준):
 *   1. "평가 결과" — B7=현장명 F7=점검기간 J7=점검자 N7=주요작업, C9=공사금액 F9=공사기간 J9=담당자.
 *      P15/P16/P17=보정계수 하위 3건. G21:G25=가감점(현재 UI 없이 0으로 리셋).
 *      D13/F13/I13/L13/O13은 절대 직접 쓰지 않아야 함(수식 셀).
 *   2. "안전보건 서류부문" — 항목당 3개 등급티어. G{row}=항목점수, H{row}=의견. 만점=지적없음.
 *   3. "안전보건 현장부문" — 체크리스트 158항목. E미흡=1/F위험=2, G{row}=발건사항.
 *   4. "Sheet3" — 가감점 보정표 (고정값, 미사용).
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
  monthlyProgressFactor?: string;
  riskWorkFactor?: string;
  helperOperationFactor?: string;
  fieldFindings: FieldFindingInput[];
  docFindings: DocFindingInput[];
};

export type ComputedScores = {
  totalScore: number;
  docScore: number;
  fieldScore: number;
  deduction: number;
  correctionFactor: number;
};

export type RegularXlsxResult = {
  buffer: Buffer;
  scores: ComputedScores;
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

function readNumericCellValue(sheetXml: string, ref: string): number | null {
  const c = getCellRaw(sheetXml, ref);
  if (!c) return null;
  const vM = c.inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return null;
  const n = parseFloat(vM[1]);
  return Number.isFinite(n) ? n : null;
}

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

function stripFormulaCache(sheetXml: string): string {
  return sheetXml.replace(/(<c[^>]*>)(<f>[\s\S]*?<\/f>)<v>[^<]*<\/v>(<\/c>)/g, "$1$2$3");
}

function fixGoogleSheetsCompat(sheetXml: string): string {
  return sheetXml.replace(/SUM\(\(((?:[A-Z]+\d+,?)+)\)\)/g, "SUM($1)");
}

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
): Promise<RegularXlsxResult> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetPath = await buildSheetPathMap(zip);

  let ssXml = await zip.file("xl/sharedStrings.xml")!.async("string");
  function writeText(sheetXml: string, ref: string, text: string): string {
    const r = addSharedString(ssXml, text);
    ssXml = r.xml;
    return setCellSharedString(sheetXml, ref, r.idx);
  }

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
    s1 = setCellNumber(s1, "P15", parseFloat(data.monthlyProgressFactor ?? "1") || 1);
    s1 = setCellNumber(s1, "P16", parseFloat(data.riskWorkFactor ?? "1") || 1);
    s1 = setCellNumber(s1, "P17", parseFloat(data.helperOperationFactor ?? "1") || 1);
    for (const r of [21, 22, 23, 24, 25]) s1 = setCellNumber(s1, `G${r}`, 0);
  }

  const s3Path = sheetPath["안전보건 현장부문"];
  let s3 = "";
  const fieldGroups = new Map<number, { grades: string[]; texts: string[] }>();
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
      // 주의: PPTX 별첨 슬라이드와 텍스트를 글자 하나까지 일썱해야 하므로 추가 축쇕 절대 안 함 — 줄바꿈롱 합쇔
      const combined = texts.join("\n");
      s3 = writeText(s3, `G${row}`, combined);
    }
  }

  let fieldDeduction = 0;
  for (const { grades } of fieldGroups.values()) {
    fieldDeduction += grades.includes("위험") ? 2 : 1;
  }
  const fieldScore = Math.max(0, 50 - fieldDeduction);

  const s2Path = sheetPath["안전보건 서류부문"];
  let s2 = "";
  let stylesXml = await zip.file("xl/styles.xml")!.async("string");
  const wrapStyleCache = new Map<number, number>();
  const roleGroups: Record<string, { g: number; denom: number }> = {
    "현장소장": { g: 0, denom: 0 },
    "관리감독자": { g: 0, denom: 0 },
    "안전관리자": { g: 0, denom: 0 },
  };

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

      // 서류부문은 PPTX에 대응하는 항목이 없으몤롱 (xlsx 전용), 셀 용뇘 초과 시 Gemini 축쇕 계속 적용
      const combined = await combineFindingTexts(texts, DOC_OPINION_TOTAL_CHARS);

      const curStyleIdx = getCellStyleIndex(s2, `H${row}`);
      const { xml: newStylesXml, index: wrappedIdx } = ensureWrapTextStyle(stylesXml, curStyleIdx, wrapStyleCache);
      stylesXml = newStylesXml;
      if (wrappedIdx !== curStyleIdx) {
        s2 = setCellStyleIndex(s2, `H${row}`, wrappedIdx);
      }

      s2 = writeText(s2, `H${row}`, combined);
    }

    for (const item of DOC_CHECKLIST) {
      const roleKey = Object.keys(roleGroups).find((k) => item.group.includes(k));
      if (!roleKey) continue;
      const gVal = readNumericCellValue(s2, `G${item.row}`) ?? 0;
      const { notApplicable, max } = getDocItemTiers(s2, item.row);
      roleGroups[roleKey].g += gVal;
      if (!notApplicable) roleGroups[roleKey].denom += max;
    }
  }

  const pct = (k: string) => (roleGroups[k].denom > 0 ? (100 * roleGroups[k].g) / roleGroups[k].denom : 0);
  const docScore = pct("현장소장") / 10 + pct("관리감독자") / 5 + pct("안전관리자") / 5;

  const correctionFactor =
    ((parseFloat(data.monthlyProgressFactor ?? "1") || 1) +
      (parseFloat(data.riskWorkFactor ?? "1") || 1) +
      (parseFloat(data.helperOperationFactor ?? "1") || 1)) /
    3;
  const deduction = 0;
  const totalScore = (docScore + fieldScore) * correctionFactor + deduction;

  zip.file("xl/sharedStrings.xml", ssXml);
  zip.file("xl/styles.xml", stylesXml);

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

  const buffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return { buffer, scores: { totalScore, docScore, fieldScore, deduction, correctionFactor } };
}
