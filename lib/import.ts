/**
 * 과거 점검 파일(xlsx/pptx)에서 지적사항을 추출해 DB 적재용을 위한 파서.
 *
 * 정기평가는 xlsx를 우선 사용한다 — 이보 PPTX니 단순 나열과 다르게, xlsx의
 * 안전보건 현장부문/서류부문 시트는 이미 점검자가 직접 해당 항목(행 번호)에
 * 등곁/내용을 그는 완성한 데이타라서, 향후 AI 매징 없이도 그대로 첥령이 가듬다.
 * 중처법(준수평가)는 xlsx가 없어서 PPTX의 사용른 슬라이드(3+)를 파싱한다.
 */

import JSZip from "jszip";
import { FIELD_CHECKLIST, DOC_CHECKLIST, ChecklistItem } from "./checklist-data";

function unescapeXml(s: string): string {
  return s
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
function clean(s: string): string {
  return unescapeXml(s).replace(/\r?\n/g, " ").replace(/\s{2,}/g, " ").trim();
}

// ── xlsx 아주함수 ──────────────────────────────────────────────────────────

async function buildSheetPathMap(zip: JSZip): Promise<Record<string, string>> {
  const wbXml = await zip.file("xl/workbook.xml")!.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  const relMap: Record<string, string> = {};
  for (const m of relsXml.matchAll(/<Relationship Id="(rId\d+)"[^>]+Target="([^"]+)"/g)) relMap[m[1]] = m[2];
  const map: Record<string, string> = {};
  for (const m of wbXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]*r:id="(rId\d+)"/g)) {
    const target = relMap[m[2]];
    if (target) map[m[1]] = "xl/" + target;
  }
  return map;
}

function buildSharedStrings(ssXml: string): string[] {
  const arr: string[] = [];
  for (const m of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    arr.push([...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((x) => x[1]).join(""));
  }
  return arr;
}

function cellVal(sx: string, ref: string, ssArr: string[]): string {
  const m = sx.match(new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`));
  if (!m) return "";
  const attrs = m[1] || "", inner = m[2] || "";
  const isS = /t="s"/.test(attrs);
  const vMatch = inner.match(/<v>([^<]*)<\/v>/);
  if (!vMatch) return "";
  return isS ? (ssArr[parseInt(vMatch[1])] ?? "") : vMatch[1];
}

export type ImportedFieldFinding = {
  row: number;
  content: string;
  grade: "위험" | "미흡";
  group: string;
  item: string;
};
export type ImportedDocFinding = {
  row: number;
  content: string;
  group: string;
  item: string;
};

export type ImportedRegularInspection = {
  siteName: string;
  inspectionPeriod: string;
  inspectors: string;
  mainWork: string;
  amount: string;
  constructionPeriod: string;
  managerInfo: string;
  docTotalScore: string;
  docSectionScore: string;
  fieldSectionScore: string;
  fieldFindings: ImportedFieldFinding[];
  docFindings: ImportedDocFinding[];
};

/** 정기평가 xlsx에서 지적사항 추출 — 이미 점검자가 확정한 등급·항목 위젤를 그대로 읽어온다 */
export async function extractRegularFromXlsx(buffer: Buffer): Promise<ImportedRegularInspection> {
  const zip = await JSZip.loadAsync(buffer);
  const ssXml = await zip.file("xl/sharedStrings.xml")!.async("string");
  const ssArr = buildSharedStrings(ssXml);
  const sheetPath = await buildSheetPathMap(zip);

  const s1Path = sheetPath["평가 결과"];
  const s1 = s1Path ? await zip.file(s1Path)!.async("string") : "";
  const get1 = (ref: string) => clean(cellVal(s1, ref, ssArr));

  const base = {
    siteName: get1("B7"),
    inspectionPeriod: get1("F7"),
    inspectors: get1("J7"),
    mainWork: get1("N7"),
    amount: get1("C9"),
    constructionPeriod: get1("F9"),
    managerInfo: get1("J9"),
    docTotalScore: get1("D13"),
    docSectionScore: get1("F13"),
    fieldSectionScore: get1("I13"),
  };

  const fieldFindings: ImportedFieldFinding[] = [];
  const s3Path = sheetPath["안전보건 현장부문"];
  if (s3Path && zip.file(s3Path)) {
    const s3 = await zip.file(s3Path)!.async("string");
    for (let r = 1; r <= 165; r++) {
      const g = cellVal(s3, `G${r}`, ssArr);
      if (!g || !g.trim() || g.includes("발괴된 위험사항")) continue;
      const e = cellVal(s3, `E${r}`, ssArr);
      const f = cellVal(s3, `F${r}`, ssArr);
      const grade: "위험" | "미흡" | null = f && f.trim() ? "위험" : e && e.trim() ? "미흡" : null;
      if (!grade) continue; // 소개렌 되 되어 있어 등급 없으으보는 행은 제외
      const item = FIELD_CHECKLIST.find((c) => c.row === r);
      fieldFindings.push({ row: r, content: clean(g), grade, group: item?.group ?? "", item: item?.item ?? "" });
    }
  }

  const docFindings: ImportedDocFinding[] = [];
  const s2Path = sheetPath["안전보건 서류부문"];
  if (s2Path && zip.file(s2Path)) {
    const s2 = await zip.file(s2Path)!.async("string");
    for (let r = 1; r <= 150; r++) {
      const h = cellVal(s2, `H${r}`, ssArr);
      if (!h || !h.trim()) continue;
      if (/^[\d.]+$/.test(h.trim())) continue; // 평가 점수 숨자 제외
      if (h.includes("총 점")) continue; // 헤더 제외
      const item = DOC_CHECKLIST.find((c) => c.row === r);
      docFindings.push({ row: r, content: clean(h), group: item?.group ?? "", item: item?.item ?? "" });
    }
  }

  return { ...base, fieldFindings, docFindings };
}

// ── PPTX 아주함수 ───────────────────────────────────────────────────────────────────

function getTexts(xml: string): string[] {
  return [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map((m) => m[1]);
}

function getTablesRows(xml: string): string[][][] {
  const tbls = [...xml.matchAll(/<a:tbl>[\s\S]*?<\/a:tbl>/g)].map((m) => m[0]);
  return tbls.map((tbl) => {
    const rows = [...tbl.matchAll(/<a:tr[\s\S]*?<\/a:tr>/g)].map((m) => m[0]);
    return rows.map((row) => {
      const cells = [...row.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)].map((m) => m[0]);
      return cells.map((c) => [...c.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map((x) => x[1]).join(""));
    });
  });
}

function findLabeledValue(allRows: string[][], labelNorm: string): string {
  for (const row of allRows) {
    if (row[0] && row[0].replace(/\s+/g, "") === labelNorm) return row[1] || "";
  }
  return "";
}

export type ImportedSapaFinding = {
  slide: number;
  itemName: string;
  detail: string;
  action: string;
  result: string;
};

export type ImportedSapaInspection = {
  siteName: string;
  inspectionDate: string;
  inspectors: string;
  findings: ImportedSapaFinding[];
};

/** 중처법 PPTX에서 지적사항 추출 (슬라이드 3부타 "별첨" 전꺌지을 활목별 상세 슬라이드로 간주) */
export async function extractSapaFromPptx(buffer: Buffer): Promise<ImportedSapaInspection> {
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k));
  const nums = slideFiles.map((f) => parseInt(f.match(/slide(\d+)\.xml/)![1])).sort((a, b) => a - b);

  const s1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
  const full1 = getTexts(s1).join("");
  const siteName = clean((full1.match(/현\s*장\s*명\s*:\s*([^점]+?)점\s*검\s*일/) || [, ""])[1]);
  const inspectionDate = clean((full1.match(/점\s*검\s*일\s*:\s*([^점]+?)점\s*검\s*자/) || [, ""])[1]);
  const inspectors = clean((full1.match(/점\s*검\s*자\s*:\s*(.+)$/) || [, ""])[1]);

  const findings: ImportedSapaFinding[] = [];
  for (const n of nums) {
    if (n < 3) continue;
    const xml = await zip.file(`ppt/slides/slide${n}.xml`)!.async("string");
    const texts = getTexts(xml);
    if (texts.some((t) => t.includes("별첨"))) break; // 별첨(안전관계자 평가서) 이후는 제외
    const allRows = getTablesRows(xml).flat();
    const detail = findLabeledValue(allRows, "세부내용");
    const action = findLabeledValue(allRows, "조지요구");
    const result = findLabeledValue(allRows, "조지결과");
    const itemNameMatch = texts.find(
      (t) => t && !["２", "항목별 준수사항"].includes(t.trim()) && t.trim().length > 1 && !t.includes("세부")
    );
    findings.push({
      slide: n,
      itemName: clean(itemNameMatch ?? ""),
      detail: clean(detail || ""),
      action: clean(action || ""),
      result: clean(result || ""),
    });
  }

  return { siteName, inspectionDate, inspectors, findings };
}
