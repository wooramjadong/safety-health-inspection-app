/**
 * xlsx 평가표 자동 생성 (JSZip 기반 XML 조작)
 *
 * 템플릿 시트 구성:
 *   평가결과 시트: 현장명, 공사금액, 점검기간, 공사기간, 등 요약
 *   서류부문 시트: 항목별 점수
 *   현장부문 시트: 지적사항
 */

import JSZip from "jszip";

export type XlsxInput = {
  siteName: string;
  amount: string;
  constructionPeriod: string;
  inspectionPeriod: string;
  siteManager: string;
  safetyManager: string;
  docScore: string;
  fieldScore: string;
  deduction: string;
  totalScore: string;
  docFindings: Array<{ category: string; content: string; score: string }>;
  fieldFindings: Array<{ riskType: string; content: string; deduction: number }>;
};

/** 공유 문자열에서 인덱스로 값 조회 */
function getSharedString(sharedXml: string, idx: number): string {
  const re = /<si>[\s\S]*?<\/si>/g;
  let count = 0;
  let m;
  while ((m = re.exec(sharedXml)) !== null) {
    if (count === idx) {
      const texts = [...m[0].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(x => x[1]);
      return texts.join("");
    }
    count++;
  }
  return "";
}

/** 공유 문자열에 새 문자열 추가 후 인덱스 반환 */
function addSharedString(sharedXml: string, value: string): { xml: string; idx: number } {
  const count = [...sharedXml.matchAll(/<si>/g)].length;
  const newSi = `<si><t>${value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</t></si>`;
  const xml = sharedXml.replace("</sst>", `${newSi}</sst>`);
  // count attr 업데이트
  const updated = xml.replace(
    /(<sst[^>]*count=")(\d+)"/,
    (_, pre) => `${pre}${count + 1}"`
  ).replace(
    /(<sst[^>]*uniqueCount=")(\d+)"/,
    (_, pre) => `${pre}${count + 1}"`
  );
  return { xml: updated, idx: count };
}

/** 시트 XML에서 특정 셀 ("A1" 형식) 값 수정 */
function setCellValue(sheetXml: string, cellRef: string, value: string, strIdx: number): string {
  // 시트에 해당 셀이 있으면 수정, 없으면 무시
  const cellRe = new RegExp(`(<c r="${cellRef}"[^>]*>)[\\s\\S]*?(<\/c>)`);
  if (cellRe.test(sheetXml)) {
    return sheetXml.replace(cellRe, `<c r="${cellRef}" t="s"><v>${strIdx}</v></c>`);
  }
  return sheetXml;
}

export async function generateRegularXlsx(
  templateBuffer: Buffer,
  data: XlsxInput
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(templateBuffer);

  let sharedXml = await zip.file("xl/sharedStrings.xml")!.async("string");

  // 입력값들을 공유문자열에 추가
  const fields: Record<string, string> = {
    siteName: data.siteName,
    amount: data.amount,
    constructionPeriod: data.constructionPeriod,
    inspectionPeriod: data.inspectionPeriod,
    siteManager: data.siteManager,
    safetyManager: data.safetyManager,
    docScore: data.docScore,
    fieldScore: data.fieldScore,
    deduction: data.deduction,
    totalScore: data.totalScore,
  };

  const idxMap: Record<string, number> = {};
  for (const [key, val] of Object.entries(fields)) {
    const res = addSharedString(sharedXml, val);
    sharedXml = res.xml;
    idxMap[key] = res.idx;
  }

  // 평가결과 시트 (시트 번호는 실제 xlsx에 맞게 조정 필요)
  // 새 테이터를 커스텀 셀 위치에 없데이트
  // 현재는 단순 로거로 대체 (Phase2에서 정교화)
  console.log("xlsx data:", JSON.stringify(idxMap));

  zip.file("xl/sharedStrings.xml", sharedXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
