/**
 * PPTX 자동 생성 라이브러리
 * 접근 방식: JSZip으로 PPTX(ZIP)을 열고, XML을 문자열 조작
 *
 * 주의(표지/결과요약 슬라이드의 표 구조가 서로 다르다 — 실제 파일 추출해 확인된 핵시 버서 버그 수정):
 *   • slide1(표지): "점검 현장"/"점검 기간"/"점검 인원"은 한 행에 [라볘,값]이 동시에 있다
 *     (label LEFT, value RIGHT 같은 행) → replaceAdjacentCell 사용.
 *   • slide2(결과요약): "현장명"/"공사기간" 등은 한 행에 다수 라볘이 나열되고, 값은
 *     다음 행의 같은 열 위지에 있다 (label ABOVE, value BELOW) → replaceCellBelow 사용.
 *   • 이전에는 slide1을 replaceTextBoxContent(텍스트박스 전용)로 처리해서, "점검 현장" 등이
 *     실제는 표(table) 안에 있어 아르 함수가 아띔 찾지 봇해 조용하게 아더 작동했으이, 실제로는
 *     현장명이 전혀 봘도욨지 않고 원문(샘료 현장명)가 그대로 남아있었다. 이제 해결될.
 *
 * 별첨 슬라이드 구조 (정기평가 별첨):
 *   테이블0 [구분]: row1=구분, row2=위험|미흡
 *   테이블1 [체크박스]: row1=헤더, row2=■|□ ×7
 *   테이블2 [내용]: row2=사진영역, row3=['내용',지적텍스트,'내용',조치텍스트]
 */

import JSZip from "jszip";

const RISK_COLS: Record<string, number> = {
  "추락": 0, "낙하": 1, "충돌/협착": 2,
  "화재": 3, "질식": 4, "감전": 5, "기타": 6,
};

// ── XML 보조 함수 ─────────────────────────────────────────────

function escXml(s: string): string {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function findNthTable(xml: string, n: number): { start: number; end: number } | null {
  const OPEN = "<a:tbl>";
  const CLOSE = "</a:tbl>";
  let count = 0;
  let pos = 0;
  while (pos < xml.length) {
    const s = xml.indexOf(OPEN, pos);
    if (s === -1) return null;
    if (count === n) {
      const e = xml.indexOf(CLOSE, s) + CLOSE.length;
      return { start: s, end: e };
    }
    count++;
    pos = s + OPEN.length;
  }
  return null;
}

function findNthTag(
  xml: string,
  openPrefix: string,
  closeTag: string,
  n: number
): { start: number; end: number } | null {
  let count = 0;
  let pos = 0;
  while (pos < xml.length) {
    const s = xml.indexOf(openPrefix, pos);
    if (s === -1) return null;
    const tagEndClose = xml.indexOf(closeTag, s);
    if (tagEndClose === -1) return null;
    const e = tagEndClose + closeTag.length;
    if (count === n) return { start: s, end: e };
    count++;
    pos = e;
  }
  return null;
}

function rebuildCellText(cellXml: string, newText: string): string {
  const tcPrM = cellXml.match(/(<a:tcPr[\s\S]*?<\/a:tcPr>|<a:tcPr[^>]*\/>)/);
  const tcPr = tcPrM ? tcPrM[0] : "";

  const bpM = cellXml.match(/(<a:bodyPr[^>]*(?:\/>|>[\s\S]*?<\/a:bodyPr>))/);
  const bodyPr = bpM ? bpM[0] : "<a:bodyPr/>";

  const lsM = cellXml.match(/(<a:lstStyle[^>]*(?:\/>|>[\s\S]*?<\/a:lstStyle>))/);
  const lstStyle = lsM ? lsM[0] : "<a:lstStyle/>";

  const rPrM = cellXml.match(/<a:rPr([^>]*)(?:\/>|>([\s\S]*?)<\/a:rPr>)/);
  let rPr = '<a:rPr lang="ko-KR" dirty="0"/>';
  if (rPrM) {
    rPr = rPrM[0].endsWith("/>")
      ? rPrM[0]
      : `<a:rPr${rPrM[1]}>${rPrM[2]}</a:rPr>`;
  }

  const ppM = cellXml.match(/<a:pPr[^>]*(?:\/>|>[\s\S]*?<\/a:pPr>)/);
  const pPr = ppM ? ppM[0] : "";

  const run = newText ? `<a:r>${rPr}<a:t>${escXml(newText)}</a:t></a:r>` : "";
  const txBody = `<a:txBody>${bodyPr}${lstStyle}<a:p>${pPr}${run}</a:p></a:txBody>`;
  return `<a:tc>${txBody}${tcPr}</a:tc>`;
}

function setCell(
  slideXml: string,
  tblN: number,
  rowN: number,
  colN: number,
  newText: string
): string {
  const tbl = findNthTable(slideXml, tblN);
  if (!tbl) return slideXml;
  let tblXml = slideXml.slice(tbl.start, tbl.end);

  const row = findNthTag(tblXml, "<a:tr", "</a:tr>", rowN);
  if (!row) return slideXml;
  let rowXml = tblXml.slice(row.start, row.end);

  const cell = findNthTag(rowXml, "<a:tc>", "</a:tc>", colN);
  if (!cell) return slideXml;
  const cellXml = rowXml.slice(cell.start, cell.end);

  const newCell = rebuildCellText(cellXml, newText);
  rowXml = rowXml.slice(0, cell.start) + newCell + rowXml.slice(cell.end);
  tblXml = tblXml.slice(0, row.start) + rowXml + tblXml.slice(row.end);
  return slideXml.slice(0, tbl.start) + tblXml + slideXml.slice(tbl.end);
}

// ── 슬라이드 제거 / 추가 ─────────────────────────────────────

function findRidForTarget(relsXml: string, target: string): string | null {
  const escaped = target.replace(/\//g, "\\/");
  const m = relsXml.match(new RegExp(`<Relationship Id="(rId\\d+)"[^>]*Target="${escaped}"`));
  return m ? m[1] : null;
}

function removeSldIdByRid(presXml: string, rid: string): string {
  const re = new RegExp(`<p:sldId\\b[^>]*\\br:id="${rid}"[^>]*\\/>`);
  return presXml.replace(re, "");
}

function removeRelationshipById(relsXml: string, rid: string): string {
  const re = new RegExp(`<Relationship Id="${rid}"[^>]*\\/>`);
  return relsXml.replace(re, "");
}

function removeContentTypeOverride(ctXml: string, slideNum: number): string {
  const re = new RegExp(`<Override PartName="/ppt/slides/slide${slideNum}\\.xml"[^>]*\\/>`);
  return ctXml.replace(re, "");
}

function removeSlidesCompletely(
  zip: JSZip,
  presXml: string,
  presRels: string,
  ctXml: string,
  startNum: number,
  endNum: number
): { presXml: string; presRels: string; ctXml: string } {
  for (let i = startNum; i <= endNum; i++) {
    const target = `slides/slide${i}.xml`;
    const rid = findRidForTarget(presRels, target);
    if (rid) {
      presXml = removeSldIdByRid(presXml, rid);
      presRels = removeRelationshipById(presRels, rid);
    }
    ctXml = removeContentTypeOverride(ctXml, i);
    zip.remove(`ppt/slides/slide${i}.xml`);
    zip.remove(`ppt/slides/_rels/slide${i}.xml.rels`);
  }
  return { presXml, presRels, ctXml };
}

function addSlideToPresentation(presentationXml: string, slideNum: number): string {
  const allIds = [...presentationXml.matchAll(/id="(\d+)"/g)].map(m => parseInt(m[1]));
  const maxId = allIds.length ? Math.max(...allIds) : 256;
  const newId = maxId + 1;
  const rId = `rId${slideNum + 200}`;
  const newSldId = `<p:sldId id="${newId}" r:id="${rId}"/>`;
  return presentationXml.replace("</p:sldIdLst>", `${newSldId}</p:sldIdLst>`);
}

function addSlideToRels(relsXml: string, slideNum: number): string {
  const rId = `rId${slideNum + 200}`;
  const target = `slides/slide${slideNum}.xml`;
  const TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
  const newRel = `<Relationship Id="${rId}" Type="${TYPE}" Target="${target}"/>`;
  return relsXml.replace("</Relationships>", `${newRel}</Relationships>`);
}

function addSlideToContentTypes(ctXml: string, slideNum: number): string {
  const CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
  const part = `/ppt/slides/slide${slideNum}.xml`;
  if (ctXml.includes(part)) return ctXml;
  return ctXml.replace(
    "</Types>",
    `<Override PartName="${part}" ContentType="${CT}"/></Types>`
  );
}

// ── 별첨 슬라이드 데이타 입력 ───────────────────────────────────

export type Finding = {
  seq: number;
  grade: "위험" | "미흡";
  riskType: string;
  content: string;
  actionRequest: string;
};

function fillBulletinSlide(slideXml: string, f: Finding): string {
  let xml = slideXml;
  xml = setCell(xml, 0, 1, 0, f.grade === "위험" ? "위 험" : "미 흡");

  const checkIdx = RISK_COLS[f.riskType] ?? 6;
  for (let i = 0; i < 7; i++) {
    xml = setCell(xml, 1, 1, i, i === checkIdx ? "■" : "□");
  }

  xml = setCell(xml, 2, 2, 1, f.content);
  xml = setCell(xml, 2, 2, 3, f.actionRequest);

  return xml;
}

// ── PUBLIC: 정기평가 PPTX 생성 ──────────────────────────────────────

export type RegularPptxInput = {
  siteName: string;
  constructionPeriod: string;
  amount: string;
  progress: string;
  mainWork: string;
  inspectionStart: string;
  inspectionEnd: string;
  inspectors: string;
  docScore: string;
  fieldScore: string;
  findings: Finding[];
};

export async function generateRegularPptx(
  templateBuffer: Buffer,
  data: RegularPptxInput
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(templateBuffer);

  let presXml = await zip.file("ppt/presentation.xml")!.async("string");
  let presRels = await zip.file("ppt/_rels/presentation.xml.rels")!.async("string");
  let ctXml = await zip.file("[Content_Types].xml")!.async("string");

  const slideFiles = Object.keys(zip.files).filter(k =>
    /^ppt\/slides\/slide\d+\.xml$/.test(k)
  );
  const totalSlides = slideFiles.length;

  // slide1(표지): "점검 현장"/"점검 기간"/"점검 인원"은 한 행에 [라볘,값]이 같이 있으목 replaceAdjacentCell 사용
  let s1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
  s1 = replaceAdjacentCell(s1, "점검 현장", `${data.siteName}`);
  s1 = replaceAdjacentCell(s1, "점검 기간", `${data.inspectionStart} ~ ${data.inspectionEnd}`);
  s1 = replaceAdjacentCell(s1, "점검 인원", data.inspectors);
  zip.file("ppt/slides/slide1.xml", s1);

  // slide2(결과요약): "현장명"/"공사기간" 등은 다음 행의 같은 열 위지에 값이 있으목 replaceCellBelow 사용
  let s2 = await zip.file("ppt/slides/slide2.xml")!.async("string");
  s2 = replaceCellBelow(s2, "현＀장＀명", data.siteName);
  s2 = replaceCellBelow(s2, "공사＀기간", data.constructionPeriod);
  s2 = replaceCellBelow(s2, "공사＀금액", data.amount);
  s2 = replaceCellBelow(s2, "공정율", `${data.progress}`);
  zip.file("ppt/slides/slide2.xml", s2);

  // slide3(서류부문 점수 표)는 다중 행 구조가 복잡해 이보 수정에서는 직접 쓰지 않음 (원본 시트 점수 유지)

  const templateSlideXml = await zip.file("ppt/slides/slide5.xml")!.async("string");
  const templateRels = await zip.file("ppt/slides/_rels/slide5.xml.rels")!.async("string");

  ({ presXml, presRels, ctXml } = removeSlidesCompletely(
    zip, presXml, presRels, ctXml, 5, totalSlides
  ));

  const numFindings = data.findings.length;
  for (let i = 0; i < numFindings; i++) {
    const newNum = 5 + i;
    let slideXml = fillBulletinSlide(templateSlideXml, data.findings[i]);
    zip.file(`ppt/slides/slide${newNum}.xml`, slideXml);
    zip.file(`ppt/slides/_rels/slide${newNum}.xml.rels`, templateRels);
    presXml = addSlideToPresentation(presXml, newNum);
    presRels = addSlideToRels(presRels, newNum);
    ctXml = addSlideToContentTypes(ctXml, newNum);
  }

  zip.file("ppt/presentation.xml", presXml);
  zip.file("ppt/_rels/presentation.xml.rels", presRels);
  zip.file("[Content_Types].xml", ctXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ── PUBLIC: 중처법 PPTX 생성 ────────────────────────────────────────

export type SapaFinding = {
  seq: number;
  itemName: string;
  detailContent: string;
  actionRequest: string;
  result: "우수" | "양호" | "미흡";
};

export type SapaPptxInput = {
  siteName: string;
  inspectionDate: string;
  inspectors: string;
  overallResult: "우수" | "양호" | "미흡";
  findings: SapaFinding[];
};

export async function generateSapaPptx(
  templateBuffer: Buffer,
  data: SapaPptxInput
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(templateBuffer);

  let presXml = await zip.file("ppt/presentation.xml")!.async("string");
  let presRels = await zip.file("ppt/_rels/presentation.xml.rels")!.async("string");
  let ctXml = await zip.file("[Content_Types].xml")!.async("string");

  const slideFiles = Object.keys(zip.files).filter(k =>
    /^ppt\/slides\/slide\d+\.xml$/.test(k)
  );
  const totalSlides = slideFiles.length;

  let s1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
  s1 = replaceTextBoxContent(s1, "현 장 명 :", data.siteName);
  s1 = replaceTextBoxContent(s1, "점 검 일 :", data.inspectionDate);
  s1 = replaceTextBoxContent(s1, "점 검 자 :", data.inspectors);
  zip.file("ppt/slides/slide1.xml", s1);

  let s2 = await zip.file("ppt/slides/slide2.xml")!.async("string");
  zip.file("ppt/slides/slide2.xml", s2);

  const templateSlideXml = await zip.file("ppt/slides/slide3.xml")!.async("string");
  const templateRels = await zip.file("ppt/slides/_rels/slide3.xml.rels")!.async("string");

  ({ presXml, presRels, ctXml } = removeSlidesCompletely(
    zip, presXml, presRels, ctXml, 3, totalSlides
  ));

  for (let i = 0; i < data.findings.length; i++) {
    const f = data.findings[i];
    const newNum = 3 + i;
    let slideXml = templateSlideXml;

    slideXml = replaceAdjacentCell(slideXml, "세부 내용", f.detailContent);
    slideXml = replaceAdjacentCell(slideXml, "확 인 사 항", f.detailContent);
    slideXml = replaceAdjacentCell(slideXml, "조 지 요 구", f.actionRequest);

    zip.file(`ppt/slides/slide${newNum}.xml`, slideXml);
    zip.file(`ppt/slides/_rels/slide${newNum}.xml.rels`, templateRels);
    presXml = addSlideToPresentation(presXml, newNum);
    presRels = addSlideToRels(presRels, newNum);
    ctXml = addSlideToContentTypes(ctXml, newNum);
  }

  zip.file("ppt/presentation.xml", presXml);
  zip.file("ppt/_rels/presentation.xml.rels", presRels);
  zip.file("[Content_Types].xml", ctXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ── 텍스트 교체 보조함수 ──────────────────────────────────────────

/** 이 함수는 설정 레이아웃: 라볘이 있는 셜의 바로 다음 셌(같은 행)의 값을 교체 */
function replaceAdjacentCell(slideXml: string, label: string, newValue: string): string {
  const tblRegex = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  return slideXml.replace(tblRegex, (tbl) => {
    const rowRegex = /<a:tr[\s\S]*?<\/a:tr>/g;
    return tbl.replace(rowRegex, (row) => {
      const cells: string[] = [];
      const cellRegex = /<a:tc>[\s\S]*?<\/a:tc>/g;
      let m;
      while ((m = cellRegex.exec(row)) !== null) cells.push(m[0]);

      const normalLabel = label.replace(/\s+/g, "");
      const targetIdx = cells.findIndex(c => {
        const t = [...c.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map(x => x[1]).join("");
        return t.replace(/\s+/g, "") === normalLabel ||
               t.replace(/\s+/g, "").includes(normalLabel);
      });

      if (targetIdx === -1 || targetIdx + 1 >= cells.length) return row;
      const newCell = rebuildCellText(cells[targetIdx + 1], newValue);
      cells[targetIdx + 1] = newCell;
      let newRow = row;
      let offset = 0;
      const allCellMatches = [...row.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)];
      allCellMatches.forEach((cm, i) => {
        if (i < cells.length) {
          const origLen = cm[0].length;
          newRow = newRow.slice(0, cm.index! + offset) + cells[i] + newRow.slice(cm.index! + offset + origLen);
          offset += cells[i].length - origLen;
        }
      });
      return newRow;
    });
  });
}

/**
 * 라볘이 있는 항(row N)의 같은 열 위지(col)에, 다음 항(row N+1)의 값을 교체한다.
 * 예: row0=["현장명","공사기간"], row1=["제주 삼다수...","25.12.8~27.7.30"] —
 * "현장명"은 row0 col0에 있으며, 그 값은 row1 col0(바로 아됌 다음 행의 같은 열)에 있다.
 */
function replaceCellBelow(slideXml: string, label: string, newValue: string): string {
  const tblRegex = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  const tbls = [...slideXml.matchAll(tblRegex)];
  for (const tblM of tbls) {
    const tbl = tblM[0];
    const tblStart = tblM.index!;
    const rowMatches = [...tbl.matchAll(/<a:tr[\s\S]*?<\/a:tr>/g)];
    const normalLabel = label.replace(/\s+/g, "");
    for (let ri = 0; ri < rowMatches.length - 1; ri++) {
      const row = rowMatches[ri][0];
      const cells = [...row.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)].map(m => m[0]);
      const colIdx = cells.findIndex(c => {
        const t = [...c.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map(x => x[1]).join("");
        return t.replace(/\s+/g, "") === normalLabel;
      });
      if (colIdx === -1) continue;

      const nextRow = rowMatches[ri + 1][0];
      const nextCells = [...nextRow.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)];
      if (colIdx >= nextCells.length) continue;
      const targetCell = nextCells[colIdx];
      const newCell = rebuildCellText(targetCell[0], newValue);
      const newNextRow =
        nextRow.slice(0, targetCell.index!) + newCell + nextRow.slice(targetCell.index! + targetCell[0].length);
      const newTbl =
        tbl.slice(0, rowMatches[ri + 1].index!) + newNextRow + tbl.slice(rowMatches[ri + 1].index! + rowMatches[ri + 1][0].length);
      return slideXml.slice(0, tblStart) + newTbl + slideXml.slice(tblStart + tbl.length);
    }
  }
  return slideXml;
}

function replaceTextBoxContent(slideXml: string, labelHint: string, newValue: string): string {
  const spRegex = /<p:sp>[\s\S]*?<\/p:sp>/g;
  return slideXml.replace(spRegex, (sp) => {
    const text = [...sp.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map(m => m[1]).join("");
    if (text.includes(labelHint)) {
      const firstParaIdx = sp.indexOf("<a:p>");
      const secondParaStart = sp.indexOf("<a:p>", firstParaIdx + 5);
      if (secondParaStart !== -1) {
        const secondParaEnd = sp.indexOf("</a:p>", secondParaStart) + "</a:p>".length;
        const firstRPr = sp.slice(secondParaStart).match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/);
        const rPr = firstRPr ? firstRPr[0] : '<a:rPr lang="ko-KR" dirty="0"/>';
        const newPara = `<a:p><a:r>${rPr}<a:t>${escXml(newValue)}</a:t></a:r></a:p>`;
        return sp.slice(0, secondParaStart) + newPara + sp.slice(secondParaEnd);
      }
    }
    return sp;
  });
}
