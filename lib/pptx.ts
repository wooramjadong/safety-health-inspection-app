/**
 * PPTX 자동 생성 라이브러리
 * 접근 방식: JSZip으로 PPTX(ZIP)을 열고, XML을 문자열 조작
 *
 * 필딩 슬라이드 구조 (정기평가 별첨):
 *   테이블0 [구분]: row1=구넶, row2=위험|미흡
 *   테이블1 [체크박스]: row1=헤더, row2=■|□ ×7
 *   테이블2 [내용]: row2=사진영역, row3=['내용',지적텍스트,'내용',조치텍스트]
 *
 * 주의: 이전에 서로 생성된 별첨 슬라이드(5번이후)는 내용내용마다 개수가 다릅니다.
 * 그뚜 제거하릔 떨어다 paint 해서 slide.xml 파일만 지우고 presentation.xml/
 * rels/[Content_Types].xml에 단글이르어 있는 참조떤 안 지우면 PPTX가 손상릌어
 * PowerPoint/python-pptx에서 열리지 않아서, removeSlideReferences()로 며명하게 제거합니다.
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

// ── 슬라이드 제거 / 추가 ──────────────────────────────────────

/** rels XML에서 펹정 타겟(slides/slideN.xml)을 가리키는 rId 조회 */
function findRidForTarget(relsXml: string, target: string): string | null {
  const escaped = target.replace(/\//g, "\\/");
  const m = relsXml.match(new RegExp(`<Relationship Id="(rId\\d+)"[^>]*Target="${escaped}"`));
  return m ? m[1] : null;
}

/** presentation.xml에서 해당 rId를 가리키릔 <p:sldId> 삭제 */
function removeSldIdByRid(presXml: string, rid: string): string {
  const re = new RegExp(`<p:sldId\\b[^>]*\\br:id="${rid}"[^>]*\\/>`);
  return presXml.replace(re, "");
}

/** rels XML에서 해당 rId의 <Relationship> 삭제 */
function removeRelationshipById(relsXml: string, rid: string): string {
  const re = new RegExp(`<Relationship Id="${rid}"[^>]*\\/>`);
  return relsXml.replace(re, "");
}

/** [Content_Types].xml에서 해당 슬라이드 Override 삭제 */
function removeContentTypeOverride(ctXml: string, slideNum: number): string {
  const re = new RegExp(`<Override PartName="/ppt/slides/slide${slideNum}\\.xml"[^>]*\\/>`);
  return ctXml.replace(re, "");
}

/**
 * 슬라이듌 startNum~endNum(포함)을 완전하개 제거: 파일 삭제 +
 * presentation.xml의 <p:sldId>, rels의 <Relationship>, [Content_Types].xml의
 * Override꺌지 다 제거해서 단글이르어 있는 참조를 남겨도륍 손상 방지.
 */
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
  const rId = `rId${slideNum + 200}`; // 기존 rId와 확실하게 안 고쳙하도록 덩 단위 올랐
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

  let s1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
  s1 = replaceTextBoxContent(s1, "점검 현장", `${data.siteName}`);
  s1 = replaceTextBoxContent(s1, "점검 기간", `${data.inspectionStart} ~ ${data.inspectionEnd}`);
  s1 = replaceTextBoxContent(s1, "점검 인원", data.inspectors);
  zip.file("ppt/slides/slide1.xml", s1);

  let s2 = await zip.file("ppt/slides/slide2.xml")!.async("string");
  s2 = replaceAdjacentCell(s2, "현　장　명", data.siteName);
  s2 = replaceAdjacentCell(s2, "공사　기간", data.constructionPeriod);
  s2 = replaceAdjacentCell(s2, "공사　금액", data.amount);
  s2 = replaceAdjacentCell(s2, "공정율", `${data.progress}`);
  zip.file("ppt/slides/slide2.xml", s2);

  let s3 = await zip.file("ppt/slides/slide3.xml")!.async("string");
  s3 = replaceAdjacentCell(s3, "현장 소장", data.docScore);
  zip.file("ppt/slides/slide3.xml", s3);

  const templateSlideXml = await zip.file("ppt/slides/slide5.xml")!.async("string");
  const templateRels = await zip.file("ppt/slides/_rels/slide5.xml.rels")!.async("string");

  // 기존 별첨 슬라이드(5~끝) 완전하개 제거 — 단그릑이름에서는
  // presentation.xml의 sldId, rels의 Relationship도 함ꪌ 상제해야
  // PPTX가 손상릌지 않뀈듬니다.
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

  // 기존 3번 이후 슬라이듌 완전하개 제거 (참조 함ꪌ 상제)
  ({ presXml, presRels, ctXml } = removeSlidesCompletely(
    zip, presXml, presRels, ctXml, 3, totalSlides
  ));

  for (let i = 0; i < data.findings.length; i++) {
    const f = data.findings[i];
    const newNum = 3 + i;
    let slideXml = templateSlideXml;

    slideXml = replaceAdjacentCell(slideXml, "세부 내용", f.detailContent);
    slideXml = replaceAdjacentCell(slideXml, "확 인 사 항", f.detailContent);
    slideXml = replaceAdjacentCell(slideXml, "조 치 요 구", f.actionRequest);

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
