/**
 * PPTX 자동 생성 라이브러리
 * 접근 방식: JSZip으로 PPTX(ZIP)을 열고, XML을 문자열 조작
 *
 * 필딩 슬라이드 구조 (정기평가 별첨):
 *   테이블0 [구분]: row1=구년, row2=위험|미흡
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

/** N번째 <a:tbl> 위치 반환 (PPTX 안 테이블은 중첩 없음) */
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
  openPrefix: string, // e.g. "<a:tr"
  closeTag: string,   // e.g. "</a:tr>"
  n: number
): { start: number; end: number } | null {
  let count = 0;
  let pos = 0;
  while (pos < xml.length) {
    const s = xml.indexOf(openPrefix, pos);
    if (s === -1) return null;
    // 자체단혀("/>")인지 확인
    const tagEndClose = xml.indexOf(closeTag, s);
    if (tagEndClose === -1) return null;
    const e = tagEndClose + closeTag.length;
    if (count === n) return { start: s, end: e };
    count++;
    pos = e;
  }
  return null;
}

/** 셀의 txBody 콘텐츠를 새 텍스트로 교체 (tcPr 보존) */
function rebuildCellText(cellXml: string, newText: string): string {
  // tcPr (셀 테두리/내지) 보존
  const tcPrM = cellXml.match(/(<a:tcPr[\s\S]*?<\/a:tcPr>|<a:tcPr[^>]*\/>)/);
  const tcPr = tcPrM ? tcPrM[0] : "";

  // bodyPr
  const bpM = cellXml.match(/(<a:bodyPr[^>]*(?:\/>|>[\s\S]*?<\/a:bodyPr>))/);
  const bodyPr = bpM ? bpM[0] : "<a:bodyPr/>";

  // lstStyle
  const lsM = cellXml.match(/(<a:lstStyle[^>]*(?:\/>|>[\s\S]*?<\/a:lstStyle>))/);
  const lstStyle = lsM ? lsM[0] : "<a:lstStyle/>";

  // 첫 번째 런의 rPr (폰트 정보 보존)
  const rPrM = cellXml.match(/<a:rPr([^>]*)(?:\/>|>([\s\S]*?)<\/a:rPr>)/);
  let rPr = '<a:rPr lang="ko-KR" dirty="0"/>';
  if (rPrM) {
    rPr = rPrM[0].endsWith("/>")
      ? rPrM[0]
      : `<a:rPr${rPrM[1]}>${rPrM[2]}</a:rPr>`;
  }

  // pPr
  const ppM = cellXml.match(/<a:pPr[^>]*(?:\/>|>[\s\S]*?<\/a:pPr>)/);
  const pPr = ppM ? ppM[0] : "";

  const run = newText ? `<a:r>${rPr}<a:t>${escXml(newText)}</a:t></a:r>` : "";
  const txBody = `<a:txBody>${bodyPr}${lstStyle}<a:p>${pPr}${run}</a:p></a:txBody>`;
  return `<a:tc>${txBody}${tcPr}</a:tc>`;
}

/** 슬라이드 XML에서 tbl[tblN].row[rowN].cell[colN] 내용 교체 */
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

// ── 슬라이드 복제 / 프레젠테이션.xml 업데이트 ─────────────────────

async function cloneSlide(
  zip: JSZip,
  srcSlideNum: number,   // 복사원본 슬라이드 번호 (1-based)
  newSlideNum: number    // 새 슬라이드 번호 (1-based)
): Promise<void> {
  const srcPath = `ppt/slides/slide${srcSlideNum}.xml`;
  const srcRelsPath = `ppt/slides/_rels/slide${srcSlideNum}.xml.rels`;
  const dstPath = `ppt/slides/slide${newSlideNum}.xml`;
  const dstRelsPath = `ppt/slides/_rels/slide${newSlideNum}.xml.rels`;

  const srcXml = await zip.file(srcPath)!.async("string");
  const srcRels = await zip.file(srcRelsPath)!.async("string");

  zip.file(dstPath, srcXml);
  // 레이아웃 관계만 유지 (이미지 rId는 Phase2에서 처리)
  zip.file(dstRelsPath, srcRels);
}

function addSlideToPresentation(presentationXml: string, slideNum: number): string {
  // 마지막 <p:sldId> 다음에 새 슬라이드 삽입
  const allIds = [...presentationXml.matchAll(/id="(\d+)"/g)].map(m => parseInt(m[1]));
  const maxId = allIds.length ? Math.max(...allIds) : 256;
  const newId = maxId + 1;
  const rId = `rId${slideNum + 100}`; // 충돌 방지
  const newSldId = `<p:sldId id="${newId}" r:id="${rId}"/>`;

  // sldIdLst의 마지막 항목 다음에 삽입
  return presentationXml.replace("</p:sldIdLst>", `${newSldId}</p:sldIdLst>`);
}

function addSlideToRels(relsXml: string, slideNum: number): string {
  const rId = `rId${slideNum + 100}`;
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

// ── 별첨 슬라이드 데이터 입력 ──────────────────────────────────────

export type Finding = {
  seq: number;
  grade: "위험" | "미흡";
  riskType: string;   // "추락", "낙하", "충돌/협착", "화재", "질식", "감전", "기타"
  content: string;    // 확정문구 (35자 이내)
  actionRequest: string;
};

/** 별첨 슬라이드 1장에 데이터 입력 */
function fillBulletinSlide(slideXml: string, f: Finding): string {
  let xml = slideXml;

  // [Table 0] 구분: row1 → 위험 | 미흡
  xml = setCell(xml, 0, 1, 0, f.grade === "위험" ? "위 험" : "미 흡");

  // [Table 1] 체크박스: row1, col0-6
  const checkIdx = RISK_COLS[f.riskType] ?? 6;
  for (let i = 0; i < 7; i++) {
    xml = setCell(xml, 1, 1, i, i === checkIdx ? "■" : "□");
  }

  // [Table 2] 내용: row2(idx=2), col1 = 지적내용
  xml = setCell(xml, 2, 2, 1, f.content);

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

  // 프레젠테이션.xml으로 슬라이드 목록 파악
  let presXml = await zip.file("ppt/presentation.xml")!.async("string");
  let presRels = await zip.file("ppt/_rels/presentation.xml.rels")!.async("string");
  let ctXml = await zip.file("[Content_Types].xml")!.async("string");

  // 현재 심라이드 수 파악
  const slideFiles = Object.keys(zip.files).filter(k =>
    /^ppt\/slides\/slide\d+\.xml$/.test(k)
  );
  const totalSlides = slideFiles.length;

  // 슬라이드 1-4: 텍스트 대입
  // — 슬라이드 1 (표지)
  let s1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
  // 해당 셀 내용 교체 (라벨 다음 셀 = 값)
  s1 = replaceTextBoxContent(s1, "점검 현장", `${data.siteName}`);
  s1 = replaceTextBoxContent(s1, "점검 기간", `${data.inspectionStart} ~ ${data.inspectionEnd}`);
  s1 = replaceTextBoxContent(s1, "점검 인원", data.inspectors);
  zip.file("ppt/slides/slide1.xml", s1);

  // — 슬라이드 2 (결과요약)
  let s2 = await zip.file("ppt/slides/slide2.xml")!.async("string");
  s2 = replaceAdjacentCell(s2, "현　장　명", data.siteName);
  s2 = replaceAdjacentCell(s2, "공사　기간", data.constructionPeriod);
  s2 = replaceAdjacentCell(s2, "공사　금액", data.amount);
  s2 = replaceAdjacentCell(s2, "공정율", `${data.progress}`);
  zip.file("ppt/slides/slide2.xml", s2);

  // — 슬라이드 3, 4: 점수만 대입 (내용은 사용자 입력)
  let s3 = await zip.file("ppt/slides/slide3.xml")!.async("string");
  s3 = replaceAdjacentCell(s3, "현장 소장", data.docScore);
  zip.file("ppt/slides/slide3.xml", s3);

  // 슬라이드 5를 별첨 템플릿으로 사용
  const templateSlideXml = await zip.file("ppt/slides/slide5.xml")!.async("string");

  // 기존 별첨 슬라이드 (5번이후) 모두 제거
  for (let i = 5; i <= totalSlides; i++) {
    zip.remove(`ppt/slides/slide${i}.xml`);
    zip.remove(`ppt/slides/_rels/slide${i}.xml.rels`);
  }

  // 지적사항수만큼 슬라이드 생성
  const templateRels = await zip.file("ppt/slides/_rels/slide5.xml.rels")!.async("string");
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

  // 마지막에 컴파일된 presentation.xml, rels, CT 저장
  zip.file("ppt/presentation.xml", presXml);
  zip.file("ppt/_rels/presentation.xml.rels", presRels);
  zip.file("[Content_Types].xml", ctXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ── PUBLIC: 중처법 PPTX 생성 ────────────────────────────────────────

export type SapaFinding = {
  seq: number;
  itemName: string;      // 안전보건 예산편성 등
  detailContent: string; // 확인사항 세부내용
  actionRequest: string; // 조치요구
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

  // 슬라이드 1 (표지)
  let s1 = await zip.file("ppt/slides/slide1.xml")!.async("string");
  s1 = replaceTextBoxContent(s1, "현 장 명 :", data.siteName);
  s1 = replaceTextBoxContent(s1, "점 검 일 :", data.inspectionDate);
  s1 = replaceTextBoxContent(s1, "점 검 자 :", data.inspectors);
  zip.file("ppt/slides/slide1.xml", s1);

  // 슬라이드 2 (결과요약) — 우수/양호/미흡 체크표시 (Phase2에서 정교화)
  // Phase1: 표라만 변경
  let s2 = await zip.file("ppt/slides/slide2.xml")!.async("string");
  zip.file("ppt/slides/slide2.xml", s2);

  // 슬라이드 3을 준수사항 템플릿으로 사용
  const templateSlideXml = await zip.file("ppt/slides/slide3.xml")!.async("string");
  const templateRels = await zip.file("ppt/slides/_rels/slide3.xml.rels")!.async("string");

  // 기존 3번 이후 슬라이드 제거
  for (let i = 3; i <= totalSlides; i++) {
    zip.remove(`ppt/slides/slide${i}.xml`);
    zip.remove(`ppt/slides/_rels/slide${i}.xml.rels`);
  }

  for (let i = 0; i < data.findings.length; i++) {
    const f = data.findings[i];
    const newNum = 3 + i;
    let slideXml = templateSlideXml;

    // 항목명 교체
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

/** 라벨에 인접한 셀의 값을 교체 (label 첫 발견 시 다음 셀) */
function replaceAdjacentCell(slideXml: string, label: string, newValue: string): string {
  // 모든 테이블의 모든 row에서 라벨 셀 찾기
  const tblRegex = /<a:tbl>[\s\S]*?<\/a:tbl>/g;
  return slideXml.replace(tblRegex, (tbl) => {
    const rowRegex = /<a:tr[\s\S]*?<\/a:tr>/g;
    return tbl.replace(rowRegex, (row) => {
      // 이 row에 label이 있으면 다음 셀 교체
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
      // row재조립 (셀이 자면 cells가 다 포함)
      let newRow = row;
      // 각 셀을 순서대로 교체
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

/** 텍스트박스(라벨 포함)의 다음 스팔 내용 교체 */
function replaceTextBoxContent(slideXml: string, labelHint: string, newValue: string): string {
  // 라벨 힌트가 있는 <p:sp>를 찾아 그 다음에 오는 실제 값 텍스트 교체
  // 단순 문자열 대체: label힌트를 가진 sp에서 값 텍스트 일부 교체
  const spRegex = /<p:sp>[\s\S]*?<\/p:sp>/g;
  return slideXml.replace(spRegex, (sp) => {
    const text = [...sp.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map(m => m[1]).join("");
    if (text.includes(labelHint)) {
      // 라벨 힌트 다음에 오는 <a:t> 수정
      // 라벨이 있는 셋에서 다음 <a:p>를 찾아 교체
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
