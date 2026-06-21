/**
 * 로컬 단독 테스트 스크립트 — Google Drive/Sheets/API 키 없이 실행 가능
 *
 * 목적: 다른 컴퓨터(또는 다른 AI 어시스턴트)가 이 레포만 clone 받아서, Google 서비스
 * 계정이나 Drive 템플릿 ID 같은 비공개 설정 없이도 코드를 수정하고 바로 결과를
 * 눈으로 확인할 수 있게 한다. generateRegularXlsx/generateRegularPptx는 Drive나
 * Sheets에 전혀 의존하지 않고 템플릿 Buffer + 데이터 객체만 받으므로, templates/
 * 폴더에 들어있는 실제 템플릿 파일을 그대로 읽어서 바로 호출할 수 있다.
 *
 * 실행 방법:
 *   npm install
 *   npm run test:local
 *
 * 결과물은 output/ 폴더에 생성된다 (xlsx/pptx). GEMINI_API_KEY가 없으면 AI 매칭/요약은
 * 키워드 기반 폴백으로 동작한다 — 정상이며 에러가 아니다.
 */

import * as fs from "fs";
import * as path from "path";
import { generateRegularXlsx } from "../lib/xlsx";
import { generateRegularPptx } from "../lib/pptx";
import { detectRiskType } from "../lib/gemini";

const ROOT = path.join(__dirname, "..");
const TEMPLATES_DIR = path.join(ROOT, "templates");
const OUTPUT_DIR = path.join(ROOT, "output");

const XLSX_TEMPLATE = path.join(TEMPLATES_DIR, "정기안전보건 평가표(제주 삼다수).xlsx");
const PPTX_TEMPLATE = path.join(TEMPLATES_DIR, "01. 정기 안전보건 평가 결과(260611_제주 삼다수).pptx");

// ── 더미 입력값 — 자유롭게 바꿔서 테스트해도 됨 ──────────────────────────────
const SITE_NAME = "테스트현장 신축공사";

const FIELD_FINDINGS = [
  { content: "3층 슬라브 단부 안전난간대 미설치 및 작업자 접근 통제 미흡", grade: "위험" as const },
  { content: "지하 1층 환기구 개구부 덮개 미설치", grade: "위험" as const },
  { content: "가설계단 손잡이 일부 구간 미설치", grade: "미흡" as const },
];

const DOC_FINDINGS = [
  { content: "산업안전보건관리비 사용계획서 재작성 필요" },
  { content: "위험성평가 결과 게시 누락" },
];

async function main() {
  if (!fs.existsSync(XLSX_TEMPLATE) || !fs.existsSync(PPTX_TEMPLATE)) {
    console.error("템플릿 파일을 찾을 수 없습니다. templates/ 폴더를 확인하세요:");
    console.error(" -", XLSX_TEMPLATE);
    console.error(" -", PPTX_TEMPLATE);
    process.exit(1);
  }
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

  console.log("① xlsx 생성 중...");
  const xlsxTplBuf = fs.readFileSync(XLSX_TEMPLATE);
  const { buffer: xlsxBuf, scores } = await generateRegularXlsx(xlsxTplBuf, {
    siteName: SITE_NAME,
    inspectionPeriod: "2026-06-20 ~ 2026-06-21",
    inspectors: "김안전 부장, 박점검 과장",
    civilWorkDetail: "지하 2층 토공사 및 흙막이 보강",
    concreteWorkDetail: "지상 3~8층 골조 콘크리트 타설",
    wetWorkDetail: "지하 1층 화장실 방수 및 미장",
    amount: "8500000000",
    constructionPeriod: "2025-09-01 ~ 2026-12-31",
    managerInfo: "이현장 소장 정안전 차장",
    monthlyProgressFactor: "1.02",   // 유효값: 1, 1.02, 1.04, 1.06
    riskWorkFactor: "1.01",          // 유효값: 1, 1.01, 1.02, 1.03
    helperOperationFactor: "1.02",   // 유효값: 1, 1.02
    fieldFindings: FIELD_FINDINGS,
    docFindings: DOC_FINDINGS,
  });
  console.log("   점수 계산 결과:", JSON.stringify(scores, null, 2));

  console.log("\n② pptx 생성 중 (xlsx와 동일한 scores 사용)...");
  const pptxTplBuf = fs.readFileSync(PPTX_TEMPLATE);
  const findings = FIELD_FINDINGS.map((f, idx) => ({
    seq: idx + 1,
    grade: f.grade,
    riskType: detectRiskType(f.content),
    content: f.content,
    actionRequest: "즉시 시정 및 재발방지 대책 수립",
  }));
  const pptxBuf = await generateRegularPptx(pptxTplBuf, {
    siteName: SITE_NAME,
    constructionPeriod: "2025-09-01 ~ 2026-12-31",
    amount: "8500000000",
    progress: "35",
    civilWorkDetail: "지하 2층 토공사 및 흙막이 보강",
    concreteWorkDetail: "지상 3~8층 골조 콘크리트 타설",
    wetWorkDetail: "지하 1층 화장실 방수 및 미장",
    inspectionStart: "2026-06-20",
    inspectionEnd: "2026-06-21",
    inspectors: "김안전 부장, 박점검 과장",
    siteManager: "이현장 소장",
    safetyManager: "정안전 차장",
    scores,
    findings,
  });

  const xlsxOut = path.join(OUTPUT_DIR, `${SITE_NAME}_정기평가.xlsx`);
  const pptxOut = path.join(OUTPUT_DIR, `${SITE_NAME}_정기평가.pptx`);
  fs.writeFileSync(xlsxOut, xlsxBuf);
  fs.writeFileSync(pptxOut, pptxBuf);

  console.log("\n✅ 완료. output/ 폴더를 확인하세요:");
  console.log(" -", xlsxOut);
  console.log(" -", pptxOut);

  if (!process.env.GEMINI_API_KEY) {
    console.log("\n참고: GEMINI_API_KEY가 설정되지 않아 AI 매칭/요약은 키워드 기반 폴백으로 동작했습니다.");
    console.log("      실제 서비스에서는 .env.local에 GEMINI_API_KEY를 넣으면 더 정확한 매칭이 됩니다.");
  }
}

main().catch((e) => {
  console.error("❌ 테스트 실패:", e);
  process.exit(1);
});
