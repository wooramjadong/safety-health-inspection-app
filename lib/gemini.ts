import { ChecklistItem } from "./checklist-data";

// 글자수 한도 (셀 크기·폰트 기준 분석값)
export const CELL_LIMITS = {
  별첨_내용: 35,    // PPTX 별첨 10.1cm × 1.3cm, 11pt, 약 2줄
  서류부문: 35,     // PPTX 서류부문 슬라이드 15.7cm × 0.8cm, 10.5pt, 1줄
  현장부문_내용: 30, // PPTX 현장부문 셋은 셀
  조치요구: 40,
} as const;

// xlsx 안전보건 서류부문 H열(의견란) 실제측: 열너바 20.33, 3행 병합(65.7pt), 9pt 썮은고딜.
// 한 줄에 약 10자, 쵝대 5줄 정도 들어감 — 안전하게 45자로 제한.
export const DOC_OPINION_TOTAL_CHARS = 45;

/**
 * 텍스트가 maxChars를 초과하으면 Gemini로 요약, 이하으면 원문 반환
 * 폰트 크기는 고정 — 텍스트 내용 자신을 줄임 (Excel shrinkToFit은 원잘 폰트 크기를 바꿌서 사용하지 않음)
 */
export async function summarizeForCell(
  text: string,
  maxChars: number = CELL_LIMITS.별첨_내용
): Promise<string> {
  if (!text || text.length <= maxChars) return text;

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.warn("GEMINI_API_KEY 없음 — 단순 절삭");
    return text.slice(0, maxChars);
  }

  try {
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{
            parts: [{
              text: `안전보건 점검 지적사항을 ${maxChars}자 이내로 핵시번 간결하게 요약하세요.\n안전 용어와 핵시 위험 내용은 반드시 유지하세요.\n요약문만 출릵하고 다른 말은 하지 마세요.\n\n원문: ${text}`,
            }],
          }],
          generationConfig: { maxOutputTokens: 80, temperature: 0.1 },
        }),
      }
    );
    const data = await res.json();
    const summary: string = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    if (summary && summary.length <= maxChars) return summary;
    return summary.slice(0, maxChars) || text.slice(0, maxChars);
  } catch (e) {
    console.error("Gemini 요약 실패:", e);
    return text.slice(0, maxChars);
  }
}

/**
 * 여러 지적사항이 같은 셀(공유 체크리스트 항목)에 매징될 뙄 하나로 합쳐 쓰는다.
 * 줄변경(\n)으로 구분해 쓰고, 합쳐쁌 글자수가 한도를 초과하면 각 건을 건수대로 나뉘어 Gemini로
 * 감단해띴 총 글자수가 한도 안에 들어가게 한다. 폰트 크기는 안 목 좀이고 내용만 줄인다.
 */
export async function combineFindingTexts(
  texts: string[],
  totalCharBudget: number = DOC_OPINION_TOTAL_CHARS
): Promise<string> {
  const filtered = texts.filter((t) => t && t.trim());
  if (filtered.length === 0) return "";
  if (filtered.length === 1) {
    const t = filtered[0];
    return t.length <= totalCharBudget ? t : await summarizeForCell(t, totalCharBudget);
  }

  const joined = filtered.join("\n");
  if (joined.length <= totalCharBudget) return joined;

  // 합쳐서 동과하면 항목당 권돔 너이어 개별 요약 (줄바꿈 기호 타면은 제외하고 계산)
  const perItemBudget = Math.max(8, Math.floor((totalCharBudget - (filtered.length - 1)) / filtered.length));
  const shortened = await Promise.all(filtered.map((t) => summarizeForCell(t, perItemBudget)));
  return shortened.join("\n");
}

/** 그그적 경로 각장 검월 권타 추출 (Gemini 불필요) */
export function detectRiskType(text: string): string {
  if (/추락|난간|개구부|고소|작업발판|단부/.test(text)) return "추락";
  if (/낙하|뱄똈|낙석/.test(text)) return "낙하";
  if (/충돌|협착|꿸임/.test(text)) return "충돌/협착";
  if (/화재|가스|폭발|용접/.test(text)) return "화재";
  if (/질식|밀폐/.test(text)) return "질식";
  if (/감전|전선|누전|접지/.test(text)) return "감전";
  return "기타";
}

/**
 * 지적사항 텍스트와 가장 관리있는 체크리스트 항(row)을 찾는다.
 * xlsx의 안전보건 현장부문/서류부문 시트는 158건/47건의 고정 점검항목으로
 * 구성되어 있어, PPTX 지적사항이 어느 항목에 해당하는지 AI가 자동 매징한다.
 *
 * @returns 매징된 체크리스트의 row 번호 (xlsx 행 번호). 실패 시 null.
 */
export async function matchChecklistRow(
  findingText: string,
  checklist: ChecklistItem[]
): Promise<number | null> {
  if (!findingText || checklist.length === 0) return null;

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.warn("GEMINI_API_KEY 없음 — 키워드 매징으로 대신");
    return fallbackKeywordMatch(findingText, checklist);
  }

  const listText = checklist
    .map((c, i) => `${i}. [${c.group}] ${c.item}`)
    .join("\n");

  const prompt =
    `다음은 건설현장 안전보건 점검 체크리스트 항목 목록입니다.\n` +
    `아래 "지적사항"과 내용·작업유형이 가장 일썱하는 항목의 번호 하나만 출릵하세요.\n` +
    `숫자만 출릵하고 다른 설루, 점, 텍스트는 절대 포함하지 마세요.\n\n` +
    `[체크리스트]\n${listText}\n\n[지적사항]\n${findingText}`;

  try {
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { maxOutputTokens: 10, temperature: 0 },
        }),
      }
    );
    const data = await res.json();
    const answer: string = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    const idx = parseInt(answer.match(/\d+/)?.[0] ?? "-1", 10);
    if (idx >= 0 && idx < checklist.length) return checklist[idx].row;
    console.warn("Gemini 매징 응답 해석 실패, 키워드 매징으로 대신:", answer);
    return fallbackKeywordMatch(findingText, checklist);
  } catch (e) {
    console.error("Gemini 체크리스트 매징 실패:", e);
    return fallbackKeywordMatch(findingText, checklist);
  }
}

/** Gemini 불필요/실패 시 키워드 중복도 기반 폴백 매징 */
function fallbackKeywordMatch(findingText: string, checklist: ChecklistItem[]): number | null {
  const keywords = findingText.replace(/[^\p{L}\p{N}\s]/gu, " ").split(/\s+/).filter(w => w.length >= 2);
  if (keywords.length === 0) return checklist[0]?.row ?? null;

  let bestRow: number | null = null;
  let bestScore = -1;
  for (const c of checklist) {
    const target = `${c.group} ${c.item}`;
    let score = 0;
    for (const kw of keywords) if (target.includes(kw)) score++;
    if (score > bestScore) {
      bestScore = score;
      bestRow = c.row;
    }
  }
  return bestScore > 0 ? bestRow : (checklist[0]?.row ?? null);
}
