import { ChecklistItem } from "./checklist-data";

// 자 수 한도 (셀 크기·폰트 기준 분석값)
export const CELL_LIMITS = {
  별첨_내용: 35,    // 10.1cm × 1.3cm, 11pt, 약 2줄
  서류부문: 35,     // 15.7cm × 0.8cm, 10.5pt, 1줄
  현장부문_내용: 30, // 좋은 셀
  조치요구: 40,
} as const;

/**
 * 텍스트가 maxChars를 초과하면 Gemini로 요약, 이하면 원문 반환
 * 폰트 크기는 고정 — 텍스트 내용 자신을 줄임
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
    // 요약이 어지 김리면 앞부분럸
    return summary.slice(0, maxChars) || text.slice(0, maxChars);
  } catch (e) {
    console.error("Gemini 요약 실패:", e);
    return text.slice(0, maxChars);
  }
}

/** 그앮적 경한 감지 괜타 추출 (Gemini 불필요) */
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
 * 지적사항 텍스트와 가장 관리있는 체크리스트 행(row)을 찾는다.
 * xlsx의 안전보건 현장부문/서류부문 시트는 158건/47건의 고정 점검항목으로
 * 구성되어 있어, PPTX 지적사항이 어느 항목에 해당하는지 AI가 자동 매징한다.
 *
 * @returns 매쾤된 체크리스트의 row 번호 (xlsx 행 번호). 실패 시 null.
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

/** Gemini 불필요/실패 시 키워드 중백도 기반 폴백 매징 */
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
