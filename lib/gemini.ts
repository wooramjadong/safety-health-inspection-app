// 글자수 한도 (셀 크기·폰트 기준 분석값)
export const CELL_LIMITS = {
  별첨_내용: 35,    // 10.1cm × 1.3cm, 11pt, 약 2줄
  서류부문: 35,     // 15.7cm × 0.8cm, 10.5pt, 1줄
  현장부문_내용: 30, // 좁은 셀
  조치요구: 40,
} as const;

/**
 * 텍스트가 maxChars를 초과하면 Gemini로 요약, 이하면 원문 반환
 * 폰트 크기는 고정 — 텍스트 내용 자체를 줄임
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
              text: `안전보건 점검 지적사항을 ${maxChars}자 이내로 핵심만 간결하게 요약하세요.\n안전 용어와 핵심 위험 내용은 반드시 유지하세요.\n요약문만 출력하고 다른 말은 하지 마세요.\n\n원문: ${text}`,
            }],
          }],
          generationConfig: { maxOutputTokens: 80, temperature: 0.1 },
        }),
      }
    );
    const data = await res.json();
    const summary: string = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? "";
    if (summary && summary.length <= maxChars) return summary;
    // 요약이 아직 길면 앞부분만
    return summary.slice(0, maxChars) || text.slice(0, maxChars);
  } catch (e) {
    console.error("Gemini 요약 실패:", e);
    return text.slice(0, maxChars);
  }
}

/** 규칙 기반 위험유형 추출 (Gemini 미사용) */
export function detectRiskType(text: string): string {
  if (/추락|난간|개구부|고소|작업발판|단부/.test(text)) return "추락";
  if (/낙하|비래|낙석/.test(text)) return "낙하";
  if (/충돌|협착|끼임/.test(text)) return "충돌/협착";
  if (/화재|가스|폭발|용접/.test(text)) return "화재";
  if (/질식|밀폐/.test(text)) return "질식";
  if (/감전|전선|누전|접지/.test(text)) return "감전";
  return "기타";
}
