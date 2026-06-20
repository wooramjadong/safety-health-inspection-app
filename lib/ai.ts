export async function suggestFinding(input: { memo: string }) {
  // 1차 뼈대: 실제 이미지 분석 API 연결 전 규칙 기반 추천
  const text = input.memo;
  const isFall = /난간|개구부|고소|작업발판|추락/.test(text);
  const isElectric = /전선|피복|접지|감전/.test(text);
  const riskType = isFall ? "추락" : isElectric ? "감전" : "기타";
  const grade = /미설치|미착용|개구부|난간/.test(text) ? "위험" : "미흡";
  const deduction = grade === "위험" ? 2 : 1;
  return {
    riskType,
    grade,
    deduction,
    summary: text ? text.slice(0, 40) : "지적내용을 입력하세요",
    actionRequest: riskType === "추락" ? "안전시설 설치 후 작업 실시" : "즉시 개선 후 재발방지 관리"
  };
}
