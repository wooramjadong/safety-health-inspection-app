"use client";

import { useState } from "react";

type CheckType = "정기안전보건평가" | "중처법준수평가";
type Finding = {
  no: number;
  raw: string;
  photo: string;
  level: "위험" | "미흡";
  risk: string;
  summary: string;
  point: number;
  request: string;
  fixed: boolean;
};

function recommend(raw: string) {
  if (raw.includes("난간") || raw.includes("단부") || raw.includes("추락")) {
    return {
      level: "위험" as const,
      risk: "추락",
      summary: "단부 안전난간대 미설치",
      point: 2,
      request: "안전난간대 설치 및 추락방호조치 후 작업 실시",
    };
  }
  if (raw.includes("위험성평가") || raw.includes("중처법") || raw.includes("증빙")) {
    return {
      level: "미흡" as const,
      risk: "중처법 준수",
      summary: "위험성평가 개선조치 증빙 미흡",
      point: 1,
      request: "개선대책 이행자료와 확인기록 보완",
    };
  }
  return {
    level: "미흡" as const,
    risk: "기타",
    summary: "현장 안전관리 개선 필요",
    point: 1,
    request: "현장 확인 후 적정 안전조치 실시",
  };
}

function escapeHtml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

export default function DashboardPage() {
  const [type, setType] = useState<CheckType>("정기안전보건평가");
  const [site, setSite] = useState("임시 테스트현장 A동 신축공사");
  const [inspector, setInspector] = useState("본사 안전보건팀");
  const [started, setStarted] = useState(false);
  const [token, setToken] = useState("demo-token");
  const [raw, setRaw] = useState("");
  const [photo, setPhoto] = useState("");
  const [findings, setFindings] = useState<Finding[]>([]);

  const totalPoint = findings.reduce((sum, item) => sum + item.point, 0);
  const fixedCount = findings.filter((item) => item.fixed).length;

  function startInspection() {
    setStarted(true);
    setToken(`test-${Date.now()}`);
  }

  function addFinding() {
    if (!started) return alert("먼저 점검 임시등록 후 실행하세요.");
    if (!raw.trim()) return alert("지적 원문을 입력하세요.");
    const ai = recommend(raw);
    setFindings((prev) => [...prev, { no: prev.length + 1, raw, photo, fixed: false, ...ai }]);
    setRaw("");
    setPhoto("");
  }

  function runRegularTest() {
    setType("정기안전보건평가");
    setSite("임시 테스트현장 A동 신축공사");
    setInspector("본사 안전보건팀");
    setStarted(true);
    setToken(`regular-${Date.now()}`);
    const rawText = "계단실 단부 안전난간대가 설치되지 않아 작업자 추락 위험이 있음";
    setFindings([{ no: 1, raw: rawText, photo: "before_guardrail.jpg", fixed: true, ...recommend(rawText) }]);
  }

  function runSapaTest() {
    setType("중처법준수평가");
    setSite("임시 테스트현장 B동 증축공사");
    setInspector("본사 안전보건팀");
    setStarted(true);
    setToken(`sapa-${Date.now()}`);
    const rawText = "위험성평가 개선대책 이행 확인 기록이 일부 누락되어 중처법 준수 증빙 보완 필요";
    setFindings([{ no: 1, raw: rawText, photo: "before_compliance.jpg", fixed: true, ...recommend(rawText) }]);
  }

  function downloadReport() {
    const today = new Date().toISOString().slice(0, 10);
    const rows = findings.map((item) => `
      <tr>
        <td>${item.no}</td>
        <td>${escapeHtml(item.level)}</td>
        <td>${escapeHtml(item.risk)}</td>
        <td>${escapeHtml(item.summary)}</td>
        <td>${item.point}</td>
        <td>${escapeHtml(item.request)}</td>
        <td>${escapeHtml(item.photo || "-")}</td>
        <td>${item.fixed ? "확정" : "미확정"}</td>
      </tr>`).join("");

    const html = `<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(site)}_${escapeHtml(type)}_점검결과</title>
<style>
  body { font-family: Arial, 'Noto Sans KR', sans-serif; padding: 32px; color: #111827; }
  h1 { font-size: 24px; margin-bottom: 8px; }
  .meta { border: 1px solid #ddd; padding: 12px; margin: 16px 0; }
  table { width: 100%; border-collapse: collapse; margin-top: 16px; }
  th, td { border: 1px solid #ddd; padding: 8px; font-size: 13px; text-align: left; }
  th { background: #f3f4f6; }
  .note { margin-top: 24px; font-size: 12px; color: #555; }
</style>
</head>
<body>
  <h1>${escapeHtml(type)} 점검결과서</h1>
  <div class="meta">
    <p><b>현장명:</b> ${escapeHtml(site)}</p>
    <p><b>점검자:</b> ${escapeHtml(inspector)}</p>
    <p><b>작성일:</b> ${today}</p>
    <p><b>조치링크 토큰:</b> ${escapeHtml(token)}</p>
    <p><b>총 지적건수:</b> ${findings.length}건 / <b>총 감점:</b> ${totalPoint}점 / <b>확정:</b> ${fixedCount}건</p>
  </div>
  <table>
    <thead><tr><th>번호</th><th>구분</th><th>위험유형</th><th>확정문구</th><th>감점</th><th>조치요구</th><th>조치 전 사진</th><th>상태</th></tr></thead>
    <tbody>${rows || "<tr><td colspan='8'>등록된 지적사항이 없습니다.</td></tr>"}</tbody>
  </table>
  <p class="note">※ 본 문서는 자동화 테스트용 검토 서류입니다. 향후 기존 PPTX/XLSX 템플릿 형식으로 연결 예정입니다.</p>
</body>
</html>`;

    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${site}_${type}_점검결과.html`;
    a.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="container grid">
      <h1>본사 안전보건팀 대시보드</h1>

      <section className="grid grid-3">
        <div className="card"><b>운영 현장</b><h2>20</h2></div>
        <div className="card"><b>연간 점검</b><h2>40</h2></div>
        <div className="card"><b>미조치</b><h2>{findings.length - fixedCount}</h2></div>
      </section>

      <section className="card grid">
        <h2>점검 임시등록</h2>
        <div className="grid grid-3">
          <div><label>점검유형</label><select className="input" value={type} onChange={(e) => setType(e.target.value as CheckType)}><option>정기안전보건평가</option><option>중처법준수평가</option></select></div>
          <div><label>현장명</label><input className="input" value={site} onChange={(e) => setSite(e.target.value)} /></div>
          <div><label>점검자</label><input className="input" value={inspector} onChange={(e) => setInspector(e.target.value)} /></div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button className="btn" onClick={startInspection}>점검 임시등록 후 실행</button>
          <button className="btn secondary" onClick={runRegularTest}>정기평가 최종테스트</button>
          <button className="btn secondary" onClick={runSapaTest}>중처법 최종테스트</button>
        </div>
      </section>

      {started && (
        <section className="card grid">
          <h2>현장점검 실행</h2>
          <p><b>{type}</b> / {site} / {inspector}</p>
          <p className="badge">조치링크 토큰: {token}</p>
          <div><label>조치 전 사진 파일명</label><input className="input" value={photo} onChange={(e) => setPhoto(e.target.value)} placeholder="예: IMG_001.jpg" /></div>
          <div><label>지적 원문</label><textarea className="input" rows={4} value={raw} onChange={(e) => setRaw(e.target.value)} placeholder="현장에서 확인한 내용을 입력" /></div>
          <button className="btn" onClick={addFinding}>AI 추천 후 지적사항 추가</button>
        </section>
      )}

      <section className="card">
        <h2>지적사항 및 AI 추천 결과</h2>
        <table className="table">
          <thead><tr><th>번호</th><th>사진</th><th>구분</th><th>위험유형</th><th>확정문구</th><th>감점</th><th>조치요구</th><th>상태</th></tr></thead>
          <tbody>
            {findings.map((item, index) => (
              <tr key={item.no}>
                <td>{item.no}</td>
                <td>{item.photo || "-"}</td>
                <td>{item.level}</td>
                <td>{item.risk}</td>
                <td><input className="input" value={item.summary} onChange={(e) => setFindings((prev) => prev.map((x, i) => i === index ? { ...x, summary: e.target.value } : x))} /></td>
                <td>{item.point}</td>
                <td>{item.request}</td>
                <td><button className={item.fixed ? "btn secondary" : "btn"} onClick={() => setFindings((prev) => prev.map((x, i) => i === index ? { ...x, fixed: true } : x))}>{item.fixed ? "확정완료" : "확정"}</button></td>
              </tr>
            ))}
            {findings.length === 0 && <tr><td colSpan={8}>등록된 지적사항이 없습니다.</td></tr>}
          </tbody>
        </table>
      </section>

      {started && (
        <section className="card grid">
          <h2>자료 생성 및 현장 배포</h2>
          <p>총 지적 {findings.length}건 / 총 감점 {totalPoint}점 / 확정 {fixedCount}건</p>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn" onClick={downloadReport}>검토용 점검결과서 다운로드</button>
            <a className="btn secondary" href={`/action/${token}`}>현장 조치 링크 열기</a>
          </div>
        </section>
      )}
    </main>
  );
}
