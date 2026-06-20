import { getDashboardSummary } from "@/lib/mock";

export default async function DashboardPage() {
  const summary = await getDashboardSummary();
  return (
    <main className="container grid">
      <h1>본사 안전보건팀 대시보드</h1>
      <section className="grid grid-3">
        <div className="card"><b>운영 현장</b><h2>{summary.sites}</h2></div>
        <div className="card"><b>연간 점검</b><h2>{summary.inspections}</h2></div>
        <div className="card"><b>미조치</b><h2>{summary.openActions}</h2></div>
      </section>
      <section className="card">
        <h2>점검 등록</h2>
        <form className="grid" action="/api/inspections" method="post">
          <div><label>점검유형</label><select name="type" className="input"><option>정기안전보건평가</option><option>중처법준수평가</option></select></div>
          <div><label>현장명</label><input name="siteName" className="input" placeholder="예: 광주KT&G 기숙사 현장" /></div>
          <div><label>점검자</label><input name="inspectors" className="input" placeholder="예: 현태완 부장, 최우람 과장" /></div>
          <button className="btn" type="submit">점검 임시등록</button>
        </form>
      </section>
      <section className="card">
        <h2>지적사항 등록 흐름</h2>
        <table className="table"><tbody>
          <tr><th>1</th><td>조치 전 사진 1장 업로드</td></tr>
          <tr><th>2</th><td>AI 위험유형·구분·감점·지적내용 추천</td></tr>
          <tr><th>3</th><td>본사 담당자가 수정 후 확정</td></tr>
          <tr><th>4</th><td>PPTX/XLSX/PDF 생성 및 Drive 저장</td></tr>
        </tbody></table>
      </section>
    </main>
  );
}
