export default function HomePage() {
  return (
    <main className="container">
      <section className="card">
        <h1>안전보건 점검관리 웹앱</h1>
        <p>정기안전보건평가와 중대재해처벌법 준수평가를 반자동화하는 1차 개발 뼈대입니다.</p>
        <div style={{ display: "flex", gap: 12, marginTop: 20 }}>
          <a className="btn" href="/dashboard">본사 대시보드</a>
          <a className="btn secondary" href="/action/demo-token">현장 조치화면 예시</a>
        </div>
      </section>
    </main>
  );
}
