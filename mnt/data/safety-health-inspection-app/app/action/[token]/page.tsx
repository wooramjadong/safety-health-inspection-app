export default function ActionPage({ params }: { params: { token: string } }) {
  return (
    <main className="container grid">
      <section className="card">
        <h1>현장 조치결과 제출</h1>
        <p className="badge">조치 링크 토큰: {params.token}</p>
        <p>현장 안전관리자는 로그인 없이 지적사항별 조치내용과 조치 후 사진 1장을 제출합니다.</p>
      </section>
      <section className="card">
        <h2>지적사항 1</h2>
        <p><b>지적내용:</b> 계단 및 단부 안전난간대 미설치</p>
        <form className="grid" action="/api/action-submit" method="post" encType="multipart/form-data">
          <input type="hidden" name="token" value={params.token} />
          <div><label>조치내용</label><textarea name="actionText" rows={4} placeholder="예: 계단 단부에 안전난간대를 설치하고 작업자 통행 전 확인 완료" /></div>
          <div><label>조치 후 사진</label><input className="input" type="file" name="afterPhoto" accept="image/*" /></div>
          <button className="btn" type="submit">제출</button>
        </form>
      </section>
    </main>
  );
}
