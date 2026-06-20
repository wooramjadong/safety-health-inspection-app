"use client";
import { use, useState, useEffect } from "react";

type Finding = {
  id: string;
  content: string;
  actionRequest: string;
  grade: string;
  hasPhoto?: boolean;
};

export default function ActionPage({ params }: { params: Promise<{ token: string }> }) {
  const { token } = use(params);
  const [findings, setFindings] = useState<Finding[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [uploaded, setUploaded] = useState<Record<string, boolean>>({});
  const [submitting, setSubmitting] = useState<Record<string, boolean>>({});
  const [done, setDone] = useState(false);

  useEffect(() => {
    // 토큰으로 점검에 연결된 지적사항 조회
    fetch(`/api/findings-by-token?token=${token}`)
      .then((r) => r.json())
      .then((d) => {
        if (d.error) setError(d.error);
        else setFindings(d);
      })
      .catch(() => setError("데이터를 불러올 수 없습니다."))
      .finally(() => setLoading(false));
  }, [token]);

  async function handleUpload(findingId: string, file: File, comment: string) {
    setSubmitting((p) => ({ ...p, [findingId]: true }));
    try {
      const form = new FormData();
      form.append("file", file);
      form.append("token", token);
      form.append("findingId", findingId);
      form.append("comment", comment);
      const res = await fetch("/api/upload-photo", { method: "POST", body: form });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setUploaded((p) => ({ ...p, [findingId]: true }));
    } catch (e: any) {
      alert("업로드 실패: " + e.message);
    } finally {
      setSubmitting((p) => ({ ...p, [findingId]: false }));
    }
  }

  async function handleFinalSubmit() {
    if (!confirm("모든 조치사진을 제출하시겠습니까?")) return;
    const res = await fetch("/api/action-submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token }),
    });
    if (res.ok) setDone(true);
    else alert("제출 실패합니다.");
  }

  if (loading) return <FullCenter>로딩 중...</FullCenter>;
  if (error) return <FullCenter>오류: {error}</FullCenter>;
  if (done) return (
    <FullCenter>
      <div className="text-center">
        <p className="text-2xl font-bold text-green-600 mb-2">제출 완료</p>
        <p className="text-gray-600">조치사진이 성공적으로 제출되었습니다.</p>
      </div>
    </FullCenter>
  );

  const allUploaded = findings.length > 0 && findings.every((f) => uploaded[f.id]);

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white border-b px-6 py-4 shadow-sm">
        <h1 className="text-xl font-bold text-gray-900">조치사진 제출</h1>
        <p className="text-sm text-gray-500 mt-0.5">각 지적사항에 조치사진을 업로드해 주세요.</p>
      </header>

      <main className="max-w-2xl mx-auto px-4 py-6 space-y-4">
        {findings.map((f, idx) => (
          <FindingCard
            key={f.id}
            seq={idx + 1}
            finding={f}
            isUploaded={!!uploaded[f.id]}
            isSubmitting={!!submitting[f.id]}
            onUpload={handleUpload}
          />
        ))}

        {findings.length > 0 && (
          <button
            onClick={handleFinalSubmit}
            disabled={!allUploaded}
            className="w-full bg-green-600 text-white py-3 rounded-xl font-medium text-sm hover:bg-green-700 disabled:opacity-40 disabled:cursor-not-allowed mt-4"
          >
            {allUploaded ? "조치결과 제출" : `해야 할 항목 남음 (${findings.length - Object.keys(uploaded).length}개)`}
          </button>
        )}
      </main>
    </div>
  );
}

function FindingCard({
  seq, finding, isUploaded, isSubmitting, onUpload,
}: {
  seq: number;
  finding: Finding;
  isUploaded: boolean;
  isSubmitting: boolean;
  onUpload: (id: string, file: File, comment: string) => void;
}) {
  const [file, setFile] = useState<File | null>(null);
  const [comment, setComment] = useState("");
  const [preview, setPreview] = useState("");

  function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (!f) return;
    setFile(f);
    setPreview(URL.createObjectURL(f));
  }

  return (
    <div className={`bg-white rounded-xl shadow-sm p-4 border-l-4 ${
      isUploaded ? "border-green-500" : finding.grade === "위험" ? "border-red-500" : "border-yellow-500"
    }`}>
      <div className="flex items-start justify-between mb-2">
        <div>
          <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${
            finding.grade === "위험" ? "bg-red-100 text-red-700" : "bg-yellow-100 text-yellow-700"
          }`}>{finding.grade}</span>
          <span className="ml-2 text-xs text-gray-400">#{seq}</span>
        </div>
        {isUploaded && <span className="text-xs text-green-600 font-medium">✓ 업로드완료</span>}
      </div>
      <p className="text-sm text-gray-800 mb-1 font-medium">{finding.content}</p>
      <p className="text-xs text-gray-500 mb-3">조치요구: {finding.actionRequest}</p>

      {!isUploaded && (
        <div className="space-y-2">
          <label className="block">
            <span className="text-xs text-gray-500">사진 선택</span>
            <input type="file" accept="image/*" capture="environment"
              onChange={handleFile}
              className="mt-1 block w-full text-sm text-gray-600 file:mr-3 file:py-1.5 file:px-3 file:rounded file:border-0 file:text-xs file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
          </label>
          {preview && <img src={preview} alt="미리보기" className="w-full h-32 object-cover rounded" />}
          <textarea
            value={comment} onChange={(e) => setComment(e.target.value)}
            placeholder="조치 내용 메모 (선택)"
            rows={2}
            className="w-full border rounded px-3 py-2 text-sm resize-none" />
          <button
            onClick={() => file && onUpload(finding.id, file, comment)}
            disabled={!file || isSubmitting}
            className="w-full bg-blue-600 text-white py-2 rounded-lg text-sm font-medium hover:bg-blue-700 disabled:opacity-40"
          >
            {isSubmitting ? "업로드 중..." : "사진 업로드"}
          </button>
        </div>
      )}
    </div>
  );
}

function FullCenter({ children }: { children: React.ReactNode }) {
  return (
    <div className="min-h-screen flex items-center justify-center text-gray-600">{children}</div>
  );
}
