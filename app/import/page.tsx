"use client";
import { useState } from "react";
import { useRouter } from "next/navigation";

type FieldFinding = { row: number; content: string; grade: "위험" | "미흡"; group: string; item: string };
type DocFinding = { row: number; content: string; group: string; item: string };
type RegularData = {
  siteName: string; inspectionPeriod: string; inspectors: string; mainWork: string;
  amount: string; constructionPeriod: string; managerInfo: string;
  docTotalScore: string; docSectionScore: string; fieldSectionScore: string;
  fieldFindings: FieldFinding[]; docFindings: DocFinding[];
};
type SapaFinding = { slide: number; itemName: string; detail: string; action: string; result: string };
type SapaData = { siteName: string; inspectionDate: string; inspectors: string; findings: SapaFinding[] };

export default function ImportPage() {
  const router = useRouter();
  const [type, setType] = useState<"정기평가" | "준수평가">("정기평가");
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [confirming, setConfirming] = useState(false);
  const [error, setError] = useState("");
  const [regularData, setRegularData] = useState<RegularData | null>(null);
  const [sapaData, setSapaData] = useState<SapaData | null>(null);

  async function handleParse() {
    if (!file) { alert("파일을 선택하세요."); return; }
    setLoading(true);
    setError("");
    setRegularData(null);
    setSapaData(null);
    try {
      const form = new FormData();
      form.append("file", file);
      form.append("type", type);
      const res = await fetch("/api/import-historical", { method: "POST", body: form });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error);
      if (type === "정기평가") setRegularData(json.data);
      else setSapaData(json.data);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  async function handleConfirm() {
    const data = type === "정기평가" ? regularData : sapaData;
    if (!data) return;
    setConfirming(true);
    try {
      const res = await fetch("/api/import-historical/confirm", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type, data }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error);
      alert(`임포트 완료: 지적사항 ${json.importedFindings}개 저장되었습니다.`);
      router.push("/dashboard");
    } catch (e: any) {
      alert("오류: " + e.message);
    } finally {
      setConfirming(false);
    }
  }

  function updateFieldFinding(idx: number, patch: Partial<FieldFinding>) {
    if (!regularData) return;
    const next = [...regularData.fieldFindings];
    next[idx] = { ...next[idx], ...patch };
    setRegularData({ ...regularData, fieldFindings: next });
  }
  function removeFieldFinding(idx: number) {
    if (!regularData) return;
    setRegularData({ ...regularData, fieldFindings: regularData.fieldFindings.filter((_, i) => i !== idx) });
  }
  function updateDocFinding(idx: number, patch: Partial<DocFinding>) {
    if (!regularData) return;
    const next = [...regularData.docFindings];
    next[idx] = { ...next[idx], ...patch };
    setRegularData({ ...regularData, docFindings: next });
  }
  function removeDocFinding(idx: number) {
    if (!regularData) return;
    setRegularData({ ...regularData, docFindings: regularData.docFindings.filter((_, i) => i !== idx) });
  }
  function updateSapaFinding(idx: number, patch: Partial<SapaFinding>) {
    if (!sapaData) return;
    const next = [...sapaData.findings];
    next[idx] = { ...next[idx], ...patch };
    setSapaData({ ...sapaData, findings: next });
  }
  function removeSapaFinding(idx: number) {
    if (!sapaData) return;
    setSapaData({ ...sapaData, findings: sapaData.findings.filter((_, i) => i !== idx) });
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white border-b px-6 py-4 shadow-sm">
        <h1 className="text-xl font-bold text-gray-900">과거 점검파일 가져오기</h1>
        <p className="text-sm text-gray-500 mt-0.5">이미 수행된 점검 파일을 업로드해 지적사항을 DB에 누적 저장합니다. 저장 전에 내용을 확인·수정할 수 있습니다.</p>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-8 space-y-6">
        <section className="bg-white rounded-xl shadow-sm p-6">
          <div className="flex gap-4 mb-4">
            {(["정기평가", "준수평가"] as const).map((t) => (
              <label key={t} className="flex items-center gap-2 cursor-pointer">
                <input type="radio" checked={type === t} onChange={() => { setType(t); setFile(null); setRegularData(null); setSapaData(null); }} className="accent-blue-600" />
                <span className="text-sm font-medium">{t} {t === "정기평가" ? "(.xlsx)" : "(.pptx)"}</span>
              </label>
            ))}
          </div>
          <div className="flex gap-3 items-center">
            <input
              type="file"
              accept={type === "정기평가" ? ".xlsx" : ".pptx"}
              onChange={(e) => setFile(e.target.files?.[0] ?? null)}
              className="text-sm flex-1 border rounded px-3 py-2"
            />
            <button onClick={handleParse} disabled={!file || loading}
              className="bg-blue-600 text-white text-sm px-4 py-2 rounded hover:bg-blue-700 disabled:opacity-50">
              {loading ? "분석 중..." : "파싱 보이기"}
            </button>
          </div>
          {error && <p className="text-sm text-red-600 mt-3">오류: {error}</p>}
        </section>

        {regularData && (
          <>
            <section className="bg-white rounded-xl shadow-sm p-6">
              <h2 className="text-base font-semibold text-gray-800 mb-3">기본정보</h2>
              <div className="grid grid-cols-2 gap-2 text-sm">
                <p><span className="text-gray-500">현장명:</span> {regularData.siteName}</p>
                <p><span className="text-gray-500">점검기간:</span> {regularData.inspectionPeriod}</p>
                <p><span className="text-gray-500">점검자:</span> {regularData.inspectors}</p>
                <p><span className="text-gray-500">공사기간:</span> {regularData.constructionPeriod}</p>
                <p><span className="text-gray-500">공사금액:</span> {regularData.amount}</p>
                <p><span className="text-gray-500">담당자:</span> {regularData.managerInfo}</p>
              </div>
            </section>

            <section className="bg-white rounded-xl shadow-sm p-6">
              <h2 className="text-base font-semibold text-gray-800 mb-3">현장부문 지적사항 ({regularData.fieldFindings.length})</h2>
              <div className="space-y-2">
                {regularData.fieldFindings.map((f, i) => (
                  <div key={i} className="border rounded-lg p-3 flex items-start gap-2">
                    <span className={`text-xs px-2 py-0.5 rounded shrink-0 ${f.grade === "위험" ? "bg-red-100 text-red-700" : "bg-yellow-100 text-yellow-700"}`}>{f.grade}</span>
                    <div className="flex-1">
                      <textarea value={f.content} onChange={(e) => updateFieldFinding(i, { content: e.target.value })}
                        rows={1} className="w-full text-sm border-0 focus:ring-0 resize-none p-0" />
                      <p className="text-xs text-gray-400 mt-1">{f.group} / {f.item} (row{f.row})</p>
                    </div>
                    <button onClick={() => removeFieldFinding(i)} className="text-xs text-red-500 shrink-0">제거</button>
                  </div>
                ))}
              </div>
            </section>

            <section className="bg-white rounded-xl shadow-sm p-6">
              <h2 className="text-base font-semibold text-gray-800 mb-3">서류부문 지적사항 ({regularData.docFindings.length})</h2>
              <div className="space-y-2">
                {regularData.docFindings.map((f, i) => (
                  <div key={i} className="border rounded-lg p-3 flex items-start gap-2">
                    <div className="flex-1">
                      <textarea value={f.content} onChange={(e) => updateDocFinding(i, { content: e.target.value })}
                        rows={1} className="w-full text-sm border-0 focus:ring-0 resize-none p-0" />
                      <p className="text-xs text-gray-400 mt-1">{f.group} / {f.item} (row{f.row})</p>
                    </div>
                    <button onClick={() => removeDocFinding(i)} className="text-xs text-red-500 shrink-0">제거</button>
                  </div>
                ))}
              </div>
            </section>
          </>
        )}

        {sapaData && (
          <>
            <section className="bg-white rounded-xl shadow-sm p-6">
              <h2 className="text-base font-semibold text-gray-800 mb-3">기본정보</h2>
              <div className="grid grid-cols-2 gap-2 text-sm">
                <p><span className="text-gray-500">현장명:</span> {sapaData.siteName}</p>
                <p><span className="text-gray-500">점검일:</span> {sapaData.inspectionDate}</p>
                <p><span className="text-gray-500">점검자:</span> {sapaData.inspectors}</p>
              </div>
            </section>
            <section className="bg-white rounded-xl shadow-sm p-6">
              <h2 className="text-base font-semibold text-gray-800 mb-3">항목별 지적사항 ({sapaData.findings.length})</h2>
              <div className="space-y-2">
                {sapaData.findings.map((f, i) => (
                  <div key={i} className="border rounded-lg p-3 flex items-start gap-2">
                    <div className="flex-1">
                      <p className="text-xs text-gray-400 mb-1">{f.itemName}</p>
                      <textarea value={f.detail} onChange={(e) => updateSapaFinding(i, { detail: e.target.value })}
                        rows={2} className="w-full text-sm border rounded px-2 py-1 resize-none" />
                    </div>
                    <button onClick={() => removeSapaFinding(i)} className="text-xs text-red-500 shrink-0">제거</button>
                  </div>
                ))}
              </div>
            </section>
          </>
        )}

        {(regularData || sapaData) && (
          <div className="flex justify-end gap-3">
            <button onClick={() => router.push("/dashboard")} className="px-5 py-2 rounded-lg border text-sm font-medium hover:bg-gray-50">대시보드로</button>
            <button onClick={handleConfirm} disabled={confirming}
              className="px-5 py-2 rounded-lg bg-green-600 text-white text-sm font-medium hover:bg-green-700 disabled:opacity-50">
              {confirming ? "저장 중..." : "확인하고 DB에 저장"}
            </button>
          </div>
        )}
      </main>
    </div>
  );
}
