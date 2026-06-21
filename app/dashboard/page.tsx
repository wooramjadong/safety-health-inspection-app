"use client";
import { useEffect, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";

type Inspection = {
  id: string;
  siteName: string;
  type: string;
  inspectionStart: string;
  status: string;
  pptxUrl?: string;
  xlsxUrl?: string;
  actionLink?: string;
};

export default function DashboardPage() {
  const router = useRouter();
  const [inspections, setInspections] = useState<Inspection[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch("/api/inspections")
      .then((r) => r.json())
      .then((d) => setInspections(d))
      .finally(() => setLoading(false));
  }, []);

  function goToReview(id: string, type: string) {
    if (type === "정기평가") {
      router.push(`/inspection/${id}/review`);
    } else {
      // 준수평가는 체크리스트 매징이 없으목 고댁 바로 생성
      handleDirectGenerate(id);
    }
  }

  async function handleDirectGenerate(id: string) {
    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ inspectionId: id }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      alert(`생성 완료\nPPTX: ${data.pptxUrl}\n조치링크: ${data.actionLink}`);
      const updated = await fetch("/api/inspections").then((r) => r.json());
      setInspections(updated);
    } catch (e: any) {
      alert("오류: " + e.message);
    }
  }

  async function copyLink(link: string) {
    await navigator.clipboard.writeText(link);
    alert("링크 복사되었습니다.");
  }

  const statusBadge = (s: string) => {
    const map: Record<string, string> = {
      "점검완료": "bg-blue-100 text-blue-800",
      "생성완료": "bg-green-100 text-green-800",
      "조치중": "bg-yellow-100 text-yellow-800",
      "조치완료": "bg-gray-100 text-gray-700",
    };
    return map[s] ?? "bg-gray-100 text-gray-700";
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white border-b px-6 py-4 flex items-center justify-between shadow-sm">
        <h1 className="text-xl font-bold text-gray-900">안전보건 점검 관리</h1>
        <Link
          href="/inspection/new"
          className="bg-blue-600 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-700 transition"
        >
          + 점검 등록
        </Link>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8">
        {loading ? (
          <p className="text-gray-500 text-center py-20">로딩 중...</p>
        ) : inspections.length === 0 ? (
          <div className="text-center py-20">
            <p className="text-gray-500 mb-4">등록된 점검이 없습니다.</p>
            <Link href="/inspection/new" className="text-blue-600 underline text-sm">첫 점검 등록하기</Link>
          </div>
        ) : (
          <div className="bg-white rounded-xl shadow-sm overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 text-gray-600 text-xs uppercase">
                <tr>
                  <th className="px-4 py-3 text-left">현장명</th>
                  <th className="px-4 py-3 text-left">유형</th>
                  <th className="px-4 py-3 text-left">점검일</th>
                  <th className="px-4 py-3 text-left">상태</th>
                  <th className="px-4 py-3 text-left">다운로드</th>
                  <th className="px-4 py-3 text-left">조치링크</th>
                  <th className="px-4 py-3 text-left">작업</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {inspections.map((insp) => (
                  <tr key={insp.id} className="hover:bg-gray-50">
                    <td className="px-4 py-3 font-medium text-gray-900">{insp.siteName}</td>
                    <td className="px-4 py-3 text-gray-600">{insp.type}</td>
                    <td className="px-4 py-3 text-gray-600">{insp.inspectionStart}</td>
                    <td className="px-4 py-3">
                      <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${statusBadge(insp.status)}`}>
                        {insp.status}
                      </span>
                    </td>
                    <td className="px-4 py-3">
                      {insp.pptxUrl ? (
                        <div className="flex gap-2">
                          <a href={insp.pptxUrl} target="_blank" className="text-blue-600 hover:underline">PPTX</a>
                          {insp.xlsxUrl && (
                            <a href={insp.xlsxUrl} target="_blank" className="text-green-600 hover:underline">xlsx</a>
                          )}
                        </div>
                      ) : <span className="text-gray-400">미생성</span>}
                    </td>
                    <td className="px-4 py-3">
                      {insp.actionLink ? (
                        <button
                          onClick={() => copyLink(insp.actionLink!)}
                          className="text-xs text-blue-600 border border-blue-200 rounded px-2 py-0.5 hover:bg-blue-50"
                        >
                          링크 복사
                        </button>
                      ) : <span className="text-gray-400">미생성</span>}
                    </td>
                    <td className="px-4 py-3">
                      <button
                        onClick={() => goToReview(insp.id, insp.type)}
                        className={insp.pptxUrl
                          ? "text-gray-500 text-xs px-3 py-1 rounded border hover:bg-gray-50"
                          : "bg-blue-600 text-white text-xs px-3 py-1 rounded hover:bg-blue-700"}
                      >
                        {insp.pptxUrl ? "재생성" : "생성"}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </main>
    </div>
  );
}
