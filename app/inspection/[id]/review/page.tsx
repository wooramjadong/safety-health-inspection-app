"use client";
import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";

type ChecklistItem = { row: number; group: string; item: string };
type Match = {
  findingId: string;
  content: string;
  grade?: string;
  matchedRow: number | null;
  matchedGroup: string;
  matchedItem: string;
};

export default function ReviewPage() {
  const router = useRouter();
  const params = useParams();
  const inspectionId = params.id as string;

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [fieldMatches, setFieldMatches] = useState<Match[]>([]);
  const [docMatches, setDocMatches] = useState<Match[]>([]);
  const [fieldChecklist, setFieldChecklist] = useState<ChecklistItem[]>([]);
  const [docChecklist, setDocChecklist] = useState<ChecklistItem[]>([]);
  const [overrides, setOverrides] = useState<Record<string, number>>({});
  const [generating, setGenerating] = useState(false);

  useEffect(() => {
    fetch("/api/match-preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ inspectionId }),
    })
      .then((r) => r.json())
      .then((d) => {
        if (d.error) { setError(d.error); return; }
        setFieldMatches(d.fieldMatches ?? []);
        setDocMatches(d.docMatches ?? []);
        setFieldChecklist(d.fieldChecklist ?? []);
        setDocChecklist(d.docChecklist ?? []);
      })
      .catch(() => setError("매징 분석을 불러올 수 없습니다."))
      .finally(() => setLoading(false));
  }, [inspectionId]);

  function setOverride(findingId: string, row: number) {
    setOverrides((prev) => ({ ...prev, [findingId]: row }));
  }

  async function handleConfirm() {
    setGenerating(true);
    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ inspectionId, matchOverrides: overrides }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      alert(`생성 완료\nPPTX: ${data.pptxUrl}\n조치링크: ${data.actionLink}`);
      router.push("/dashboard");
    } catch (e: any) {
      alert("오류: " + e.message);
    } finally {
      setGenerating(false);
    }
  }

  if (loading) {
    return <div className="min-h-screen flex items-center justify-center text-gray-500">AI 매징 분석 중...</div>;
  }
  if (error) {
    return <div className="min-h-screen flex items-center justify-center text-gray-500">오류: {error}</div>;
  }

  const noFindings = fieldMatches.length === 0 && docMatches.length === 0;

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white border-b px-6 py-4 shadow-sm">
        <h1 className="text-xl font-bold text-gray-900">AI 매징 확인</h1>
        <p className="text-sm text-gray-500 mt-0.5">지적사항이 어느 점검항목(현장부문/서류부문)에 매징되었는지 확인하고, 필요하시맩 항목을 변경한 다음 생성하세요.</p>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-8 space-y-8">
        {fieldMatches.length > 0 && (
          <section className="bg-white rounded-xl shadow-sm p-6">
            <h2 className="text-base font-semibold text-gray-800 mb-4">현장부문 지적사항 ({fieldMatches.length})</h2>
            <div className="space-y-3">
              {fieldMatches.map((m) => (
                <MatchRow key={m.findingId} match={m} checklist={fieldChecklist}
                  onChange={(row) => setOverride(m.findingId, row)} />
              ))}
            </div>
          </section>
        )}

        {docMatches.length > 0 && (
          <section className="bg-white rounded-xl shadow-sm p-6">
            <h2 className="text-base font-semibold text-gray-800 mb-4">서류부문 지적사항 ({docMatches.length})</h2>
            <div className="space-y-3">
              {docMatches.map((m) => (
                <MatchRow key={m.findingId} match={m} checklist={docChecklist}
                  onChange={(row) => setOverride(m.findingId, row)} />
              ))}
            </div>
          </section>
        )}

        {noFindings && (
          <p className="text-gray-500 text-center py-10">지적사항이 없습니다. 다음 다음으로 바로 생성하셔률 됩니다.</p>
        )}

        <div className="flex justify-end gap-3">
          <button onClick={() => router.back()}
            className="px-5 py-2 rounded-lg border text-sm font-medium hover:bg-gray-50">
            취소
          </button>
          <button onClick={handleConfirm} disabled={generating}
            className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-medium hover:bg-blue-700 disabled:opacity-50">
            {generating ? "생성 중..." : "확인하고 생성"}
          </button>
        </div>
      </main>
    </div>
  );
}

function MatchRow({
  match, checklist, onChange,
}: {
  match: Match;
  checklist: ChecklistItem[];
  onChange: (row: number) => void;
}) {
  const [selected, setSelected] = useState<number>(match.matchedRow ?? checklist[0]?.row ?? 0);
  const groups = Array.from(new Set(checklist.map((c) => c.group)));
  const matchedButMissing = match.matchedRow === null;

  return (
    <div className="border rounded-lg p-3">
      <p className="text-sm text-gray-800 mb-2">
        {match.content}
        {match.grade && <span className="ml-2 text-xs text-gray-400">[{match.grade}]</span>}
      </p>
      <div className="flex items-center gap-2">
        <span className="text-xs text-gray-400 whitespace-nowrap">매징항목:</span>
        <select
          value={selected}
          onChange={(e) => {
            const row = parseInt(e.target.value, 10);
            setSelected(row);
            onChange(row);
          }}
          className="flex-1 border rounded px-2 py-1.5 text-sm"
        >
          {groups.map((g) => (
            <optgroup key={g} label={g}>
              {checklist.filter((c) => c.group === g).map((c) => (
                <option key={c.row} value={c.row}>{c.item}</option>
              ))}
            </optgroup>
          ))}
        </select>
      </div>
      {matchedButMissing && (
        <p className="text-xs text-amber-600 mt-1">⚠ AI가 적합한 항목을 찾지 봇해 금 항목이 자동 선택되었습니다. 확인해주세요.</p>
      )}
    </div>
  );
}
