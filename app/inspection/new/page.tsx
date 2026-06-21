"use client";
import { useState, useEffect } from "react";
import { useRouter } from "next/navigation";

type Finding = {
  id: string;
  section: "서류" | "현장";
  grade?: "위험" | "미흡";
  content: string;
  actionRequest: string;
};

const INITIAL_FINDING = (): Finding => ({
  id: crypto.randomUUID(),
  section: "현장",
  grade: "미흡",
  content: "",
  actionRequest: "",
});

// 템플릿 원본 드롭다운 옵션 (xlsx 평가결과 시트 P15/P16/P17 데이터유효성검사 값 그대로 — 임의값 입력 금지)
const MONTHLY_PROGRESS_OPTIONS = [
  { label: "1~5%", value: "1" },
  { label: "6~8%", value: "1.02" },
  { label: "8~10%", value: "1.04" },
  { label: "10%이상", value: "1.06" },
];
const RISK_WORK_OPTIONS = [
  { label: "일반", value: "1" },
  { label: "거푸집&동바리", value: "1.01" },
  { label: "철골설치", value: "1.02" },
  { label: "T/C사용", value: "1.03" },
];
const HELPER_OPERATION_OPTIONS = [
  { label: "1명이상 운영", value: "1" },
  { label: "미운영", value: "1.02" },
];

export default function NewInspectionPage() {
  const router = useRouter();
  const [submitting, setSubmitting] = useState(false);

  // 기본 정보
  const [type, setType] = useState<"정기평가" | "준수평가">("정기평가");
  const [siteName, setSiteName] = useState("");
  const [constructionPeriod, setConstructionPeriod] = useState("");
  const [amount, setAmount] = useState("");
  const [progress, setProgress] = useState("");
  const [inspectionStart, setInspectionStart] = useState("");
  const [inspectionEnd, setInspectionEnd] = useState("");
  const [inspectors, setInspectors] = useState("");
  const [siteManager, setSiteManager] = useState("");
  const [safetyManager, setSafetyManager] = useState("");
  const [docScore, setDocScore] = useState("");
  const [fieldScore, setFieldScore] = useState("");
  // 보정계수 하위 3건 (xlsx P15/P16/P17 → O13=SUM(P15:P17)/3 자동 계산)
  const [monthlyProgressFactor, setMonthlyProgressFactor] = useState("1");
  const [riskWorkFactor, setRiskWorkFactor] = useState("1");
  const [helperOperationFactor, setHelperOperationFactor] = useState("1");
  // 주요작업 상세 3건 (xlsx N7 + pptx 슬라이드2 표1[주요작업]에 동일하게 반영됨)
  const [civilWorkDetail, setCivilWorkDetail] = useState("");
  const [concreteWorkDetail, setConcreteWorkDetail] = useState("");
  const [wetWorkDetail, setWetWorkDetail] = useState("");

  // 지적사항
  const [findings, setFindings] = useState<Finding[]>([INITIAL_FINDING()]);

  function updateFinding(id: string, patch: Partial<Finding>) {
    setFindings((prev) => prev.map((f) => (f.id === id ? { ...f, ...patch } : f)));
  }
  function removeFinding(id: string) {
    setFindings((prev) => prev.filter((f) => f.id !== id));
  }
  function addFinding() {
    setFindings((prev) => [...prev, INITIAL_FINDING()]);
  }

  function handleSectionChange(id: string, section: "서류" | "현장") {
    if (section === "서류") {
      updateFinding(id, { section, grade: undefined });
    } else {
      updateFinding(id, { section, grade: "미흡" });
    }
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!siteName || !inspectionStart) {
      alert("현장명과 점검일은 필수입니다.");
      return;
    }
    setSubmitting(true);
    try {
      const inspRes = await fetch("/api/inspections", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type, siteName, constructionPeriod, amount, progress,
          inspectionStart, inspectionEnd, inspectors,
          siteManager, safetyManager, docScore, fieldScore,
          monthlyProgressFactor, riskWorkFactor, helperOperationFactor,
          civilWorkDetail, concreteWorkDetail, wetWorkDetail,
          status: "점검완료",
        }),
      });
      const { id: inspectionId } = await inspRes.json();

      await Promise.all(
        findings.filter((f) => f.content.trim()).map((f) =>
          fetch("/api/findings", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ ...f, inspectionId }),
          })
        )
      );

      router.push("/dashboard");
    } catch (e: any) {
      alert("오류: " + e.message);
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white border-b px-6 py-4 shadow-sm">
        <h1 className="text-xl font-bold text-gray-900">점검 등록</h1>
      </header>

      <form onSubmit={handleSubmit} className="max-w-3xl mx-auto px-6 py-8 space-y-8">
        {/* 점검 유형 */}
        <section className="bg-white rounded-xl shadow-sm p-6">
          <h2 className="text-base font-semibold text-gray-800 mb-4">점검 유형</h2>
          <div className="flex gap-4">
            {(["정기평가", "준수평가"] as const).map((t) => (
              <label key={t} className="flex items-center gap-2 cursor-pointer">
                <input type="radio" checked={type === t} onChange={() => setType(t)} className="accent-blue-600" />
                <span className="text-sm font-medium">{t}</span>
              </label>
            ))}
          </div>
        </section>

        {/* 기본 정보 */}
        <section className="bg-white rounded-xl shadow-sm p-6 space-y-4">
          <h2 className="text-base font-semibold text-gray-800">기본 정보</h2>
          <div className="grid grid-cols-2 gap-4">
            <Field label="현장명 *" value={siteName} onChange={setSiteName} />
            <Field label="점검자" value={inspectors} onChange={setInspectors} />
            <Field label="점검 시작일 *" type="date" value={inspectionStart} onChange={setInspectionStart} />
            <Field label="점검 종료일" type="date" value={inspectionEnd} onChange={setInspectionEnd} />
          </div>
          {type === "정기평가" && (
            <>
              <div className="grid grid-cols-2 gap-4 pt-2">
                <Field label="공사기간" value={constructionPeriod} onChange={setConstructionPeriod} />
                <Field label="공사금액" value={amount} onChange={setAmount} />
                <Field label="공정율(%)" value={progress} onChange={setProgress} />
                <Field label="현장소장" value={siteManager} onChange={setSiteManager} />
                <Field label="안전관리자" value={safetyManager} onChange={setSafetyManager} />
                <Field label="서류부문 점수" value={docScore} onChange={setDocScore} />
                <Field label="현장부문 점수" value={fieldScore} onChange={setFieldScore} />
              </div>

              <div className="pt-4 border-t">
                <p className="text-sm font-medium text-gray-700 mb-3">보정계수 (3건 평균으로 자동 산정)</p>
                <div className="grid grid-cols-3 gap-4">
                  <SelectField label="월 공정률 보정" value={monthlyProgressFactor} onChange={setMonthlyProgressFactor} options={MONTHLY_PROGRESS_OPTIONS} />
                  <SelectField label="주 위험공종 진행" value={riskWorkFactor} onChange={setRiskWorkFactor} options={RISK_WORK_OPTIONS} />
                  <SelectField label="안전보조원 운영" value={helperOperationFactor} onChange={setHelperOperationFactor} options={HELPER_OPERATION_OPTIONS} />
                </div>
              </div>

              <div className="pt-4 border-t">
                <p className="text-sm font-medium text-gray-700 mb-3">주요작업 (PPTX 슬라이드2 표와 xlsx에 동일하게 반영)</p>
                <div className="grid grid-cols-1 gap-3">
                  <Field label="토목공사 상세" value={civilWorkDetail} onChange={setCivilWorkDetail} />
                  <Field label="철콘공사 상세" value={concreteWorkDetail} onChange={setConcreteWorkDetail} />
                  <Field label="습식공사 상세" value={wetWorkDetail} onChange={setWetWorkDetail} />
                </div>
              </div>
            </>
          )}
        </section>

        {/* 지적사항 */}
        <section className="bg-white rounded-xl shadow-sm p-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-base font-semibold text-gray-800">지적사항 ({findings.length})</h2>
            <button type="button" onClick={addFinding}
              className="text-sm text-blue-600 border border-blue-200 rounded px-3 py-1 hover:bg-blue-50">
              + 추가
            </button>
          </div>
          <div className="space-y-4">
            {findings.map((f, idx) => (
              <div key={f.id} className="border rounded-lg p-4 space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-sm font-medium text-gray-700">#{idx + 1}</span>
                  {findings.length > 1 && (
                    <button type="button" onClick={() => removeFinding(f.id)}
                      className="text-xs text-red-500 hover:text-red-700">제거</button>
                  )}
                </div>
                <div className={`grid gap-3 ${f.section === "현장" ? "grid-cols-2" : "grid-cols-1"}`}>
                  <div>
                    <label className="text-xs text-gray-500 mb-1 block">부문</label>
                    <select value={f.section} onChange={(e) => handleSectionChange(f.id, e.target.value as any)}
                      className="w-full border rounded px-2 py-1.5 text-sm">
                      <option value="현장">현장</option>
                      <option value="서류">서류</option>
                    </select>
                  </div>
                  {f.section === "현장" && (
                    <div>
                      <label className="text-xs text-gray-500 mb-1 block">등급</label>
                      <select value={f.grade ?? "미흡"} onChange={(e) => updateFinding(f.id, { grade: e.target.value as any })}
                        className="w-full border rounded px-2 py-1.5 text-sm">
                        <option value="미흡">미흡</option>
                        <option value="위험">위험</option>
                      </select>
                    </div>
                  )}
                </div>
                <div>
                  <label className="text-xs text-gray-500 mb-1 block">지적 내용</label>
                  <textarea value={f.content} onChange={(e) => updateFinding(f.id, { content: e.target.value })}
                    rows={2} placeholder="지적 내용을 입력하세요"
                    className="w-full border rounded px-3 py-2 text-sm resize-none" />
                </div>
                <div>
                  <label className="text-xs text-gray-500 mb-1 block">조치요구사항</label>
                  <textarea value={f.actionRequest} onChange={(e) => updateFinding(f.id, { actionRequest: e.target.value })}
                    rows={2} placeholder="조치요구사항을 입력하세요"
                    className="w-full border rounded px-3 py-2 text-sm resize-none" />
                </div>
              </div>
            ))}
          </div>
        </section>

        <div className="flex gap-3 justify-end">
          <button type="button" onClick={() => router.back()}
            className="px-5 py-2 rounded-lg border text-sm font-medium hover:bg-gray-50">
            취소
          </button>
          <button type="submit" disabled={submitting}
            className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-medium hover:bg-blue-700 disabled:opacity-50">
            {submitting ? "저장 중..." : "저장"}
          </button>
        </div>
      </form>
    </div>
  );
}

function Field({
  label, value, onChange, type = "text",
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  type?: string;
}) {
  return (
    <div>
      <label className="text-xs text-gray-500 mb-1 block">{label}</label>
      <input type={type} value={value} onChange={(e) => onChange(e.target.value)}
        className="w-full border rounded px-3 py-2 text-sm" />
    </div>
  );
}

function SelectField({
  label, value, onChange, options,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  options: { label: string; value: string }[];
}) {
  return (
    <div>
      <label className="text-xs text-gray-500 mb-1 block">{label}</label>
      <select value={value} onChange={(e) => onChange(e.target.value)}
        className="w-full border rounded px-2 py-1.5 text-sm">
        {options.map((o) => (
          <option key={o.value} value={o.value}>{o.label}</option>
        ))}
      </select>
    </div>
  );
}
