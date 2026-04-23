'use client';

import Image from 'next/image';
import { useState, useCallback, DragEvent } from 'react';
import { HOSPITAL_CONFIGS } from '@/lib/hospitalConfig';

const HOSPITALS = Object.entries(HOSPITAL_CONFIGS).map(([k, v]) => ({ key: k, name: v.name }));

export default function Home() {
  const now = new Date();
  const [files, setFiles]           = useState<File[]>([]);
  const [dragging, setDragging]     = useState(false);
  const [hospitalKey, setHospitalKey] = useState('mellow_sinsa');
  const [hospitalName, setHospitalName] = useState(HOSPITAL_CONFIGS.mellow_sinsa.name);
  const [year, setYear]     = useState(now.getFullYear());
  const [month, setMonth]   = useState(now.getMonth() + 1);
  const [teamName, setTeamName] = useState('열정의시간 마케팅팀');
  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState('');
  const [done, setDone]       = useState(false);

  const addFiles = useCallback((newFiles: FileList | null) => {
    if (!newFiles) return;
    const excelFiles = Array.from(newFiles).filter(f => /\.(xlsx|xls)$/i.test(f.name));
    setFiles(prev => {
      const names = new Set(prev.map(f => f.name));
      return [...prev, ...excelFiles.filter(f => !names.has(f.name))];
    });
    setDone(false);
    setError('');
  }, []);

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragging(false);
    addFiles(e.dataTransfer.files);
  };

  const handleHospitalChange = (key: string) => {
    setHospitalKey(key);
    setHospitalName(HOSPITAL_CONFIGS[key]?.name || '');
  };

  const handleGenerate = async () => {
    if (!files.length) { setError('엑셀 파일을 먼저 업로드하세요.'); return; }
    setLoading(true); setError(''); setDone(false);

    const fd = new FormData();
    files.forEach(f => fd.append('files', f));
    fd.append('hospitalKey',  hospitalKey);
    fd.append('hospitalName', hospitalName);
    fd.append('year',         String(year));
    fd.append('month',        String(month));
    fd.append('teamName',     teamName);

    try {
      const res = await fetch('/api/generate', { method: 'POST', body: fd });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `서버 오류 (${res.status})`);
      }
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = `${hospitalName}_${year}년${month}월_매출분析보고서.pptx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setDone(true);
    } catch (e) {
      setError(String(e));
    } finally {
      setLoading(false);
    }
  };

  const yearOptions = Array.from({ length: 4 }, (_, i) => now.getFullYear() - 1 + i);

  return (
    <div className="min-h-screen" style={{ background: '#F8F8F9' }}>

      {/* ── Header ─────────────────────────────────────── */}
      <header style={{ background: 'var(--red)' }} className="shadow-md">
        <div className="max-w-6xl mx-auto px-6 py-3 flex items-center gap-3">
          <Image src="/logo_passion.png" alt="열정의시간" width={36} height={36} className="rounded-md" />
          <div>
            <h1 className="text-base font-bold text-white leading-tight tracking-tight">
              CRM 매출분석 보고서 생성기
            </h1>
            <p className="text-xs text-red-200">열정의시간 마케팅팀</p>
          </div>
        </div>
      </header>

      {/* ── Top accent bar ──────────────────────────────── */}
      <div className="h-0.5" style={{ background: 'var(--yellow)' }} />

      <div className="max-w-6xl mx-auto px-6 py-7 flex gap-6 items-start">

        {/* ── Sidebar ──────────────────────────────────── */}
        <aside className="w-72 shrink-0 space-y-4">

          {/* Settings card */}
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="px-5 py-3 flex items-center gap-2" style={{ background: 'var(--red)' }}>
              <span className="text-sm font-bold text-white">⚙ 보고서 설정</span>
            </div>
            <div className="p-5 space-y-4">

              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">병원 선택</label>
                <select
                  value={hospitalKey}
                  onChange={e => handleHospitalChange(e.target.value)}
                  className="w-full text-sm border border-slate-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-400 bg-white"
                >
                  {HOSPITALS.map(h => <option key={h.key} value={h.key}>{h.name}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">병원명 (표지)</label>
                <input
                  type="text" value={hospitalName}
                  onChange={e => setHospitalName(e.target.value)}
                  className="w-full text-sm border border-slate-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-400"
                />
              </div>

              <div className="flex gap-2">
                <div className="flex-1">
                  <label className="block text-xs font-semibold text-slate-500 mb-1.5">년도</label>
                  <select value={year} onChange={e => setYear(+e.target.value)}
                    className="w-full text-sm border border-slate-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-400 bg-white">
                    {yearOptions.map(y => <option key={y} value={y}>{y}</option>)}
                  </select>
                </div>
                <div className="flex-1">
                  <label className="block text-xs font-semibold text-slate-500 mb-1.5">월</label>
                  <select value={month} onChange={e => setMonth(+e.target.value)}
                    className="w-full text-sm border border-slate-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-400 bg-white">
                    {Array.from({ length: 12 }, (_, i) => i + 1).map(m =>
                      <option key={m} value={m}>{m}월</option>
                    )}
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-500 mb-1.5">팀명</label>
                <input type="text" value={teamName}
                  onChange={e => setTeamName(e.target.value)}
                  className="w-full text-sm border border-slate-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-400" />
              </div>
            </div>
          </div>

          {/* Supported files */}
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
            <p className="text-xs font-semibold text-slate-500 mb-3">📋 지원 파일 유형</p>
            {[
              ['오더판매내역', '가장 상세한 분석'],
              ['분야별집계', '카테고리 요약'],
              ['내원경로별', '채널 요약'],
              ['기타 매출 엑셀', '자동 인식'],
            ].map(([t, s]) => (
              <div key={t} className="flex items-start gap-2 py-1.5">
                <span className="w-1.5 h-1.5 rounded-full mt-1.5 shrink-0" style={{ background: 'var(--red)' }} />
                <div>
                  <span className="text-xs font-medium text-slate-700">{t}</span>
                  <span className="text-xs text-slate-400 ml-1">— {s}</span>
                </div>
              </div>
            ))}
          </div>
        </aside>

        {/* ── Main ─────────────────────────────────────── */}
        <main className="flex-1 space-y-4">

          {/* Upload area */}
          <div
            className={`bg-white rounded-2xl shadow-sm border-2 border-dashed transition-all p-10 text-center cursor-pointer
              ${dragging ? 'border-red-400 bg-red-50' : 'border-slate-200 hover:border-red-300 hover:bg-red-50/30'}`}
            onDragOver={e => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            onClick={() => document.getElementById('fileInput')?.click()}
          >
            <input id="fileInput" type="file" accept=".xlsx,.xls" multiple hidden
              onChange={e => addFiles(e.target.files)} />
            <div className="text-5xl mb-3">📁</div>
            <p className="font-bold text-slate-700 text-base">엑셀 파일을 드래그하거나 클릭하여 업로드</p>
            <p className="text-sm text-slate-400 mt-1.5">
              CRM에서 받은 <span className="font-semibold">.xlsx / .xls</span> 파일 (여러 파일 동시 가능)
            </p>
          </div>

          {/* File list */}
          {files.length > 0 && (
            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
              <div className="flex items-center justify-between mb-3">
                <span className="text-sm font-bold text-slate-700">
                  업로드된 파일
                  <span className="ml-1.5 text-xs font-normal text-white px-2 py-0.5 rounded-full" style={{ background: 'var(--red)' }}>
                    {files.length}개
                  </span>
                </span>
                <button onClick={() => { setFiles([]); setDone(false); }}
                  className="text-xs text-slate-400 hover:text-red-500 transition-colors">전체 삭제</button>
              </div>
              <div className="space-y-2">
                {files.map((f, i) => (
                  <div key={i} className="flex items-center justify-between rounded-xl px-4 py-2.5"
                    style={{ background: '#FFF5F5' }}>
                    <div className="flex items-center gap-2">
                      <span className="text-base">📄</span>
                      <span className="text-sm font-medium text-slate-700">{f.name}</span>
                      <span className="text-xs text-slate-400">{(f.size / 1024).toFixed(0)}KB</span>
                    </div>
                    <button onClick={() => setFiles(prev => prev.filter((_, j) => j !== i))}
                      className="text-xs text-slate-300 hover:text-red-500 transition-colors ml-2 font-bold">✕</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Error */}
          {error && (
            <div className="rounded-xl p-4 text-sm border" style={{ background: '#FFF0EF', borderColor: '#FECACA', color: 'var(--red-dk)' }}>
              ⚠️ {error}
            </div>
          )}

          {/* Success */}
          {done && (
            <div className="bg-green-50 border border-green-200 rounded-xl p-4 text-sm text-green-700 font-medium">
              ✅ 보고서가 다운로드되었습니다!
            </div>
          )}

          {/* Generate button */}
          <button
            onClick={handleGenerate}
            disabled={loading || !files.length}
            style={
              loading || !files.length
                ? { background: '#D1D5DB' }
                : { background: 'var(--red)' }
            }
            className="w-full py-4 rounded-2xl font-bold text-white text-base transition-all shadow-md disabled:cursor-not-allowed hover:opacity-90 active:scale-[0.99]"
          >
            {loading ? (
              <span className="flex items-center justify-center gap-2">
                <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24" fill="none">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                </svg>
                분析 중... (30초~1분 소요)
              </span>
            ) : '🚀  PPT 보고서 생성'}
          </button>

          {/* Guide (shown when no files) */}
          {!files.length && (
            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6">
              <p className="font-bold text-sm text-slate-700 mb-4">💡 사용 방법</p>
              <div className="space-y-3">
                {[
                  '왼쪽에서 병원, 기간, 팀명을 설정하세요',
                  'CRM에서 받은 엑셀 파일을 위 영역에 업로드하세요',
                  '보고서 생성 버튼을 클릭하세요',
                  'PPT 파일이 자동으로 다운로드됩니다',
                ].map((t, i) => (
                  <div key={i} className="flex items-start gap-3">
                    <span className="w-6 h-6 rounded-full text-xs font-bold text-white flex items-center justify-center shrink-0 mt-0.5"
                      style={{ background: 'var(--red)' }}>
                      {i + 1}
                    </span>
                    <span className="text-sm text-slate-600">{t}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </main>
      </div>

      {/* ── Footer ──────────────────────────────────────── */}
      <footer className="text-center py-5 text-xs text-slate-400 mt-4 border-t border-slate-100">
        <div className="flex items-center justify-center gap-2">
          <Image src="/logo_passion.png" alt="" width={16} height={16} className="opacity-60" />
          <span>열정의시간 마케팅팀 · ceo@timeofpassion.com</span>
        </div>
      </footer>
    </div>
  );
}
