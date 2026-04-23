'use client';

import { useState, useCallback, DragEvent } from 'react';
import { HOSPITAL_CONFIGS } from '@/lib/hospitalConfig';

const HOSPITALS = Object.entries(HOSPITAL_CONFIGS).map(([k, v]) => ({ key: k, name: v.name }));

export default function Home() {
  const now = new Date();
  const [files, setFiles] = useState<File[]>([]);
  const [dragging, setDragging] = useState(false);
  const [hospitalKey, setHospitalKey] = useState('mellow_sinsa');
  const [hospitalName, setHospitalName] = useState(HOSPITAL_CONFIGS.mellow_sinsa.name);
  const [year, setYear]   = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const [teamName, setTeamName] = useState('열정의시간 마케팅팀');
  const [loading, setLoading]   = useState(false);
  const [error, setError]       = useState('');
  const [done, setDone]         = useState(false);

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
    e.preventDefault(); setDragging(false);
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
      a.download = `${hospitalName}_${year}년${month}월_매출분석보고서.pptx`;
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
    <div className="min-h-screen bg-slate-50">
      {/* Header */}
      <header style={{ background: 'var(--navy)' }} className="text-white shadow-lg">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center gap-3">
          <span className="text-2xl">📊</span>
          <div>
            <h1 className="text-lg font-bold leading-tight">CRM 매출분석 보고서 자동 생성기</h1>
            <p className="text-xs text-blue-200">열정의시간 마케팅팀</p>
          </div>
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-6 py-8 flex gap-6">
        {/* Sidebar */}
        <aside className="w-72 shrink-0">
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 space-y-5">
            <h2 className="font-bold text-sm text-slate-700 flex items-center gap-1">
              ⚙️ 보고서 설정
            </h2>

            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">병원 선택</label>
              <select
                value={hospitalKey}
                onChange={e => handleHospitalChange(e.target.value)}
                className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2"
                style={{ '--tw-ring-color': 'var(--blue)' } as React.CSSProperties}
              >
                {HOSPITALS.map(h => <option key={h.key} value={h.key}>{h.name}</option>)}
              </select>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">병원명 (표지)</label>
              <input
                type="text" value={hospitalName}
                onChange={e => setHospitalName(e.target.value)}
                className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2"
              />
            </div>

            <div className="flex gap-2">
              <div className="flex-1">
                <label className="block text-xs font-medium text-slate-600 mb-1">년도</label>
                <select value={year} onChange={e => setYear(+e.target.value)}
                  className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2">
                  {yearOptions.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
              <div className="flex-1">
                <label className="block text-xs font-medium text-slate-600 mb-1">월</label>
                <select value={month} onChange={e => setMonth(+e.target.value)}
                  className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2">
                  {Array.from({ length: 12 }, (_, i) => i + 1).map(m =>
                    <option key={m} value={m}>{m}월</option>
                  )}
                </select>
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">팀명</label>
              <input type="text" value={teamName}
                onChange={e => setTeamName(e.target.value)}
                className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2" />
            </div>
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 mt-4">
            <h3 className="font-bold text-xs text-slate-600 mb-3">📋 지원 파일 유형</h3>
            {['오더판매내역 (가장 상세)', '분야별집계', '내원경로별', '기타 매출 엑셀'].map(t => (
              <div key={t} className="flex items-center gap-2 text-xs text-slate-600 py-1">
                <span className="w-1.5 h-1.5 rounded-full bg-blue-400 shrink-0" />
                {t}
              </div>
            ))}
          </div>
        </aside>

        {/* Main */}
        <main className="flex-1 space-y-5">
          {/* Upload area */}
          <div
            className={`bg-white rounded-xl shadow-sm border-2 border-dashed transition-colors p-8 text-center cursor-pointer
              ${dragging ? 'border-blue-400 bg-blue-50' : 'border-slate-300 hover:border-blue-300'}`}
            onDragOver={e => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            onClick={() => document.getElementById('fileInput')?.click()}
          >
            <input id="fileInput" type="file" accept=".xlsx,.xls" multiple hidden
              onChange={e => addFiles(e.target.files)} />
            <div className="text-4xl mb-3">📁</div>
            <p className="font-semibold text-slate-700">엑셀 파일을 드래그하거나 클릭하여 업로드</p>
            <p className="text-sm text-slate-500 mt-1">CRM에서 받은 .xlsx / .xls 파일 (여러 파일 동시 가능)</p>
          </div>

          {/* File list */}
          {files.length > 0 && (
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
              <div className="flex items-center justify-between mb-3">
                <h3 className="font-semibold text-sm text-slate-700">업로드된 파일 ({files.length}개)</h3>
                <button onClick={() => { setFiles([]); setDone(false); }}
                  className="text-xs text-slate-400 hover:text-red-500">전체 삭제</button>
              </div>
              <div className="space-y-2">
                {files.map((f, i) => (
                  <div key={i} className="flex items-center justify-between bg-slate-50 rounded-lg px-3 py-2">
                    <div className="flex items-center gap-2">
                      <span className="text-green-500 text-sm">📄</span>
                      <span className="text-sm text-slate-700">{f.name}</span>
                      <span className="text-xs text-slate-400">{(f.size / 1024).toFixed(0)}KB</span>
                    </div>
                    <button onClick={() => setFiles(prev => prev.filter((_, j) => j !== i))}
                      className="text-xs text-slate-400 hover:text-red-500 ml-2">✕</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Error */}
          {error && (
            <div className="bg-red-50 border border-red-200 rounded-xl p-4 text-sm text-red-700">
              ⚠️ {error}
            </div>
          )}

          {/* Success */}
          {done && (
            <div className="bg-green-50 border border-green-200 rounded-xl p-4 text-sm text-green-700">
              ✅ 보고서가 다운로드되었습니다!
            </div>
          )}

          {/* Generate button */}
          <button
            onClick={handleGenerate}
            disabled={loading || !files.length}
            style={{ background: loading || !files.length ? undefined : 'var(--navy)' }}
            className={`w-full py-4 rounded-xl font-bold text-white text-base transition-all
              ${loading || !files.length
                ? 'bg-slate-300 cursor-not-allowed'
                : 'hover:opacity-90 active:scale-[0.99] shadow-lg'
              }`}
          >
            {loading ? (
              <span className="flex items-center justify-center gap-2">
                <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                </svg>
                분석 중... (30초~1분 소요)
              </span>
            ) : '🚀 PPT 보고서 생성'}
          </button>

          {/* Guide */}
          {!files.length && (
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
              <h3 className="font-bold text-sm text-slate-700 mb-3">💡 사용 방법</h3>
              {[
                ['1', '왼쪽에서 병원, 기간, 팀명을 설정하세요'],
                ['2', 'CRM에서 받은 엑셀 파일을 업로드하세요'],
                ['3', '보고서 생성 버튼을 클릭하세요'],
                ['4', 'PPT 파일이 자동으로 다운로드됩니다'],
              ].map(([n, t]) => (
                <div key={n} className="flex items-start gap-3 py-2">
                  <span className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold text-white shrink-0"
                    style={{ background: 'var(--blue)' }}>{n}</span>
                  <span className="text-sm text-slate-600">{t}</span>
                </div>
              ))}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
