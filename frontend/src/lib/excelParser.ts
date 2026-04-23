import * as XLSX from 'xlsx';

export interface OrderRow {
  chart_no: string;
  date: Date | null;
  order_name: string;
  amount: number;
  channel: string;
  doctor: string;
  nationality: string;
  original_category: string;
}

// ── 컬럼 alias 매핑 (최대한 포괄적으로) ───────────────────────────────────────
const ALIASES: Record<string, string[]> = {
  chart_no: [
    '차트번호', '차트 번호', '차트no', '차트No', '차트NO',
    '환자번호', '환자 번호', '환자id', '환자ID', '등록번호',
    'chart_no', 'chartno', 'chart no', 'pt.no', 'pt no', 'ptno',
    'patient_id', 'patientid', 'id', 'no', '번호',
  ],
  date: [
    '진료일', '진료 일', '날짜', '일자', '방문일', '시술일', '접수일',
    '내원일', '결제일', '처리일', '발생일', '처방일', '수납일',
    '진료날짜', '서비스일', '예약일', 'date', '진료일자',
  ],
  order_name: [
    '오더명', '오더 명', '시술명', '항목명', '항목', '품목명',
    '처방명', '처방', '처치명', '처치', '진료항목', '서비스명',
    '시술', '내용', '제품명', '서비스', '진료내용', '수납항목',
    '행위명', '행위', '처방내용', 'order_name', 'item', 'service',
    '의료행위', '진료', '항목명칭',
  ],
  amount: [
    '매출금액', '매출 금액', '금액', '결제금액', '매출액', '판매금액',
    '수납금액', '청구금액', '실매출', '합계금액', '총금액', '총매출',
    '실결제금액', '최종금액', '순매출', '단가', '가격', '비용',
    '매출', '수익', 'amount', 'price', 'total', 'revenue', 'fee',
    '처방금액', '진료비', '수납', '결제액',
  ],
  channel: [
    '내원경로', '내원 경로', '유입경로', '채널', '매체',
    '광고매체', '경로', '소개경로', '내원방법', '유입', '채널명',
    '마케팅채널', '유입채널', '소스', 'channel', 'source', 'medium',
    '마케팅경로', '방문경로', '유입소스',
  ],
  doctor: [
    '진료의', '담당의', '의사', '원장', '담당의사',
    '의료진', '시술의', '진료의사', '담당원장', '의사명',
    '원장명', 'doctor', 'physician', '처방의',
  ],
  nationality: [
    '국적', '국가', '외국인여부', '외국인', '국적코드',
    'nationality', 'country', '내외국인',
  ],
  original_category: [
    '분야', '카테고리', '분류', '시술분류',
    '진료과목', '부서', '과목', '항목분류', '시술종류',
    'category', 'type', '구분', '종류', '영역',
  ],
};

// ── 유틸리티 ──────────────────────────────────────────────────────────────────

function normalizeKey(s: string): string {
  return String(s ?? '').trim().toLowerCase().replace(/\s+/g, '');
}

function findColumn(headers: string[], target: string): string | null {
  const aliases = (ALIASES[target] || []).map(normalizeKey);
  for (const h of headers) {
    const norm = normalizeKey(h);
    if (aliases.includes(norm)) return h;
    if (aliases.some(a => norm.includes(a) || a.includes(norm))) return h;
  }
  return null;
}

function cleanAmount(v: unknown): number {
  if (typeof v === 'number') return isFinite(v) && v > 0 ? v : 0;
  const s = String(v ?? '')
    .replace(/,/g, '')
    .replace(/원/g, '')
    .replace(/₩/g, '')
    .replace(/\$/g, '')
    .replace(/\s/g, '')
    .trim();
  // 음수 (괄호 표현) 제외
  if (s.startsWith('(') && s.endsWith(')')) return 0;
  if (s.startsWith('-')) return 0;
  const n = parseFloat(s);
  return isNaN(n) || n <= 0 ? 0 : n;
}

function parseDate(v: unknown): Date | null {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  // Excel 시리얼 숫자 처리
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  const s = String(v).trim();
  // YYYYMMDD 형식
  if (/^\d{8}$/.test(s)) {
    return new Date(`${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`);
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ── 헤더 행 자동 탐지 (제목행·날짜행 건너뛰기) ────────────────────────────────

function findHeaderRowIndex(rawRows: unknown[][]): number {
  const maxScan = Math.min(15, rawRows.length);
  let bestIdx = 0, bestScore = 0;

  for (let i = 0; i < maxScan; i++) {
    const row = rawRows[i];
    if (!Array.isArray(row) || row.length < 2) continue;
    const rowStr = row.map(v => normalizeKey(String(v ?? ''))).join(' ');
    let score = 0;
    for (const aliases of Object.values(ALIASES)) {
      if (aliases.map(normalizeKey).some(a => rowStr.includes(a))) score++;
    }
    if (score > bestScore) { bestScore = score; bestIdx = i; }
  }
  return bestIdx;
}

// ── 시트 타입 감지 (스코어 기반) ──────────────────────────────────────────────

type SheetType = '오더판매내역' | '분야별집계' | '내원경로별' | '범용';

function detectType(headers: string[]): SheetType {
  const hs = headers.map(normalizeKey).join(' ');
  const has = (field: string) =>
    (ALIASES[field] || []).map(normalizeKey).some(a => hs.includes(a));

  const orderScore =
    (has('order_name') ? 3 : 0) +
    (has('chart_no')   ? 2 : 0) +
    (has('amount')     ? 1 : 0) +
    (has('date')       ? 1 : 0);

  const catScore =
    (has('original_category') ? 3 : 0) +
    (has('amount')            ? 1 : 0);

  const chScore =
    (has('channel') ? 3 : 0) +
    (has('amount')  ? 1 : 0);

  if (orderScore >= 3) return '오더판매내역';
  if (catScore  >= 3) return '분야별집계';
  if (chScore   >= 3) return '내원경로별';
  // 차트번호+금액만 있어도 오더로 시도
  if (has('chart_no') && has('amount')) return '오더판매내역';
  // 금액 컬럼이라도 있으면 범용으로 처리
  if (has('amount')) return '범용';
  return '범용';
}

// ── 범용 파서 — 컬럼명 모르는 병원 포맷 대응 ─────────────────────────────────

function parseFallback(
  headers: string[],
  rows: Record<string, unknown>[]
): OrderRow[] {
  // 숫자값이 큰 컬럼 중 금액 컬럼 자동 추정
  const numericCols = headers.filter(h => {
    const vals = rows.slice(0, 20).map(r => cleanAmount(r[h]));
    const positives = vals.filter(v => v > 1000);
    return positives.length >= 3;
  });
  if (!numericCols.length) return [];

  // 금액 컬럼: 가장 큰 평균값을 가진 숫자 컬럼
  const amtCol = numericCols.sort((a, b) => {
    const avgA = rows.slice(0, 20).reduce((s, r) => s + cleanAmount(r[a]), 0);
    const avgB = rows.slice(0, 20).reduce((s, r) => s + cleanAmount(r[b]), 0);
    return avgB - avgA;
  })[0];

  // 텍스트 컬럼 중 가장 다양한 값 = 시술명 후보
  const textCols = headers.filter(h => h !== amtCol).filter(h => {
    const vals = rows.slice(0, 20).map(r => String(r[h] ?? '').trim()).filter(Boolean);
    const unique = new Set(vals);
    return unique.size >= 3 && vals.every(v => v.length < 50);
  });
  const nameCol = textCols[0] ?? null;

  // 날짜처럼 보이는 컬럼
  const dateCol = headers.find(h => {
    const vals = rows.slice(0, 10).map(r => parseDate(r[h]));
    return vals.filter(Boolean).length >= 3;
  }) ?? null;

  return rows
    .map(row => ({
      chart_no:          '',
      date:              dateCol ? parseDate(row[dateCol]) : null,
      order_name:        nameCol ? String(row[nameCol] ?? '').trim() : '(항목 미상)',
      amount:            cleanAmount(row[amtCol]),
      channel:           '미기재',
      doctor:            '미기재',
      nationality:       '미기재',
      original_category: '',
    } as OrderRow))
    .filter(r => r.amount > 0);
}

// ── ParsedData + 파싱 메타 ────────────────────────────────────────────────────

export interface ParsedData {
  orders: OrderRow[];
  categorySummary: { category: string; amount: number }[];
  channelSummary: { channel: string; amount: number }[];
  parseInfo: {
    filesRead: number;
    sheetsRead: number;
    orderRows: number;
    missingCols: string[];  // 못 찾은 컬럼
    warnings: string[];
  };
}

// ── 메인 파서 ─────────────────────────────────────────────────────────────────

export async function parseExcelFiles(files: File[]): Promise<ParsedData> {
  const result: ParsedData = {
    orders: [],
    categorySummary: [],
    channelSummary: [],
    parseInfo: { filesRead: 0, sheetsRead: 0, orderRows: 0, missingCols: [], warnings: [] },
  };

  for (const file of files) {
    result.parseInfo.filesRead++;
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      // raw 2D 배열로 먼저 읽어 헤더 행 탐지
      const rawData = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null }) as unknown[][];
      if (rawData.length < 2) continue;

      const headerRowIdx = findHeaderRowIndex(rawData);
      const headerRow    = rawData[headerRowIdx] as unknown[];
      const headers      = headerRow.map(v => String(v ?? '').trim()).filter(Boolean);
      if (!headers.length) continue;

      // headerRowIdx부터 다시 json으로 파싱
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
        defval: null,
        range: headerRowIdx,
      });
      if (!rows.length) continue;

      result.parseInfo.sheetsRead++;
      const type = detectType(headers);

      if (type === '오더판매내역') {
        const colMap: Record<string, string | null> = {};
        for (const target of Object.keys(ALIASES)) {
          colMap[target] = findColumn(headers, target);
        }

        // 못 찾은 중요 컬럼 기록
        const criticalCols = ['order_name', 'amount'];
        for (const col of criticalCols) {
          if (!colMap[col]) result.parseInfo.missingCols.push(col);
        }

        const amtCol = colMap.amount;
        if (!amtCol) {
          result.parseInfo.warnings.push(`[${sheetName}] 금액 컬럼을 찾을 수 없어 건너뜀`);
          continue;
        }

        const parsed = rows
          .map(row => ({
            chart_no:          String(row[colMap.chart_no  ?? ''] ?? '').trim(),
            date:              parseDate(row[colMap.date       ?? '']),
            order_name:        String(row[colMap.order_name ?? ''] ?? '').trim() || '(미상)',
            amount:            cleanAmount(row[amtCol]),
            channel:           String(row[colMap.channel     ?? ''] ?? '미기재').trim() || '미기재',
            doctor:            String(row[colMap.doctor      ?? ''] ?? '미기재').trim() || '미기재',
            nationality:       String(row[colMap.nationality ?? ''] ?? '미기재').trim() || '미기재',
            original_category: String(row[colMap.original_category ?? ''] ?? '').trim(),
          } as OrderRow))
          .filter(r => r.amount > 0);

        result.orders.push(...parsed);
        result.parseInfo.orderRows += parsed.length;
        continue;
      }

      if (type === '분야별집계') {
        const catCol = findColumn(headers, 'original_category');
        const amtCol = findColumn(headers, 'amount');
        if (catCol && amtCol) {
          for (const row of rows) {
            const cat = String(row[catCol] ?? '').trim();
            const amt = cleanAmount(row[amtCol]);
            if (cat && amt > 0) result.categorySummary.push({ category: cat, amount: amt });
          }
        }
        continue;
      }

      if (type === '내원경로별') {
        const chCol  = findColumn(headers, 'channel') ?? findColumn(headers, 'original_category');
        const amtCol = findColumn(headers, 'amount');
        if (chCol && amtCol) {
          for (const row of rows) {
            const ch  = String(row[chCol] ?? '').trim();
            const amt = cleanAmount(row[amtCol]);
            if (ch && amt > 0) result.channelSummary.push({ channel: ch, amount: amt });
          }
        }
        continue;
      }

      // 범용 fallback — 어떤 포맷이든 최대한 데이터 추출
      if (type === '범용') {
        const fallback = parseFallback(headers, rows);
        if (fallback.length > 0) {
          result.orders.push(...fallback);
          result.parseInfo.orderRows += fallback.length;
          result.parseInfo.warnings.push(
            `[${sheetName}] 표준 포맷 아님 → 범용 파서로 ${fallback.length}행 추출`
          );
        }
      }
    }
  }

  return result;
}
