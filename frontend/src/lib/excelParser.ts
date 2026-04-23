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

const ALIASES: Record<string, string[]> = {
  chart_no:          ['차트번호', '차트 번호', '환자번호', '환자번호', '차트'],
  date:              ['진료일', '진료 일', '날짜', '일자', '방문일', '시술일', '접수일'],
  order_name:        ['오더명', '오더 명', '시술명', '항목명', '항목', '품목명', '처방명'],
  amount:            ['매출금액', '매출 금액', '금액', '결제금액', '매출액', '판매금액'],
  channel:           ['내원경로', '내원 경로', '유입경로', '채널', '매체'],
  doctor:            ['진료의', '담당의', '의사', '원장', '담당의사'],
  nationality:       ['국적', '국가'],
  original_category: ['분야', '카테고리', '분류', '시술분류'],
};

function findColumn(headers: string[], target: string): string | null {
  const aliases = ALIASES[target] || [];
  for (const h of headers) {
    const clean = String(h).trim();
    if (aliases.includes(clean)) return h;
    if (aliases.some(a => clean.includes(a))) return h;
  }
  return null;
}

function cleanAmount(v: unknown): number {
  if (typeof v === 'number') return v;
  const s = String(v ?? '').replace(/,/g, '').replace(/원/g, '').trim();
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function parseDate(v: unknown): Date | null {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function detectType(headers: string[]): string {
  const hs = headers.join(' ');
  const hasOrder  = ALIASES.order_name.some(a => hs.includes(a));
  const hasAmount = ALIASES.amount.some(a => hs.includes(a));
  const hasChart  = ALIASES.chart_no.some(a => hs.includes(a));
  const hasCat    = ALIASES.original_category.some(a => hs.includes(a));
  const hasCh     = ALIASES.channel.some(a => hs.includes(a));
  if (hasOrder && hasAmount && hasChart) return '오더판매내역';
  if (hasCat && hasAmount) return '분야별집계';
  if (hasCh && hasAmount) return '내원경로별';
  return '기타';
}

export interface ParsedData {
  orders: OrderRow[];
  categorySummary: { category: string; amount: number }[];
  channelSummary: { channel: string; amount: number }[];
}

export async function parseExcelFiles(files: File[]): Promise<ParsedData> {
  const result: ParsedData = { orders: [], categorySummary: [], channelSummary: [] };

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null });
      if (!rows.length) continue;

      const headers = Object.keys(rows[0]);
      const type = detectType(headers);

      if (type === '오더판매내역') {
        const colMap: Record<string, string> = {};
        for (const target of Object.keys(ALIASES)) {
          const found = findColumn(headers, target);
          if (found) colMap[target] = found;
        }

        const parsed = rows
          .map(row => ({
            chart_no:          String(row[colMap.chart_no]  ?? '').trim(),
            date:              parseDate(row[colMap.date]),
            order_name:        String(row[colMap.order_name] ?? ''),
            amount:            cleanAmount(row[colMap.amount]),
            channel:           String(row[colMap.channel]    ?? '미기재'),
            doctor:            String(row[colMap.doctor]     ?? '미기재'),
            nationality:       String(row[colMap.nationality] ?? '미기재'),
            original_category: String(row[colMap.original_category] ?? ''),
          } as OrderRow))
          .filter(r => r.amount > 0);

        result.orders.push(...parsed);
        break;
      }

      if (type === '분야별집계') {
        const catCol = findColumn(headers, 'original_category');
        const amtCol = findColumn(headers, 'amount');
        if (catCol && amtCol) {
          for (const row of rows) {
            const cat = String(row[catCol] ?? '');
            const amt = cleanAmount(row[amtCol]);
            if (cat && amt > 0) result.categorySummary.push({ category: cat, amount: amt });
          }
        }
      }

      if (type === '내원경로별') {
        const chCol  = findColumn(headers, 'channel') || findColumn(headers, 'original_category');
        const amtCol = findColumn(headers, 'amount');
        if (chCol && amtCol) {
          for (const row of rows) {
            const ch  = String(row[chCol] ?? '');
            const amt = cleanAmount(row[amtCol]);
            if (ch && amt > 0) result.channelSummary.push({ channel: ch, amount: amt });
          }
        }
      }
    }
  }

  return result;
}
