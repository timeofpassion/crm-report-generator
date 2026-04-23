import { OrderRow, ParsedData } from './excelParser';
import { HospitalConfig, categorize } from './hospitalConfig';

export interface KPIs {
  totalRevenue: number;
  totalVisits: number;
  uniquePatients: number;
  avgPerPatient: number;
  avgPerVisit: number;
}

export interface CategorySales {
  category: string; revenue: number; count: number; pct: number;
}
export interface ProcedureSales {
  name: string; revenue: number; count: number; pct: number; category: string;
}
export interface DoctorSales {
  doctor: string; category: string; revenue: number; count: number;
}
export interface ChannelSales {
  channel: string; revenue: number; count: number; avg: number;
}
export interface NationalitySales {
  nationality: string; revenue: number; count: number; avg: number;
}
export interface CrossSelling {
  crossSellRate: number;
  avgMultiplier: number;
  topCombinations: { combo: string; count: number }[];
}
export interface DailyTrend { date: string; revenue: number; count: number; }
export interface WeekdayTrend { name: string; revenue: number; count: number; avg: number; }
export interface PatientSpend { chartNo: string; total: number; visits: number; }

export interface AnalysisResult {
  kpis: KPIs;
  categorySales: CategorySales[];
  topProcedures: ProcedureSales[];
  doctorSales: DoctorSales[];
  channelSales: ChannelSales[];
  nationalitySales: NationalitySales[];
  crossSelling: CrossSelling | null;
  dailyTrends: DailyTrend[];
  weekdayTrends: WeekdayTrend[];
  patientSpending: PatientSpend[];
  keyFindings: string[];
  insights: { doctor: string[]; marketing: string[]; operations: string[] };
}

const WEEKDAY = ['월', '화', '수', '목', '금', '토', '일'];

function groupBy<T>(arr: T[], key: (x: T) => string): Map<string, T[]> {
  const m = new Map<string, T[]>();
  for (const x of arr) {
    const k = key(x);
    if (!m.has(k)) m.set(k, []);
    m.get(k)!.push(x);
  }
  return m;
}

function sumBy<T>(arr: T[], fn: (x: T) => number) {
  return arr.reduce((s, x) => s + fn(x), 0);
}

function fmtKrw(v: number): string {
  if (v >= 1e8) return `${(v / 1e8).toFixed(1)}억원`;
  if (v >= 1e4) return `${Math.round(v / 1e4)}만원`;
  return `${v.toLocaleString()}원`;
}

export class DataAnalyzer {
  constructor(private cfg: HospitalConfig) {}

  analyze(data: ParsedData): AnalysisResult {
    const orders = data.orders;
    const result = {} as AnalysisResult;

    // Categorize
    const tagged = orders.map(r => ({
      ...r,
      category: categorize(r.order_name, this.cfg.categoryKeywords),
    }));

    // KPIs
    if (tagged.length) {
      const totalRevenue   = sumBy(tagged, r => r.amount);
      const totalVisits    = tagged.length;
      const patients       = new Set(tagged.filter(r => r.chart_no).map(r => r.chart_no)).size || totalVisits;
      result.kpis = {
        totalRevenue,
        totalVisits,
        uniquePatients: patients,
        avgPerPatient:  totalRevenue / patients,
        avgPerVisit:    totalRevenue / totalVisits,
      };
    } else {
      result.kpis = { totalRevenue: 0, totalVisits: 0, uniquePatients: 0, avgPerPatient: 0, avgPerVisit: 0 };
    }

    // Category sales
    const catGroups = groupBy(tagged, r => r.category);
    const totalRev = result.kpis.totalRevenue || 1;
    result.categorySales = Array.from(catGroups.entries())
      .map(([category, rows]) => ({
        category, revenue: sumBy(rows, r => r.amount), count: rows.length,
        pct: (sumBy(rows, r => r.amount) / totalRev) * 100,
      }))
      .sort((a, b) => b.revenue - a.revenue);

    // Top individual procedures (시술별 TOP 순위)
    const procGroups = groupBy(tagged.filter(r => r.order_name), r => r.order_name);
    result.topProcedures = Array.from(procGroups.entries())
      .map(([name, rows]) => {
        const revenue = sumBy(rows, r => r.amount);
        return {
          name,
          revenue,
          count: rows.length,
          pct: (revenue / totalRev) * 100,
          category: rows[0].category,
        };
      })
      .sort((a, b) => b.revenue - a.revenue)
      .slice(0, 20);

    // Doctor sales
    const docRows = tagged.filter(r => r.doctor && r.doctor !== '미기재');
    if (docRows.length) {
      const pairs = groupBy(docRows, r => `${r.doctor}___${r.category}`);
      result.doctorSales = Array.from(pairs.entries())
        .map(([key, rows]) => {
          const [doctor, category] = key.split('___');
          return { doctor, category, revenue: sumBy(rows, r => r.amount), count: rows.length };
        })
        .sort((a, b) => b.revenue - a.revenue);
    } else {
      result.doctorSales = [];
    }

    // Channel sales
    const chRows = tagged.filter(r => r.channel && r.channel !== '미기재');
    if (chRows.length) {
      const chGroups = groupBy(chRows, r => r.channel);
      result.channelSales = Array.from(chGroups.entries())
        .map(([channel, rows]) => {
          const revenue = sumBy(rows, r => r.amount);
          return { channel, revenue, count: rows.length, avg: revenue / rows.length };
        })
        .sort((a, b) => b.revenue - a.revenue);
    } else if (data.channelSummary.length) {
      result.channelSales = data.channelSummary.map(r => ({
        channel: r.channel, revenue: r.amount, count: 0, avg: 0,
      })).sort((a, b) => b.revenue - a.revenue);
    } else {
      result.channelSales = [];
    }

    // Nationality
    const natRows = tagged.filter(r => r.nationality && r.nationality !== '미기재');
    if (natRows.length) {
      const natGroups = groupBy(natRows, r => r.nationality);
      result.nationalitySales = Array.from(natGroups.entries())
        .map(([nationality, rows]) => {
          const revenue = sumBy(rows, r => r.amount);
          return { nationality, revenue, count: rows.length, avg: revenue / rows.length };
        })
        .sort((a, b) => b.revenue - a.revenue);
    } else {
      result.nationalitySales = [];
    }

    // Patient spending
    const ptGroups = groupBy(tagged.filter(r => r.chart_no), r => r.chart_no);
    result.patientSpending = Array.from(ptGroups.entries())
      .map(([chartNo, rows]) => ({
        chartNo, total: sumBy(rows, r => r.amount), visits: rows.length,
      }))
      .sort((a, b) => b.total - a.total);

    // Cross-selling
    if (ptGroups.size > 0) {
      const ptCatCounts = Array.from(ptGroups.entries()).map(([, rows]) => {
        const cats = new Set(rows.map(r => r.category));
        return { total: sumBy(rows, r => r.amount), catCount: cats.size, cats: [...cats].sort() };
      });
      const multi = ptCatCounts.filter(p => p.catCount >= 2);
      const single = ptCatCounts.filter(p => p.catCount === 1);
      const crossRate = ptGroups.size > 0 ? (multi.length / ptGroups.size) * 100 : 0;
      const singleAvg = single.length ? sumBy(single, p => p.total) / single.length : 1;
      const multiAvg  = multi.length  ? sumBy(multi,  p => p.total) / multi.length  : 0;

      const comboMap = new Map<string, number>();
      for (const p of multi) {
        const key = p.cats.join(' + ');
        comboMap.set(key, (comboMap.get(key) || 0) + 1);
      }
      const topCombos = [...comboMap.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(([combo, count]) => ({ combo, count }));

      result.crossSelling = {
        crossSellRate: crossRate,
        avgMultiplier: singleAvg > 0 ? multiAvg / singleAvg : 0,
        topCombinations: topCombos,
      };
    } else {
      result.crossSelling = null;
    }

    // Daily trends
    const datedRows = tagged.filter(r => r.date instanceof Date);
    if (datedRows.length) {
      const dayGroups = groupBy(datedRows, r => (r.date as Date).toISOString().slice(0, 10));
      result.dailyTrends = [...dayGroups.entries()]
        .map(([date, rows]) => ({ date, revenue: sumBy(rows, r => r.amount), count: rows.length }))
        .sort((a, b) => a.date.localeCompare(b.date));

      const wdGroups = groupBy(datedRows, r => String((r.date as Date).getDay()));
      result.weekdayTrends = Array.from({ length: 7 }, (_, i) => {
        const day = (i + 1) % 7;
        const rows = wdGroups.get(String(day)) || [];
        const revenue = sumBy(rows, r => r.amount);
        return { name: WEEKDAY[i], revenue, count: rows.length, avg: rows.length ? revenue / rows.length : 0 };
      });
    } else {
      result.dailyTrends = [];
      result.weekdayTrends = [];
    }

    result.keyFindings = this.generateFindings(result);
    result.insights     = this.generateInsights(result);
    return result;
  }

  private generateFindings(r: AnalysisResult): string[] {
    const findings: string[] = [];
    if (r.categorySales.length) {
      const top = r.categorySales[0];
      findings.push(`최고 매출 카테고리: ${top.category} (${top.pct.toFixed(1)}% · ${fmtKrw(top.revenue)})`);
    }
    if (r.topProcedures.length) {
      const top = r.topProcedures[0];
      findings.push(`매출 1위 시술: ${top.name} (${fmtKrw(top.revenue)} · ${top.count}건)`);
    }
    if (r.channelSales.length) {
      const top = r.channelSales[0];
      findings.push(`주요 유입경로: ${top.channel} (방문 ${top.count.toLocaleString()}건)`);
    }
    if (r.crossSelling) {
      const cs = r.crossSelling;
      findings.push(`복합시술 비율: ${cs.crossSellRate.toFixed(1)}% · 단일 대비 ${cs.avgMultiplier.toFixed(1)}배 객단가`);
    }
    findings.push(`평균 객단가: ${fmtKrw(r.kpis.avgPerPatient)} (총 ${r.kpis.uniquePatients.toLocaleString()}명)`);
    return findings;
  }

  private generateInsights(r: AnalysisResult): AnalysisResult['insights'] {
    const ins = { doctor: [] as string[], marketing: [] as string[], operations: [] as string[] };

    if (r.categorySales.length >= 2) {
      const top2 = r.categorySales.slice(0, 2).map(c => c.category).join(' / ');
      ins.doctor.push(`핵심 매출 시술 ${top2} — 집중 역량 강화 권고`);
    }
    if (r.topProcedures.length) {
      ins.doctor.push(`매출 1위 시술 '${r.topProcedures[0].name}' 수요 강세 지속 — 예약 슬롯 확대 검토`);
    }
    if (r.patientSpending.length >= 10) {
      const p90 = r.patientSpending[Math.floor(r.patientSpending.length * 0.1)].total;
      ins.doctor.push(`상위 10% 객단가 ${fmtKrw(p90)} — VIP 관리 프로그램 도입 검토`);
    }

    if (r.channelSales.length >= 1) {
      const topVol = r.channelSales[0];
      ins.marketing.push(`유입량 1위 '${topVol.channel}' (${topVol.count}건) — 예산 유지 및 광고 소재 리프레시 권장`);
    }
    if (r.channelSales.length >= 2) {
      const byEff = [...r.channelSales].filter(c => c.count >= 3).sort((a, b) => b.avg - a.avg);
      if (byEff.length) {
        ins.marketing.push(`건당 매출 최고 채널 '${byEff[0].channel}' (${fmtKrw(byEff[0].avg)}) — 투자 확대 우선 고려`);
      }
    }
    if (r.crossSelling && r.crossSelling.crossSellRate < 30) {
      ins.marketing.push(`복합시술 비율 ${r.crossSelling.crossSellRate.toFixed(1)}% — 패키지 프로모션으로 객단가 제고 가능`);
    }

    if (r.weekdayTrends.length) {
      const best  = r.weekdayTrends.reduce((a, b) => a.revenue > b.revenue ? a : b);
      const worst = r.weekdayTrends.filter(d => d.count > 0).reduce((a, b) => a.revenue < b.revenue ? a : b);
      ins.operations.push(`최다 매출 ${best.name}요일 — 스태프 배치 강화 및 예약 시스템 최적화`);
      if (worst.name !== best.name) {
        ins.operations.push(`${worst.name}요일 매출 저조 — 한정 이벤트 기획으로 방문 유도`);
      }
    }
    if (r.kpis.uniquePatients > 0) {
      const revisitRate = r.patientSpending.filter(p => p.visits >= 2).length / r.kpis.uniquePatients * 100;
      if (revisitRate > 0) {
        ins.operations.push(`재방문 환자 비율 ${revisitRate.toFixed(1)}% — 리텐션 프로그램 강화 권장`);
      }
    }
    return ins;
  }
}
