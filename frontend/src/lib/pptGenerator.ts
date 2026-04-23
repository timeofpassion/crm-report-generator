import PptxGenJS from 'pptxgenjs';
import { AnalysisResult } from './dataAnalyzer';
import { HospitalConfig } from './hospitalConfig';

// ── 색상 상수 (열정의시간 브랜드) ─────────────────────────────────────────────
const CP  = 'E63329';  // Primary  red
const CS  = '1C5FA6';  // Secondary blue
const CA  = 'F5C200';  // Accent   yellow
const CX  = '5CC8E0';  // Cyan
const CL  = 'FFF0EF';  // Light red bg
const CW  = 'FFFFFF';  // White
const CD  = '1A1A2E';  // Dark text
const CG  = '6B7280';  // Gray
const CG2 = 'F8F8F9';  // Light gray

const CHART_COLORS = [CP, CS, CA, CX, '27AE60', '9B59B6', 'E67E22', '3498DB'];
const FONT         = '맑은 고딕';
const SW = 13.33, SH = 7.5;

export interface ReportConfig {
  hospitalName: string;
  year: number;
  month: number;
  teamName: string;
}

function fmtKrw(v: number): string {
  if (v >= 1e8) return `${(v / 1e8).toFixed(1)}억원`;
  if (v >= 1e4) return `${Math.round(v / 1e4)}만원`;
  return `${v.toLocaleString()}원`;
}

type SlideIface = ReturnType<PptxGenJS['addSlide']>;

function header(slide: SlideIface, title: string, page: number, rc: ReportConfig) {
  slide.addShape('rect', { x: 0, y: 0, w: SW, h: 0.85, fill: { color: CP }, line: { color: CP } });
  slide.addText(title, { x: 0.45, y: 0.1, w: 10, h: 0.65, fontSize: 20, bold: true, color: CW, fontFace: FONT });
  slide.addText(`${page} / 17`, { x: SW - 1.1, y: 0.22, w: 0.9, h: 0.4, fontSize: 10, color: CW, align: 'right', fontFace: FONT });
  const footer = `${rc.hospitalName} | ${rc.year}년 ${rc.month}월 | ${rc.teamName}`;
  slide.addText(footer, { x: 0.4, y: SH - 0.35, w: SW - 0.8, h: 0.3, fontSize: 7, color: CG, fontFace: FONT });
}

function kpiBox(slide: SlideIface, x: number, y: number, w: number, h: number, label: string, value: string, sub = '') {
  slide.addShape('rect', { x, y, w, h, fill: { color: CL }, line: { color: 'D1E8F7' } });
  slide.addText(label, { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 0.28, fontSize: 9, color: CG, fontFace: FONT });
  slide.addText(value, { x: x + 0.1, y: y + 0.38, w: w - 0.2, h: 0.55, fontSize: 18, bold: true, color: CP, fontFace: FONT });
  if (sub) slide.addText(sub, { x: x + 0.15, y: y + 0.9, w: w - 0.3, h: 0.25, fontSize: 8, color: CG, italic: true, fontFace: FONT });
}

function addTable(
  slide: SlideIface,
  headers: string[],
  rows: string[][],
  x: number, y: number, w: number, h: number
) {
  const headerRow = headers.map(h => ({
    text: h,
    options: { bold: true, color: CW, fill: { color: CP }, align: 'center' as const, fontFace: FONT, fontSize: 8 },
  }));
  const dataRows = rows.map((row, ri) => row.map(cell => ({
    text: cell,
    options: { color: CD, fill: { color: ri % 2 === 0 ? CW : CG2 }, align: 'center' as const, fontFace: FONT, fontSize: 8 },
  })));

  slide.addTable([headerRow, ...dataRows], {
    x, y, w, h,
    rowH: h / (rows.length + 1),
    border: { pt: 0.5, color: 'E5E7EB' },
  });
}

// ── 슬라이드별 생성 함수 ─────────────────────────────────────────────────────

function slideCover(pptx: PptxGenJS, rc: ReportConfig) {
  const slide = pptx.addSlide();
  // 배경: 다크
  slide.addShape('rect', { x: 0, y: 0, w: SW, h: SH, fill: { color: '1A1A2E' }, line: { color: '1A1A2E' } });
  // 레드 세로 사이드바
  slide.addShape('rect', { x: 0, y: 0, w: 0.18, h: SH, fill: { color: CP }, line: { color: CP } });
  // 옐로우 하단 라인
  slide.addShape('rect', { x: 0, y: SH - 0.12, w: SW, h: 0.12, fill: { color: CA }, line: { color: CA } });
  // 레드 타이틀 배경
  slide.addShape('rect', { x: 0.18, y: 2.2, w: SW - 0.18, h: 1.1, fill: { color: CP }, line: { color: CP } });
  slide.addText(rc.hospitalName, { x: 0.65, y: 2.28, w: SW - 1.0, h: 0.95, fontSize: 34, bold: true, color: CW, align: 'left', fontFace: FONT });
  slide.addText('매출분석 보고서', { x: 0.65, y: 3.5, w: SW - 1.0, h: 0.75, fontSize: 22, color: 'D0D0D8', align: 'left', fontFace: FONT });
  slide.addText(`${rc.year}년 ${rc.month}월`, { x: 0.65, y: 4.3, w: SW - 1.0, h: 0.6, fontSize: 17, bold: true, color: CA, align: 'left', fontFace: FONT });
  slide.addText(rc.teamName, { x: 0.65, y: SH - 0.55, w: SW - 1.0, h: 0.42, fontSize: 10, color: '7070A0', align: 'left', fontFace: FONT });
}

function slideExecutiveSummary(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, 'Executive Summary', 2, rc);
  const { kpis } = analysis;
  const bw = 2.9, bh = 1.25, gap = 0.15, sx = 0.4;
  kpiBox(slide, sx,               1.0, bw, bh, '총 매출',    fmtKrw(kpis.totalRevenue));
  kpiBox(slide, sx + bw + gap,    1.0, bw, bh, '총 방문건수', `${kpis.totalVisits.toLocaleString()}건`);
  kpiBox(slide, sx + (bw+gap)*2,  1.0, bw, bh, '총 환자수',  `${kpis.uniquePatients.toLocaleString()}명`);
  kpiBox(slide, sx + (bw+gap)*3,  1.0, bw, bh, '평균 객단가', fmtKrw(kpis.avgPerPatient), '환자 1인당');

  slide.addShape('rect', { x: 0.4, y: 2.5, w: SW - 0.8, h: 0.35, fill: { color: CP }, line: { color: CP } });
  slide.addText('핵심 발견 (Key Findings)', { x: 0.55, y: 2.55, w: 6, h: 0.28, fontSize: 11, bold: true, color: CW, fontFace: FONT });

  for (const [i, f] of analysis.keyFindings.slice(0, 4).entries()) {
    const fy = 3.0 + i * 0.72;
    slide.addShape('rect', { x: 0.4, y: fy, w: 0.35, h: 0.35, fill: { color: CS }, line: { color: CS } });
    slide.addText(String(i + 1), { x: 0.44, y: fy + 0.04, w: 0.3, h: 0.3, fontSize: 12, bold: true, color: CW, align: 'center', fontFace: FONT });
    slide.addText(f, { x: 0.85, y: fy + 0.04, w: SW - 1.3, h: 0.6, fontSize: 11, color: CD, fontFace: FONT });
  }
}

function slideCategorySales(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '시술 카테고리별 매출 분석', 3, rc);
  const { categorySales } = analysis;
  if (!categorySales.length) return;

  slide.addChart('doughnut', [{
    name: '카테고리별 매출',
    labels: categorySales.map(r => r.category),
    values: categorySales.map(r => r.revenue),
  }], {
    x: 0.3, y: 0.9, w: 6.5, h: 5.5,
    chartColors: CHART_COLORS.slice(0, categorySales.length),
    showLegend: true, legendPos: 'b', legendFontSize: 8,
    dataLabelFormatCode: '0.0%',
    showValue: true, dataLabelColor: CW, dataLabelFontSize: 9,
    holeSize: 50,
  });

  const rows = categorySales.slice(0, 8).map(r => [r.category, fmtKrw(r.revenue), `${r.count.toLocaleString()}건`, `${r.pct.toFixed(1)}%`]);
  addTable(slide, ['카테고리', '매출', '건수', '비율'], rows, 7.0, 1.0, 5.9, 5.5);
}

function slideChannelSales(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '유입경로별 매출 분석', 7, rc);
  const { channelSales } = analysis;
  if (!channelSales.length) return;

  const top10 = channelSales.slice(0, 10);
  slide.addChart('bar', [{
    name: '채널별 매출',
    labels: top10.map(r => r.channel),
    values: top10.map(r => r.revenue),
  }], {
    x: 0.3, y: 0.9, w: 7.5, h: 5.5,
    barDir: 'bar',
    chartColors: [CS],
    showValue: true, dataLabelColor: CD, dataLabelFontSize: 8,
    catAxisLabelFontSize: 9,
    valAxisHidden: true,
    showLegend: false,
  });

  slide.addShape('rect', { x: 8.0, y: 0.9, w: 5.0, h: 0.38, fill: { color: CS }, line: { color: CS } });
  slide.addText('채널별 평균 객단가 TOP 5', { x: 8.15, y: 0.97, w: 4.5, h: 0.3, fontSize: 10, bold: true, color: CW, fontFace: FONT });
  const top5avg = [...channelSales].sort((a, b) => b.avg - a.avg).slice(0, 5);
  const avgRows = top5avg.map(r => [r.channel, fmtKrw(r.avg), `${r.count.toLocaleString()}건`]);
  addTable(slide, ['채널', '평균 객단가', '방문수'], avgRows, 8.0, 1.35, 5.0, 2.5);
}

function slideDoctorSales(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '진료의별 매출 비교', 6, rc);
  const { doctorSales } = analysis;
  if (!doctorSales.length) return;

  // Aggregate by doctor and category for grouped bar
  const doctors  = [...new Set(doctorSales.map(r => r.doctor))].slice(0, 6);
  const cats     = [...new Set(doctorSales.map(r => r.category))].slice(0, 6);
  const chartData = cats.map((cat, i) => ({
    name: cat,
    labels: doctors,
    values: doctors.map(doc => {
      const row = doctorSales.find(r => r.doctor === doc && r.category === cat);
      return row?.revenue || 0;
    }),
  }));

  slide.addChart('bar', chartData, {
    x: 0.3, y: 0.9, w: 9, h: 5.5,
    barDir: 'col',
    barGrouping: 'clustered',
    chartColors: CHART_COLORS.slice(0, cats.length),
    showLegend: true, legendPos: 'b', legendFontSize: 8,
    catAxisLabelFontSize: 9,
    valAxisHidden: true,
  });

  const docTotals = doctors.map(doc => {
    const rev = doctorSales.filter(r => r.doctor === doc).reduce((s, r) => s + r.revenue, 0);
    return [doc, fmtKrw(rev)];
  });
  addTable(slide, ['진료의', '총 매출'], docTotals, 9.5, 1.0, 3.5, 4.5);
}

function slideNationality(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '국적별 매출 & 선호 시술 분석', 9, rc);
  const { nationalitySales } = analysis;
  if (!nationalitySales.length) return;

  const DOMESTIC = new Set(['한국', '내국인', 'KOR', 'Korean', '한국인']);
  const domRev = nationalitySales.filter(r => DOMESTIC.has(r.nationality)).reduce((s, r) => s + r.revenue, 0);
  const forRev = nationalitySales.filter(r => !DOMESTIC.has(r.nationality)).reduce((s, r) => s + r.revenue, 0);

  if (domRev + forRev > 0) {
    slide.addChart('doughnut', [{
      name: '내/외국인 비율',
      labels: ['내국인', '외국인'],
      values: [domRev, forRev],
    }], {
      x: 0.3, y: 0.9, w: 5.5, h: 5.0,
      chartColors: [CS, CA],
      showLegend: true, legendPos: 'b', legendFontSize: 9,
      showValue: true, dataLabelColor: CW, dataLabelFontSize: 10,
      holeSize: 50,
    });
  }

  const rows = nationalitySales.slice(0, 10).map(r => [r.nationality, fmtKrw(r.revenue), `${r.count}건`, fmtKrw(r.avg)]);
  addTable(slide, ['국적', '매출', '방문수', '평균 객단가'], rows, 6.0, 1.0, 6.9, 5.5);
}

function slideCrossSelling(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '복합시술 (Cross-selling) 분석', 5, rc);
  const cs = analysis.crossSelling;
  if (!cs) return;

  kpiBox(slide, 0.5, 1.0, 3.5, 1.3, '복합시술 비율', `${cs.crossSellRate.toFixed(1)}%`, '2개 이상 시술 환자');
  kpiBox(slide, 4.3, 1.0, 3.5, 1.3, '객단가 배수', `${cs.avgMultiplier.toFixed(2)}x`, '단일 시술 대비');

  slide.addShape('rect', { x: 0.5, y: 2.6, w: 7.5, h: 0.38, fill: { color: CP }, line: { color: CP } });
  slide.addText('TOP 복합시술 조합', { x: 0.65, y: 2.65, w: 6, h: 0.3, fontSize: 11, bold: true, color: CW, fontFace: FONT });

  if (cs.topCombinations.length) {
    const rows = cs.topCombinations.map(r => [r.combo, `${r.count}건`]);
    addTable(slide, ['시술 조합', '횟수'], rows, 0.5, 3.1, 7.5, 3.2);
  }

  // Insight panel
  slide.addShape('rect', { x: 8.3, y: 1.0, w: 4.7, h: 5.5, fill: { color: CL }, line: { color: 'D1E8F7' } });
  slide.addText('💡 인사이트', { x: 8.5, y: 1.15, w: 4.3, h: 0.4, fontSize: 11, bold: true, color: CP, fontFace: FONT });
  const insights = [
    `복합시술 환자 ${cs.crossSellRate.toFixed(1)}% 비율`,
    `복합시술 고객 ${cs.avgMultiplier.toFixed(1)}배 높은 객단가`,
    'TOP 조합 패키지화로 추가 매출 창출 가능',
  ];
  for (const [i, ins] of insights.entries()) {
    slide.addText(`• ${ins}`, { x: 8.5, y: 1.65 + i * 0.75, w: 4.3, h: 0.7, fontSize: 10, color: CD, fontFace: FONT });
  }
}

function slideTrends(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '일별/요일별 매출 추이', 11, rc);
  const { dailyTrends, weekdayTrends } = analysis;

  if (dailyTrends.length) {
    slide.addChart('line', [{
      name: '일별 매출',
      labels: dailyTrends.map(d => d.date.slice(5)),
      values: dailyTrends.map(d => d.revenue),
    }], {
      x: 0.3, y: 0.9, w: 12.7, h: 3.5,
      chartColors: [CS],
      showLegend: false,
      lineDataSymbol: 'none',
      catAxisLabelFontSize: 7,
      valAxisHidden: true,
      title: `${rc.year}년 ${rc.month}월 일별 매출`,
    });
  }

  if (weekdayTrends.length) {
    slide.addChart('bar', [{
      name: '요일별 매출',
      labels: weekdayTrends.map(d => d.name + '요일'),
      values: weekdayTrends.map(d => d.revenue),
    }], {
      x: 0.3, y: 4.4, w: 12.7, h: 2.8,
      barDir: 'col',
      chartColors: weekdayTrends.map((d) => {
        const max = Math.max(...weekdayTrends.map(x => x.revenue));
        return d.revenue === max ? CA : CS;
      }),
      showLegend: false,
      catAxisLabelFontSize: 9,
      valAxisHidden: true,
    });
  }
}

function slidePatientSpending(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '환자 객단가 분포 & VIP 분석', 10, rc);
  const { patientSpending } = analysis;
  if (!patientSpending.length) return;

  const totals = patientSpending.map(p => p.total).sort((a, b) => a - b);
  const buckets = 12;
  const min = totals[0], max = totals[totals.length - 1];
  const step = (max - min) / buckets;
  const bucketCounts = Array.from({ length: buckets }, (_, i) => {
    const lo = min + i * step, hi = lo + step;
    return totals.filter(v => v >= lo && v < hi).length;
  });
  const bucketLabels = Array.from({ length: buckets }, (_, i) => fmtKrw(min + i * step));

  slide.addChart('bar', [{ name: '분포', labels: bucketLabels, values: bucketCounts }], {
    x: 0.3, y: 0.9, w: 7.5, h: 5.5,
    barDir: 'col',
    chartColors: [CS],
    showLegend: false,
    catAxisLabelFontSize: 7,
    valAxisHidden: false,
    title: '객단가 분포 (환자 수)',
  });

  const mean = totals.reduce((s, v) => s + v, 0) / totals.length;
  const median = totals[Math.floor(totals.length / 2)];
  const p90 = totals[Math.floor(totals.length * 0.9)];

  const stats = [['평균 객단가', fmtKrw(mean)], ['중앙값', fmtKrw(median)],
    ['상위 10% (P90)', fmtKrw(p90)], ['최고 객단가', fmtKrw(max)]];
  for (const [i, [label, value]] of stats.entries()) {
    const y = 1.4 + i * 0.7;
    slide.addShape('rect', { x: 8.0, y, w: 5.0, h: 0.6, fill: { color: CG2 }, line: { color: 'E5E7EB' } });
    slide.addText(label, { x: 8.15, y: y + 0.05, w: 3.0, h: 0.28, fontSize: 8, color: CG, fontFace: FONT });
    slide.addText(value, { x: 8.15, y: y + 0.28, w: 4.7, h: 0.28, fontSize: 13, bold: true, color: CP, fontFace: FONT });
  }
}

function slideInsights(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '종합 인사이트 & 전략 제언', 12, rc);
  const { insights } = analysis;
  const perspectives: [string, string[], string][] = [
    ['원장님 관점', insights.doctor,   CP],
    ['마케팅 관점', insights.marketing, CS],
    ['운영 관점',   insights.operations, '27AE60'],
  ];
  const cw = 4.1, gap = 0.25, sx = 0.3;

  for (const [i, [title, items, color]] of perspectives.entries()) {
    const cx = sx + i * (cw + gap);
    slide.addShape('rect', { x: cx, y: 0.9, w: cw, h: 0.45, fill: { color }, line: { color } });
    slide.addText(title, { x: cx + 0.1, y: 0.96, w: cw - 0.2, h: 0.35, fontSize: 11, bold: true, color: CW, align: 'center', fontFace: FONT });
    slide.addShape('rect', { x: cx, y: 1.35, w: cw, h: 5.6, fill: { color: CG2 }, line: { color: 'E5E7EB' } });
    for (const [j, item] of items.slice(0, 4).entries()) {
      const iy = 1.5 + j * 1.25;
      slide.addShape('rect', { x: cx + 0.1, y: iy, w: 0.25, h: 0.25, fill: { color }, line: { color } });
      slide.addText(item, { x: cx + 0.45, y: iy - 0.02, w: cw - 0.6, h: 1.1, fontSize: 9, color: CD, fontFace: FONT });
    }
  }
}

function slideAppendixCover(pptx: PptxGenJS, rc: ReportConfig) {
  const slide = pptx.addSlide();
  slide.addShape('rect', { x: 0, y: 0, w: SW, h: SH, fill: { color: CG2 }, line: { color: CG2 } });
  slide.addShape('rect', { x: 0, y: 2.5, w: SW, h: 2.5, fill: { color: CS }, line: { color: CS } });
  slide.addText('부 록', { x: 0, y: 2.85, w: SW, h: 1.2, fontSize: 36, bold: true, color: CW, align: 'center', fontFace: FONT });
  slide.addText('Appendix — 상세 데이터', { x: 0, y: 4.0, w: SW, h: 0.6, fontSize: 16, color: CW, align: 'center', fontFace: FONT });
  const footer = `${rc.hospitalName} | ${rc.year}년 ${rc.month}월`;
  slide.addText(footer, { x: 0.4, y: SH - 0.35, w: SW - 0.8, h: 0.3, fontSize: 7, color: CG, fontFace: FONT });
}

function slideVip(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '[부록] VIP 환자 분석 (상위 10%)', 14, rc);
  const { patientSpending } = analysis;
  if (!patientSpending.length) return;

  const threshold = patientSpending[Math.floor(patientSpending.length * 0.1)].total;
  const vips = patientSpending.filter(p => p.total >= threshold).slice(0, 20);

  slide.addShape('rect', { x: 0.4, y: 1.0, w: SW - 0.8, h: 0.4, fill: { color: CL }, line: { color: 'D1E8F7' } });
  slide.addText(`VIP 기준: ${fmtKrw(threshold)} 이상 · 해당 환자 ${vips.length}명`, {
    x: 0.6, y: 1.05, w: 12, h: 0.3, fontSize: 10, color: CP, fontFace: FONT,
  });

  const rows = vips.map(p => [p.chartNo, fmtKrw(p.total), `${p.visits}회`]);
  addTable(slide, ['차트번호', '총 매출', '방문횟수'], rows, 0.4, 1.5, SW - 0.8, 5.0);
}

// ── PUBLIC ────────────────────────────────────────────────────────────────────

export async function generatePPT(
  analysis: AnalysisResult,
  rc: ReportConfig,
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  _cfg: HospitalConfig
): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'CUSTOM', width: SW, height: SH });
  pptx.layout = 'CUSTOM';

  slideCover(pptx, rc);
  slideExecutiveSummary(pptx, analysis, rc);
  slideCategorySales(pptx, analysis, rc);
  slideCrossSelling(pptx, analysis, rc);
  slideDoctorSales(pptx, analysis, rc);
  slideChannelSales(pptx, analysis, rc);
  slideNationality(pptx, analysis, rc);
  slidePatientSpending(pptx, analysis, rc);
  slideTrends(pptx, analysis, rc);
  slideInsights(pptx, analysis, rc);
  slideAppendixCover(pptx, rc);
  slideVip(pptx, analysis, rc);

  return await pptx.write({ outputType: 'nodebuffer' }) as Buffer;
}
