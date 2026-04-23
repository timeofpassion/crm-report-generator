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
const CG3 = 'E5E7EB';  // Border gray

const CHART_COLORS = [CP, CS, CA, CX, '27AE60', '9B59B6', 'E67E22', '3498DB'];
const FONT         = '맑은 고딕';
const SW = 13.33, SH = 7.5;
const TOTAL = 16;

// 마케팅 채널 분류 키워드
const MKT_DIGITAL   = ['네이버', '인스타', 'sns', 'sns', '유튜브', '블로그', '광고', '카카오', '검색', '온라인', '배너', '페이스북', '틱톡', '구글', '포털'];
const MKT_REFERRAL  = ['지인', '소개', '추천'];
const MKT_REVISIT   = ['재방문', '기존', '단골'];

function classifyChannel(channel: string): string {
  const c = channel.toLowerCase();
  if (MKT_DIGITAL.some(k => c.includes(k))) return '디지털';
  if (MKT_REFERRAL.some(k => c.includes(k))) return '지인소개';
  if (MKT_REVISIT.some(k => c.includes(k))) return '재방문';
  return '기타';
}

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
  slide.addText(`${page} / ${TOTAL}`, { x: SW - 1.1, y: 0.22, w: 0.9, h: 0.4, fontSize: 10, color: CW, align: 'right', fontFace: FONT });
  const footer = `${rc.hospitalName} | ${rc.year}년 ${rc.month}월 | ${rc.teamName}`;
  slide.addText(footer, { x: 0.4, y: SH - 0.35, w: SW - 0.8, h: 0.3, fontSize: 7, color: CG, fontFace: FONT });
}

function kpiBox(slide: SlideIface, x: number, y: number, w: number, h: number, label: string, value: string, sub = '') {
  slide.addShape('rect', { x, y, w, h, fill: { color: CL }, line: { color: CG3 } });
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
    border: { pt: 0.5, color: CG3 },
  });
}

function sectionBadge(slide: SlideIface, label: string) {
  slide.addShape('rect', { x: SW - 2.4, y: 0.18, w: 2.0, h: 0.5, fill: { color: CA }, line: { color: CA } });
  slide.addText(label, { x: SW - 2.4, y: 0.22, w: 2.0, h: 0.4, fontSize: 8, bold: true, color: CD, align: 'center', fontFace: FONT });
}

// ── 슬라이드 생성 함수 ─────────────────────────────────────────────────────────

function slideCover(pptx: PptxGenJS, rc: ReportConfig) {
  const slide = pptx.addSlide();
  slide.addShape('rect', { x: 0, y: 0, w: SW, h: SH, fill: { color: '1A1A2E' }, line: { color: '1A1A2E' } });
  slide.addShape('rect', { x: 0, y: 0, w: 0.18, h: SH, fill: { color: CP }, line: { color: CP } });
  slide.addShape('rect', { x: 0, y: SH - 0.12, w: SW, h: 0.12, fill: { color: CA }, line: { color: CA } });
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

// ── [섹션 1] 시술별/수술별 매출 동향분석 ────────────────────────────────────────

function slideCategorySales(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '시술 카테고리별 매출 현황', 3, rc);
  sectionBadge(slide, '섹션 1 · 시술 동향');
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

function slideTopProcedures(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '시술별 매출 TOP 순위', 4, rc);
  sectionBadge(slide, '섹션 1 · 시술 동향');
  const { topProcedures } = analysis;

  if (!topProcedures.length) {
    slide.addText('시술별 상세 데이터 없음 — 오더판매내역 파일을 업로드하면 시술별 순위를 확인할 수 있습니다.', {
      x: 0.4, y: 3.2, w: SW - 0.8, h: 0.5, fontSize: 11, color: CG, align: 'center', fontFace: FONT,
    });
    return;
  }

  const top10 = topProcedures.slice(0, 10);
  slide.addChart('bar', [{
    name: '시술별 매출',
    labels: top10.map(r => r.name.length > 14 ? r.name.slice(0, 14) + '..' : r.name),
    values: top10.map(r => r.revenue),
  }], {
    x: 0.3, y: 0.9, w: 7.5, h: 5.8,
    barDir: 'bar',
    chartColors: top10.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]),
    showValue: true, dataLabelColor: CD, dataLabelFontSize: 8,
    catAxisLabelFontSize: 8,
    valAxisHidden: true,
    showLegend: false,
  });

  const rows = topProcedures.slice(0, 12).map((r, i) => [
    String(i + 1),
    r.name.length > 18 ? r.name.slice(0, 18) + '..' : r.name,
    fmtKrw(r.revenue),
    `${r.count}건`,
    `${r.pct.toFixed(1)}%`,
    r.category,
  ]);
  addTable(slide, ['순위', '시술명', '매출', '건수', '비율', '카테고리'], rows, 7.9, 0.9, 5.2, 5.8);
}

// ── [섹션 2] 유입경로 분석 ────────────────────────────────────────────────────

function slideChannelSales(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '유입경로별 매출 분석', 5, rc);
  sectionBadge(slide, '섹션 2 · 유입경로');
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

// ── [섹션 3] 마케팅 유입 및 매출 상관관계 ───────────────────────────────────────

function slideMarketingCorrelation(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '마케팅 유입 & 매출 상관관계', 6, rc);
  sectionBadge(slide, '섹션 3 · 마케팅 상관관계');
  const { channelSales } = analysis;

  if (!channelSales.length) {
    slide.addText('유입경로 데이터가 없습니다.', { x: 0.4, y: 3.5, w: SW - 0.8, h: 0.5, fontSize: 11, color: CG, align: 'center', fontFace: FONT });
    return;
  }

  // 채널 유형별 집계
  const typeMap = new Map<string, { revenue: number; count: number }>();
  for (const c of channelSales) {
    const t = classifyChannel(c.channel);
    const prev = typeMap.get(t) || { revenue: 0, count: 0 };
    typeMap.set(t, { revenue: prev.revenue + c.revenue, count: prev.count + c.count });
  }
  const typeOrder = ['디지털', '지인소개', '재방문', '기타'];
  const typeColors: Record<string, string> = { '디지털': CP, '지인소개': CS, '재방문': '27AE60', '기타': CG };

  // 유형별 KPI 박스
  const validTypes = typeOrder.filter(t => typeMap.has(t));
  const bw = (SW - 0.8) / validTypes.length - 0.2;
  for (const [i, t] of validTypes.entries()) {
    const d = typeMap.get(t)!;
    const avg = d.count > 0 ? d.revenue / d.count : 0;
    kpiBox(slide, 0.4 + i * (bw + 0.2), 0.95, bw, 1.15,
      t,
      fmtKrw(d.revenue),
      `유입 ${d.count}건 · 건당 ${fmtKrw(avg)}`
    );
  }

  // 왼쪽 차트: 채널별 유입건수 (볼륨)
  const byVolume = channelSales.filter(c => c.count > 0).slice(0, 8);
  slide.addChart('bar', [{
    name: '유입건수',
    labels: byVolume.map(r => r.channel),
    values: byVolume.map(r => r.count),
  }], {
    x: 0.3, y: 2.3, w: 6.3, h: 4.7,
    barDir: 'bar',
    chartColors: [CS],
    showValue: true, dataLabelColor: CD, dataLabelFontSize: 8,
    catAxisLabelFontSize: 8,
    valAxisHidden: true,
    showLegend: false,
    title: '채널별 유입 건수 (볼륨)',
  });

  // 오른쪽 표: 채널별 효율 매트릭스
  const rows = channelSales.slice(0, 10).map(r => {
    const type = classifyChannel(r.channel);
    const effLabel = r.avg >= (channelSales[0]?.avg ?? 0) * 0.8 ? '상' : r.avg >= (channelSales[0]?.avg ?? 0) * 0.4 ? '중' : '하';
    return [r.channel, `${r.count}건`, fmtKrw(r.avg), fmtKrw(r.revenue), type, effLabel];
  });
  addTable(slide, ['채널', '유입', '건당', '총매출', '유형', '효율'], rows, 6.8, 2.3, 6.3, 4.7);
}

// ── 일별/요일별 추이 ────────────────────────────────────────────────────────────

function slideTrends(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '일별/요일별 매출 추이', 7, rc);
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
    const maxRev = Math.max(...weekdayTrends.map(x => x.revenue));
    slide.addChart('bar', [{
      name: '요일별 매출',
      labels: weekdayTrends.map(d => d.name + '요일'),
      values: weekdayTrends.map(d => d.revenue),
    }], {
      x: 0.3, y: 4.4, w: 12.7, h: 2.8,
      barDir: 'col',
      chartColors: weekdayTrends.map(d => d.revenue === maxRev ? CA : CS),
      showLegend: false,
      catAxisLabelFontSize: 9,
      valAxisHidden: true,
    });
  }
}

// ── [섹션 4] 다음달 마케팅 전략방향성 제안 ──────────────────────────────────────

function slideNextMonthStrategy(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  const nextMonth = rc.month >= 12 ? 1 : rc.month + 1;
  const nextYear  = rc.month >= 12 ? rc.year + 1 : rc.year;
  header(slide, `${nextYear}년 ${nextMonth}월 마케팅 전략 방향성 제안`, 8, rc);
  sectionBadge(slide, '섹션 4 · 전략 방향성');

  // 데이터 기반 전략 카드 생성
  const cards: { icon: string; title: string; desc: string; color: string }[] = [];

  if (analysis.categorySales.length > 0) {
    const top = analysis.categorySales[0];
    const second = analysis.categorySales[1];
    cards.push({
      icon: '🎯',
      title: `집중 강화 — ${top.category}`,
      desc: `이달 매출 비중 ${top.pct.toFixed(1)}%(1위). 핵심 시술 관련 콘텐츠 및 프로모션 집중 기획. ${second ? `2위 ${second.category}(${second.pct.toFixed(1)}%)와의 격차 유지 전략 병행.` : ''}`,
      color: CP,
    });
  }

  if (analysis.channelSales.length > 0) {
    const byEff = [...analysis.channelSales].filter(c => c.count >= 2).sort((a, b) => b.avg - a.avg);
    const topEff = byEff[0];
    const topVol = analysis.channelSales[0];
    if (topEff && topEff.channel !== topVol.channel) {
      cards.push({
        icon: '📢',
        title: `고효율 채널 투자 — ${topEff.channel}`,
        desc: `건당 매출 ${fmtKrw(topEff.avg)}로 효율 1위. 예산 집중으로 ROI 극대화. 유입량 1위 '${topVol.channel}'은 현 수준 유지.`,
        color: CS,
      });
    } else {
      cards.push({
        icon: '📢',
        title: `핵심 채널 강화 — ${topVol.channel}`,
        desc: `유입건수 및 매출 모두 1위(${fmtKrw(topVol.revenue)}, ${topVol.count}건). 광고 소재 리프레시 및 예산 유지로 성과 지속.`,
        color: CS,
      });
    }
  }

  if (analysis.crossSelling && analysis.crossSelling.crossSellRate < 35) {
    const cs = analysis.crossSelling;
    cards.push({
      icon: '📦',
      title: '복합시술 패키지 프로모션',
      desc: `현재 복합시술 비율 ${cs.crossSellRate.toFixed(1)}%. TOP 조합(${cs.topCombinations[0]?.combo ?? ''}) 기반 패키지 할인 설계로 객단가 ${cs.avgMultiplier.toFixed(1)}배 수준 유도.`,
      color: '27AE60',
    });
  } else if (analysis.crossSelling) {
    const cs = analysis.crossSelling;
    cards.push({
      icon: '⭐',
      title: 'VIP 리텐션 프로그램',
      desc: `복합시술 비율 ${cs.crossSellRate.toFixed(1)}% 양호. 상위 10% 고객 대상 VIP 멤버십/선예약 혜택으로 재방문율 제고 집중.`,
      color: '27AE60',
    });
  }

  if (analysis.weekdayTrends.length > 0) {
    const active = analysis.weekdayTrends.filter(d => d.count > 0);
    if (active.length > 0) {
      const worst = active.reduce((a, b) => a.revenue < b.revenue ? a : b);
      cards.push({
        icon: '📅',
        title: `${worst.name}요일 수요 창출 이벤트`,
        desc: `${worst.name}요일 매출 최저(${fmtKrw(worst.revenue)}). 해당 요일 한정 특가 또는 이벤트 기획으로 빈 예약 슬롯 채움. SNS 사전 공지 1주 전 집행.`,
        color: CA.replace('F5', 'D4'),
      });
    }
  }

  // 2×2 카드 배치
  for (const [i, card] of cards.slice(0, 4).entries()) {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 6.45;
    const y = 1.0 + row * 2.95;
    slide.addShape('rect', { x, y, w: 6.25, h: 2.75, fill: { color: CG2 }, line: { color: CG3 } });
    slide.addShape('rect', { x, y, w: 0.14, h: 2.75, fill: { color: card.color }, line: { color: card.color } });
    slide.addText(`${card.icon}  ${card.title}`, {
      x: x + 0.25, y: y + 0.18, w: 5.85, h: 0.45,
      fontSize: 11, bold: true, color: CD, fontFace: FONT,
    });
    slide.addText(card.desc, {
      x: x + 0.25, y: y + 0.72, w: 5.85, h: 1.85,
      fontSize: 9.5, color: CG, fontFace: FONT,
    });
  }
}

// ── [섹션 5] 추가 견적사항 ────────────────────────────────────────────────────

function slideQuotation(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  const nextMonth = rc.month >= 12 ? 1 : rc.month + 1;
  const nextYear  = rc.month >= 12 ? rc.year + 1 : rc.year;
  header(slide, '추가 견적사항', 9, rc);
  sectionBadge(slide, '섹션 5 · 견적');

  slide.addShape('rect', { x: 0.4, y: 1.0, w: SW - 0.8, h: 0.38, fill: { color: CL }, line: { color: CG3 } });
  slide.addText(
    `${nextYear}년 ${nextMonth}월 마케팅 추가 서비스 견적 | ${rc.hospitalName} | ${rc.teamName}`,
    { x: 0.6, y: 1.07, w: 12, h: 0.28, fontSize: 9, color: CP, fontFace: FONT }
  );

  // 분석 기반 견적 제안 항목 자동 생성
  const items: string[][] = [];

  if (analysis.channelSales.length > 0) {
    const top = analysis.channelSales[0];
    items.push([`${top.channel} 광고 강화`, `이달 유입 1위 채널 집중 투자`, '', '', '']);
  }
  if (analysis.crossSelling && analysis.crossSelling.crossSellRate < 35) {
    items.push(['복합시술 패키지 기획', 'TOP 조합 기반 패키지 설계 및 소재 제작', '', '', '']);
  }
  if (analysis.topProcedures.length > 0) {
    items.push([`${analysis.topProcedures[0].name} 콘텐츠 제작`, '매출 1위 시술 홍보 영상/이미지 제작', '', '', '']);
  }
  items.push(['SNS 추가 캠페인', '인스타그램/페이스북 추가 광고 집행', '', '', '']);
  items.push(['이벤트 기획 및 운영', '월 특가 이벤트 기획·디자인·운영', '', '', '']);
  items.push(['리뷰 관리', 'O2O 플랫폼 리뷰 모니터링 및 대응', '', '', '']);

  // 합계 행
  items.push(['합계', '', '', '', '']);

  addTable(
    slide,
    ['항목', '내용', '단가', '수량', '합계'],
    items,
    0.4, 1.45, SW - 0.8, 5.4
  );

  slide.addText(
    '※ 위 항목은 이달 분석 결과를 바탕으로 제안된 선택 사항입니다. 금액은 별도 협의 후 확정됩니다.',
    { x: 0.4, y: 6.88, w: SW - 0.8, h: 0.28, fontSize: 7.5, color: CG, italic: true, fontFace: FONT }
  );
}

// ── 부가 분석 슬라이드 ────────────────────────────────────────────────────────

function slideCrossSelling(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '복합시술 (Cross-selling) 분석', 10, rc);
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

  slide.addShape('rect', { x: 8.3, y: 1.0, w: 4.7, h: 5.5, fill: { color: CL }, line: { color: CG3 } });
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

function slideDoctorSales(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '진료의별 매출 비교', 11, rc);
  const { doctorSales } = analysis;
  if (!doctorSales.length) return;

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
  header(slide, '국적별 매출 & 선호 시술 분석', 12, rc);
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

function slidePatientSpending(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '환자 객단가 분포 & VIP 분석', 13, rc);
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

  const mean   = totals.reduce((s, v) => s + v, 0) / totals.length;
  const median = totals[Math.floor(totals.length / 2)];
  const p90    = totals[Math.floor(totals.length * 0.9)];

  const stats = [['평균 객단가', fmtKrw(mean)], ['중앙값', fmtKrw(median)],
    ['상위 10% (P90)', fmtKrw(p90)], ['최고 객단가', fmtKrw(max)]];
  for (const [i, [label, value]] of stats.entries()) {
    const y = 1.4 + i * 0.7;
    slide.addShape('rect', { x: 8.0, y, w: 5.0, h: 0.6, fill: { color: CG2 }, line: { color: CG3 } });
    slide.addText(label, { x: 8.15, y: y + 0.05, w: 3.0, h: 0.28, fontSize: 8, color: CG, fontFace: FONT });
    slide.addText(value, { x: 8.15, y: y + 0.28, w: 4.7, h: 0.28, fontSize: 13, bold: true, color: CP, fontFace: FONT });
  }
}

function slideInsights(pptx: PptxGenJS, analysis: AnalysisResult, rc: ReportConfig) {
  const slide = pptx.addSlide();
  header(slide, '종합 인사이트', 14, rc);
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
    slide.addShape('rect', { x: cx, y: 1.35, w: cw, h: 5.6, fill: { color: CG2 }, line: { color: CG3 } });
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
  header(slide, '[부록] VIP 환자 분석 (상위 10%)', 16, rc);
  const { patientSpending } = analysis;
  if (!patientSpending.length) return;

  const threshold = patientSpending[Math.floor(patientSpending.length * 0.1)].total;
  const vips = patientSpending.filter(p => p.total >= threshold).slice(0, 20);

  slide.addShape('rect', { x: 0.4, y: 1.0, w: SW - 0.8, h: 0.4, fill: { color: CL }, line: { color: CG3 } });
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

  // 1. 표지
  slideCover(pptx, rc);
  // 2. Executive Summary
  slideExecutiveSummary(pptx, analysis, rc);
  // 섹션 1 — 시술별/수술별 매출 동향분석
  slideCategorySales(pptx, analysis, rc);     // 3
  slideTopProcedures(pptx, analysis, rc);     // 4
  // 섹션 2 — 유입경로 분석
  slideChannelSales(pptx, analysis, rc);      // 5
  // 섹션 3 — 마케팅 유입 & 매출 상관관계
  slideMarketingCorrelation(pptx, analysis, rc); // 6
  // 추이
  slideTrends(pptx, analysis, rc);            // 7
  // 섹션 4 — 다음달 마케팅 전략방향성 제안
  slideNextMonthStrategy(pptx, analysis, rc); // 8
  // 섹션 5 — 추가 견적사항
  slideQuotation(pptx, analysis, rc);         // 9
  // 부가 분석
  slideCrossSelling(pptx, analysis, rc);      // 10
  slideDoctorSales(pptx, analysis, rc);       // 11
  slideNationality(pptx, analysis, rc);       // 12
  slidePatientSpending(pptx, analysis, rc);   // 13
  slideInsights(pptx, analysis, rc);          // 14
  // 부록
  slideAppendixCover(pptx, rc);              // 15
  slideVip(pptx, analysis, rc);              // 16

  return await pptx.write({ outputType: 'nodebuffer' }) as Buffer;
}
