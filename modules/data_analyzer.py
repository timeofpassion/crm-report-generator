import pandas as pd
import numpy as np
from typing import Dict, Any, Optional


WEEKDAY_KO = ['월', '화', '수', '목', '금', '토', '일']


def _categorize(name: str, keyword_map: Dict[str, list]) -> str:
    name_str = str(name)
    for cat, keywords in keyword_map.items():
        for kw in keywords:
            if kw in name_str:
                return cat
    return '기타'


def _fmt_krw(v: float) -> str:
    if v >= 100_000_000:
        return f"{v / 100_000_000:.1f}억원"
    elif v >= 10_000:
        return f"{v / 10_000:.0f}만원"
    return f"{v:,.0f}원"


class DataAnalyzer:
    def __init__(self, hospital_config: dict):
        self.config = hospital_config
        self.kw_map: Dict[str, list] = hospital_config.get('category_keywords', {})

    def analyze(self, parsed_data: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
        order_df: Optional[pd.DataFrame] = parsed_data.get('오더판매내역')
        cat_df:   Optional[pd.DataFrame] = parsed_data.get('분야별집계')
        ch_df:    Optional[pd.DataFrame] = parsed_data.get('내원경로별')

        results: Dict[str, Any] = {}

        if order_df is not None and not order_df.empty:
            results.update(self._analyze_order(order_df))

        if cat_df is not None and not cat_df.empty and 'category_sales' not in results:
            results['original_categories'] = self._summarize_cat(cat_df)

        if ch_df is not None and not ch_df.empty and 'channel_sales' not in results:
            results['channel_sales'] = self._summarize_channel(ch_df)

        results['key_findings'] = self._generate_findings(results)
        results['insights'] = self._generate_insights(results)

        return results

    # ── 오더판매내역 분석 ──────────────────────────────────────────────────────

    def _analyze_order(self, df: pd.DataFrame) -> Dict[str, Any]:
        r: Dict[str, Any] = {}

        # 카테고리 분류
        if 'order_name' in df.columns:
            df = df.copy()
            df['category'] = df['order_name'].apply(lambda x: _categorize(x, self.kw_map))
        elif 'original_category' in df.columns:
            df = df.copy()
            df['category'] = df['original_category']
        else:
            df = df.copy()
            df['category'] = '기타'

        # KPIs
        total_rev = df['amount'].sum() if 'amount' in df.columns else 0
        total_visits = len(df)
        patients = df['chart_no'].nunique() if 'chart_no' in df.columns else total_visits
        avg_per_patient = total_rev / patients if patients else 0
        avg_per_visit   = total_rev / total_visits if total_visits else 0

        r['kpis'] = {
            'total_revenue':   total_rev,
            'total_visits':    total_visits,
            'unique_patients': patients,
            'avg_per_patient': avg_per_patient,
            'avg_per_visit':   avg_per_visit,
        }

        # 카테고리별 매출
        cat_grp = (
            df.groupby('category', as_index=False)['amount']
              .agg(revenue='sum', count='count')
              .sort_values('revenue', ascending=False)
        )
        cat_grp['pct'] = cat_grp['revenue'] / cat_grp['revenue'].sum() * 100
        r['category_sales'] = cat_grp

        # 기존(CRM) vs 재분류 비교
        if 'original_category' in df.columns:
            orig = (
                df.groupby('original_category', as_index=False)['amount']
                  .agg(orig_revenue='sum', orig_count='count')
            )
            new_ = (
                df.groupby('category', as_index=False)['amount']
                  .agg(new_revenue='sum', new_count='count')
                  .rename(columns={'category': 'original_category'})
            )
            r['reclassify_compare'] = orig.merge(new_, on='original_category', how='outer').fillna(0)

        # 진료의별
        if 'doctor' in df.columns:
            doc_grp = (
                df.groupby(['doctor', 'category'], as_index=False)['amount']
                  .agg(revenue='sum', count='count')
                  .sort_values('revenue', ascending=False)
            )
            r['doctor_sales'] = doc_grp

        # 유입경로별
        if 'channel' in df.columns:
            ch_grp = (
                df.groupby('channel', as_index=False)['amount']
                  .agg(revenue='sum', count='count')
                  .sort_values('revenue', ascending=False)
            )
            ch_grp['avg'] = ch_grp['revenue'] / ch_grp['count']
            r['channel_sales'] = ch_grp

            # 채널 × 카테고리
            try:
                pivot = df.groupby(['channel', 'category'])['amount'].sum().unstack(fill_value=0)
                r['channel_x_category'] = pivot
            except Exception:
                pass

        # 국적별
        if 'nationality' in df.columns:
            nat_grp = (
                df.groupby('nationality', as_index=False)['amount']
                  .agg(revenue='sum', count='count')
                  .sort_values('revenue', ascending=False)
            )
            nat_grp['avg'] = nat_grp['revenue'] / nat_grp['count']
            r['nationality_sales'] = nat_grp

        # 환자 객단가
        if 'chart_no' in df.columns:
            per_pt = (
                df.groupby('chart_no')['amount']
                  .agg(total='sum', visits='count')
            )
            r['patient_spending'] = per_pt

        # 복합시술 (Cross-selling)
        if 'chart_no' in df.columns:
            pt_counts = df.groupby('chart_no')['category'].nunique()
            cross_rate = (pt_counts >= 2).mean() * 100
            single_avg = per_pt.loc[pt_counts == 1, 'total'].mean() if (pt_counts == 1).any() else 1
            multi_avg  = per_pt.loc[pt_counts >= 2, 'total'].mean() if (pt_counts >= 2).any() else 0
            multiplier = multi_avg / single_avg if single_avg else 0

            # TOP 조합
            combos = (
                df.groupby('chart_no')['category']
                  .apply(lambda x: tuple(sorted(set(x))))
                  .reset_index(name='combo')
            )
            top_combos = (
                combos[combos['combo'].apply(len) >= 2]
                ['combo']
                .value_counts()
                .head(5)
                .reset_index()
            )
            top_combos.columns = ['조합', '횟수']
            top_combos['조합'] = top_combos['조합'].apply(lambda x: ' + '.join(x))

            r['cross_selling'] = {
                'cross_sell_rate': cross_rate,
                'avg_multiplier':  multiplier,
                'top_combinations': top_combos,
            }

        # 일별 / 요일별 추이
        if 'date' in df.columns:
            df_dated = df.dropna(subset=['date'])
            if not df_dated.empty:
                daily = (
                    df_dated.groupby(df_dated['date'].dt.date)['amount']
                            .agg(revenue='sum', count='count')
                            .reset_index()
                )
                daily.columns = ['date', 'revenue', 'count']
                r['daily_trends'] = daily

                df_dated = df_dated.copy()
                df_dated['weekday'] = df_dated['date'].dt.dayofweek
                wd = (
                    df_dated.groupby('weekday')['amount']
                            .agg(revenue='sum', count='count')
                            .reindex(range(7), fill_value=0)
                            .reset_index()
                )
                wd['avg'] = wd['revenue'] / wd['count'].replace(0, np.nan)
                wd['weekday_name'] = wd['weekday'].map(dict(enumerate(WEEKDAY_KO)))
                r['weekday_trends'] = wd

        # 리프팅 기기별
        if 'order_name' in df.columns:
            lifting_kws = self.kw_map.get('리프팅', [])
            lift_df = df[df['category'] == '리프팅'].copy() if '리프팅' in df['category'].values else pd.DataFrame()
            if not lift_df.empty and 'order_name' in lift_df.columns:
                lift_grp = (
                    lift_df.groupby('order_name')['amount']
                           .agg(revenue='sum', count='count')
                           .sort_values('revenue', ascending=False)
                           .head(15)
                           .reset_index()
                )
                r['lifting_detail'] = lift_grp

            # 기타 세부
            misc_df = df[df['category'] == '기타'].copy() if '기타' in df['category'].values else pd.DataFrame()
            if not misc_df.empty and 'order_name' in misc_df.columns:
                misc_grp = (
                    misc_df.groupby('order_name')['amount']
                           .agg(revenue='sum', count='count')
                           .sort_values('revenue', ascending=False)
                           .head(20)
                           .reset_index()
                )
                r['misc_detail'] = misc_grp

        return r

    # ── 분야별집계 파싱 ────────────────────────────────────────────────────────

    def _summarize_cat(self, df: pd.DataFrame) -> pd.DataFrame:
        if 'original_category' not in df.columns or 'amount' not in df.columns:
            return df
        grp = (
            df.groupby('original_category', as_index=False)['amount']
              .sum()
              .sort_values('amount', ascending=False)
        )
        grp['pct'] = grp['amount'] / grp['amount'].sum() * 100
        return grp

    def _summarize_channel(self, df: pd.DataFrame) -> pd.DataFrame:
        ch_col = 'channel' if 'channel' in df.columns else 'original_category'
        if ch_col not in df.columns or 'amount' not in df.columns:
            return df
        grp = (
            df.groupby(ch_col, as_index=False)['amount']
              .agg(revenue='sum', count='count' if 'count' in df.columns else lambda x: len(x))
              .sort_values('revenue', ascending=False)
        )
        return grp

    # ── 인사이트 생성 ─────────────────────────────────────────────────────────

    def _generate_findings(self, r: Dict) -> list:
        findings = []
        try:
            if 'category_sales' in r and not r['category_sales'].empty:
                top = r['category_sales'].iloc[0]
                findings.append(f"가장 높은 매출 카테고리: {top['category']} ({top['pct']:.1f}%)")

            if 'channel_sales' in r and not r['channel_sales'].empty:
                top_ch = r['channel_sales'].iloc[0]
                findings.append(f"주요 유입경로: {top_ch['channel']} (방문 {top_ch['count']:,}건)")

            if 'cross_selling' in r:
                cs = r['cross_selling']
                findings.append(f"복합시술 비율: {cs['cross_sell_rate']:.1f}% (단일 대비 {cs['avg_multiplier']:.1f}배 객단가)")

            if 'kpis' in r:
                kpis = r['kpis']
                findings.append(f"평균 객단가: {_fmt_krw(kpis['avg_per_patient'])} (총 {kpis['unique_patients']:,}명)")
        except Exception:
            pass
        return findings

    def _generate_insights(self, r: Dict) -> Dict[str, list]:
        insights = {'원장님 관점': [], '마케팅 관점': [], '운영 관점': []}
        try:
            if 'category_sales' in r and not r['category_sales'].empty:
                top2 = r['category_sales'].head(2)
                cats = ' / '.join(top2['category'].tolist())
                insights['원장님 관점'].append(f"핵심 매출 시술: {cats} — 집중 역량 강화 권고")
                insights['원장님 관점'].append("고객당 평균 내원 횟수 기준으로 재방문 시술 패키지 기획 검토")

            if 'channel_sales' in r and not r['channel_sales'].empty:
                top_ch = r['channel_sales'].iloc[0]
                insights['마케팅 관점'].append(f"주요 채널 '{top_ch['channel']}' 집중 투자로 ROI 극대화")
                if len(r['channel_sales']) > 1:
                    ch2 = r['channel_sales'].iloc[1]
                    insights['마케팅 관점'].append(f"2위 채널 '{ch2['channel']}' 성장 가능성 테스트 권장")

            if 'patient_spending' in r:
                pt = r['patient_spending']
                p90 = pt['total'].quantile(0.9)
                insights['원장님 관점'].append(f"상위 10% 고객 객단가: {_fmt_krw(p90)} — VIP 케어 프로그램 도입 검토")

            if 'weekday_trends' in r and not r['weekday_trends'].empty:
                wt = r['weekday_trends']
                best_day = wt.loc[wt['revenue'].idxmax(), 'weekday_name']
                worst_day = wt.loc[wt['revenue'].idxmin(), 'weekday_name']
                insights['운영 관점'].append(f"최다 매출 요일: {best_day}요일 — 스태프 배치 강화 권고")
                insights['운영 관점'].append(f"최저 매출 요일: {worst_day}요일 — 프로모션/패키지 행사 기획")

            insights['운영 관점'].append("내원 경로별 예약 시스템 최적화로 대기 시간 단축 검토")
        except Exception:
            pass
        return insights
