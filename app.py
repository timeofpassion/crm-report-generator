import os
import io
import yaml
import traceback
from datetime import datetime

import streamlit as st
import pandas as pd

from modules.excel_parser import parse_files
from modules.data_analyzer import DataAnalyzer
from modules.ppt_generator import PPTGenerator

# ── 페이지 설정 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CRM 매출분석 보고서 생성기",
    page_icon="📊",
    layout="wide",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { padding-top: 1rem; }
    .stMetric { background: #EBF3FB; border-radius: 8px; padding: 12px; }
    h1 { color: #1A3A6B; }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_configs():
    path = os.path.join(os.path.dirname(__file__), 'config', 'hospital_configs.yaml')
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def fmt_krw(v):
    v = float(v)
    if v >= 100_000_000:
        return f"₩{v/100_000_000:.1f}억"
    elif v >= 10_000:
        return f"₩{v/10_000:.0f}만"
    return f"₩{v:,.0f}"


# ── 메인 ───────────────────────────────────────────────────────────────────────
def main():
    st.title("📊 CRM 매출분석 보고서 자동 생성기")
    st.caption("CRM 엑셀 파일을 업로드하면 자동으로 매출분석 PPT 보고서를 생성합니다.")

    configs = load_configs()
    hospitals = configs['hospitals']

    # ── 사이드바 ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ 보고서 설정")

        hospital_key = st.selectbox(
            "병원 선택",
            options=list(hospitals.keys()),
            format_func=lambda k: hospitals[k]['name'],
        )
        h_cfg = hospitals[hospital_key]

        hospital_name = st.text_input("병원명 (표지)", value=h_cfg['name'])

        col1, col2 = st.columns(2)
        now = datetime.now()
        year_opts = list(range(2023, now.year + 1))
        with col1:
            year = st.selectbox("년도", year_opts, index=len(year_opts) - 1)
        with col2:
            month = st.selectbox("월", list(range(1, 13)), index=now.month - 1)

        team = st.text_input("팀명", value=h_cfg.get('team', '열정의시간 마케팅팀'))

        st.divider()
        st.subheader("📋 분석 포커스")

        focus_opts = {
            'category':    '시술 카테고리별 매출 심화',
            'channel':     '유입경로별 효율 비교',
            'cross':       '복합시술 Cross-selling',
            'doctor':      '진료의별 매출 비교',
            'nationality': '국적별 매출 & 선호 시술',
            'vip':         '객단가 분포 & VIP 분석',
            'trend':       '일별/요일별 매출 추이',
            'misc':        '기타(미분류) 상세 분석',
            'reclassify':  '기존 vs 재분류 비교',
            'lifting':     '리프팅 기기별 분석',
        }
        focus = {k: st.checkbox(v, value=True) for k, v in focus_opts.items()}

        extra = st.text_area("추가 지침 (선택)", placeholder="예: 이번 달은 리프팅 중심으로 분석해주세요")

    # ── 파일 업로드 ───────────────────────────────────────────────────────────
    st.subheader("📁 엑셀 파일 업로드")
    uploaded = st.file_uploader(
        "CRM 엑셀 파일 (오더판매내역, 분야별집계, 내원경로별 등)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="여러 파일 동시 업로드 가능. 어떤 양식이든 자동 인식됩니다.",
    )

    if not uploaded:
        with st.expander("💡 사용 방법"):
            st.markdown("""
**1단계** — 사이드바에서 병원, 기간, 팀명을 설정하세요.

**2단계** — CRM에서 받은 엑셀 파일을 드래그앤드롭 또는 버튼으로 업로드하세요.
- `오더판매내역` : 가장 상세한 분석 가능 (권장)
- `분야별집계`, `내원경로별` : 부분 분석

**3단계** — 분석 포커스를 선택하세요.

**4단계** — **보고서 생성** 버튼을 클릭하면 PPT가 자동 생성됩니다.

**5단계** — **다운로드** 버튼으로 저장하세요.
            """)
        return

    st.success(f"✅ {len(uploaded)}개 파일 업로드됨")

    # 미리보기
    with st.expander("📄 파일 미리보기", expanded=False):
        for f in uploaded:
            f.seek(0)
            try:
                df_preview = pd.read_excel(f, nrows=5)
                st.write(f"**{f.name}**")
                st.dataframe(df_preview, use_container_width=True)
            except Exception as e:
                st.error(f"{f.name}: {e}")

    st.divider()

    if st.button("🚀 보고서 생성", type="primary", use_container_width=True):
        with st.spinner("📊 분석 중... (30초~1분 소요)"):
            try:
                # 파싱
                for f in uploaded:
                    f.seek(0)
                parsed = parse_files(uploaded)

                if not parsed:
                    st.error("파일을 인식하지 못했습니다. 오더판매내역 엑셀을 업로드해주세요.")
                    return

                # 분석
                analyzer = DataAnalyzer(h_cfg)
                analysis = analyzer.analyze(parsed)

                # KPI 미리보기
                kpis = analysis.get('kpis', {})
                if kpis:
                    st.subheader("📊 분석 결과 미리보기")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("총 매출",    fmt_krw(kpis.get('total_revenue', 0)))
                    c2.metric("총 방문건수", f"{kpis.get('total_visits', 0):,}건")
                    c3.metric("총 환자수",  f"{kpis.get('unique_patients', 0):,}명")
                    c4.metric("평균 객단가", fmt_krw(kpis.get('avg_per_patient', 0)))

                if analysis.get('key_findings'):
                    st.subheader("💡 핵심 발견")
                    for f_text in analysis['key_findings']:
                        st.write(f"• {f_text}")

                # PPT 생성
                report_cfg = {
                    'hospital_name': hospital_name,
                    'year': year,
                    'month': month,
                    'team_name': team,
                    'focus': focus,
                    'extra': extra,
                }
                ppt_gen = PPTGenerator(h_cfg)
                ppt_bytes = ppt_gen.generate(analysis, report_cfg)

                st.success("✅ 보고서 생성 완료!")

                fname = f"{hospital_name}_{year}년{month}월_매출분석보고서.pptx"
                st.download_button(
                    label="📥 PPT 보고서 다운로드",
                    data=ppt_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    type="primary",
                )

            except Exception as e:
                st.error(f"오류 발생: {e}")
                with st.expander("상세 오류"):
                    st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
