# 📊 CRM 매출분석 보고서 자동 생성기

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://timeofpassion-crm.streamlit.app)

> CRM 엑셀 파일을 업로드하면 자동으로 매출분석 PPT 보고서를 생성합니다.  
> 열정의시간 마케팅팀 전용 내부 도구

---

## 기능

- **엑셀 자동 인식**: 오더판매내역 / 분야별집계 / 내원경로별 등 자동 감지
- **분석 항목**: KPI, 카테고리별, 진료의별, 유입경로별, 국적별, 복합시술, 일별/요일별 추이, 객단가 분포
- **PPT 출력**: 17슬라이드 + 부록 (최대 17페이지)
- **병원 지원**: 멜로우피부과 신사점 기본 설정 / 신규 병원 추가 가능

## 실행 방법 (로컬)

```bash
pip install -r requirements.txt
streamlit run app.py
```

또는 `CRM보고서생성기.bat` 더블클릭

## 기술 스택

Python · Streamlit · pandas · python-pptx · matplotlib

---

*열정의시간 마케팅팀 | ceo@timeofpassion.com*
