import pandas as pd
import numpy as np
import io
from typing import Dict, Optional, Tuple


COLUMN_ALIASES = {
    'chart_no':       ['차트번호', '차트 번호', '환자번호', '환자 번호', '차트', '환자id', '환자ID', 'chart'],
    'date':           ['진료일', '진료 일', '날짜', '일자', '방문일', '시술일', '접수일', '처리일'],
    'name':           ['성명', '이름', '환자명', '환자이름', '환자성명', '고객명'],
    'birth':          ['생년월일', '생년 월일', '생일', 'DOB'],
    'gender':         ['성별', '남녀', '구분'],
    'nationality':    ['국적', '국가', 'nationality', '외국인여부'],
    'doctor':         ['진료의', '진료의사', '담당의', '담당의사', '의사', '원장', '의사명', '담당원장'],
    'order_name':     ['오더명', '오더 명', '시술명', '시술 명', '항목명', '항목', '품목명', '품목', '처방명', '오더', '서비스명'],
    'amount':         ['매출금액', '매출 금액', '금액', '결제금액', '결제 금액', '매출액', '판매금액', '판매 금액', '매출', '결제'],
    'channel':        ['내원경로', '내원 경로', '유입경로', '유입 경로', '경로', '채널', '매체', '광고채널'],
    'original_category': ['분야', '카테고리', '분류', '시술분류', '오더분류'],
    'count':          ['건수', '횟수', '방문수', '방문횟수', '수량'],
}


def _find_column(df: pd.DataFrame, target: str) -> Optional[str]:
    aliases = COLUMN_ALIASES.get(target, [])
    for col in df.columns:
        col_clean = str(col).strip()
        if col_clean in aliases:
            return col
        for alias in aliases:
            if alias in col_clean:
                return col
    return None


def _clean_amount(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
              .str.replace(',', '', regex=False)
              .str.replace('원', '', regex=False)
              .str.replace(' ', '', regex=False)
              .str.strip(),
        errors='coerce'
    ).fillna(0)


def _detect_file_type(df: pd.DataFrame) -> str:
    cols = ' '.join(str(c).strip() for c in df.columns)

    has_order  = any(k in cols for k in ['오더', '시술명', '품목', '처방'])
    has_amount = any(k in cols for k in ['금액', '매출', '결제'])
    has_chart  = any(k in cols for k in ['차트', '환자번호', '환자'])
    has_cat    = any(k in cols for k in ['분야', '카테고리', '분류'])
    has_ch     = any(k in cols for k in ['내원경로', '유입경로', '경로', '채널'])

    if has_order and has_amount and has_chart:
        return '오더판매내역'
    if has_cat and has_amount:
        return '분야별집계'
    if has_ch and has_amount:
        return '내원경로별'
    if has_amount:
        return '매출요약'
    return '알수없음'


def _find_header_row(raw: pd.DataFrame) -> pd.DataFrame:
    """헤더가 첫 번째 행이 아닐 경우 실제 헤더 행을 찾아 재구성."""
    header_keywords = ['차트', '진료', '금액', '매출', '오더', '날짜', '일자', '성명', '이름']
    for i in range(min(10, len(raw))):
        row_str = ' '.join(str(v) for v in raw.iloc[i].values)
        if sum(1 for kw in header_keywords if kw in row_str) >= 2:
            new_df = raw.iloc[i + 1:].copy()
            new_df.columns = raw.iloc[i].values
            return new_df.reset_index(drop=True)
    return raw


def _parse_order_detail(df: pd.DataFrame) -> pd.DataFrame:
    df = _find_header_row(df)
    df = df.dropna(how='all').dropna(axis=1, how='all')

    rename_map = {}
    for target in COLUMN_ALIASES:
        found = _find_column(df, target)
        if found and found not in rename_map:
            rename_map[found] = target

    df = df.rename(columns=rename_map)

    if 'amount' in df.columns:
        df['amount'] = _clean_amount(df['amount'])
        df = df[df['amount'] > 0].copy()

    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'], errors='coerce')

    if 'chart_no' in df.columns:
        df['chart_no'] = df['chart_no'].astype(str).str.strip()

    return df.reset_index(drop=True)


def _parse_category_summary(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how='all').dropna(axis=1, how='all')
    rename_map = {}
    for target in ['original_category', 'amount', 'count']:
        found = _find_column(df, target)
        if found:
            rename_map[found] = target
    df = df.rename(columns=rename_map)
    if 'amount' in df.columns:
        df['amount'] = _clean_amount(df['amount'])
    return df.reset_index(drop=True)


def parse_files(uploaded_files) -> Dict[str, pd.DataFrame]:
    """여러 업로드 파일을 파싱하여 유형별로 모아 반환."""
    parser_map = {
        '오더판매내역': _parse_order_detail,
        '분야별집계':   _parse_category_summary,
        '내원경로별':   _parse_category_summary,
        '매출요약':     _parse_category_summary,
    }

    result: Dict[str, pd.DataFrame] = {}

    for file in uploaded_files:
        file.seek(0)
        raw_bytes = file.read()
        try:
            xl = pd.ExcelFile(io.BytesIO(raw_bytes))
        except Exception:
            continue

        for sheet in xl.sheet_names:
            try:
                raw = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet, header=0)
                raw = raw.dropna(how='all').dropna(axis=1, how='all')
                if raw.empty:
                    continue

                ftype = _detect_file_type(raw)
                parser = parser_map.get(ftype, lambda x: x)
                df = parser(raw.copy())

                if ftype in result:
                    result[ftype] = pd.concat([result[ftype], df], ignore_index=True)
                else:
                    result[ftype] = df

                # 오더판매내역 발견 시 해당 파일의 다른 시트는 무시
                if ftype == '오더판매내역':
                    break
            except Exception:
                continue

    return result
