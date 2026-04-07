import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", layout="wide")

def clean_text(text):
    """공백 제거 및 대문자 변환으로 매칭률 극대화"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 시트 로드 (상품코드 -> ME코드 매핑)
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {}
        for _, r in df_prod.iterrows():
            code_key = clean_text(r.get('상품코드'))
            if code_key:
                prod_map[code_key] = {
                    'me': str(r.get('ME코드', '')).strip(),
                    'nm': str(r.get('상품명', '')).strip()
                }
        
        # 2. Tesco 발주처코드 시트 로드 (D열=Key, E열=Value)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        vlookup_map = {}
        for _, r in df_store.iterrows():
            # D열(index 3): 납품처&타입, E열(index 4): 배송코드
            d_val = str(r.iloc[3]).strip()
            e_val = str(r.iloc[4]).strip()
            if d_val and d_val.lower() != 'nan':
                vlookup_map[clean_text(d_val)] = e_val
        
        return prod_map, vlookup_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (ME코드 & VLOOKUP)")

# 마스터 파일 이름 설정
MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 1. ordview 읽기 (2행부터 데이터)
            df_raw = pd.read_excel(uploaded_file, header=1)
            # 컬럼명 앞뒤 공백 제거 및 정리
            df_raw.columns = [c.strip() if isinstance(c, str) else c for c in df_raw.columns]
            # 수량이 0보다 큰 데이터만 추출
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [로직 1: 배송코드 매칭 - VLOOKUP 방식] ---
                i_val = str(row.get('납품처', '')).strip()
                q_val = str(row.get('입고타입', '')).strip().upper()
                
                # HYPER_FLOW는 FLOW로, SORTATION은 SORTER로 변환하여 마스터 D열과 일치시킴
                q_converted = q_val.replace('HYPER_FLOW', 'FLOW').replace('SORTATION', 'SORTER')
                
                # 매칭 키 생성 (I열
