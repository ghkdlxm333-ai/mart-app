import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (VLOOKUP & ME코드)", layout="wide")

def clean_text(text):
    """모든 공백을 제거하고 대문자로 변환하여 매칭 정확도 극대화"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 시트 로드 (상품코드 -> ME코드 매핑)
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        # 마스터의 '상품코드'를 키로, 'ME코드'와 '상품명'을 저장
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. Tesco 발주처코드 시트 로드 (D열=Key, E열=Value)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        vlookup_map = {}
        for _, r in df_store.iterrows():
            d_val = str(r.iloc[3]).strip() # D열: 납품처&타입
            e_val = str(r.iloc[4]).strip() # E열: 배송코드
            if d_val and d_val.lower() != 'nan':
                vlookup_map[clean_text(d_val)] = e_val
        
        return prod_map, vlookup_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (완성본)")

# 파일 경로 (서버 환경에 맞춰 확인 필요)
MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 1. ordview 파일 읽기
            df_raw = pd.read_excel(uploaded_file, header=1)
            # 컬럼명 앞뒤 공백 제거
            df_raw.columns = [c.strip() if isinstance(c, str) else c for c in df_raw.columns]
            # 낱개수량 0보다 큰 것만 필터링
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [로직 1: 배송코드 매칭] ---
                i_val = str(row.get('납품처', '')).strip() # I열
                q_val = str(row.get('입고타입', '')).strip().upper() # Q열
                
                # 변환: HYPER_FLOW -> FLOW / SORTATION -> SORTER
                q_converted = q_val.replace('HYPER_FLOW', 'FLOW').replace('SORTATION', 'SORTER')
                
                # 엑셀 VLOOKUP(I&Q, D:E, 2, 0)과 동일한 키 생성
                lookup_key = clean_text(i_val + q_converted)
                shipping_code = vlookup_map.get(lookup_key, "")

                # --- [로직 2: 상품코드 -> ME코드 변환]
