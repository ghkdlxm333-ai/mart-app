import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (공백 예외 처리)", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {}
        for _, r in df_prod.iterrows():
            barcode = str(r['상품코드']).strip().split('.')[0]
            if barcode:
                prod_map[barcode] = {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()}
        
        # 2. 배송코드 로드 (공백 제거 로직 적용)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_map = {}
        for _, r in df_store.iterrows():
            # [수정] 매칭 키에서 모든 공백을 제거하여 저장
            raw_key = str(r['납품처&타입']).strip()
            clean_key = raw_key.replace(" ", "") 
            val = str(r['배송코드']).strip()
            if clean_key and val:
                store_map[clean_key] = val
                
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (함안센터 공백 해결 버전)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [배송코드 매칭 로직 개선] ---
                # 1. ordview 데이터 추출 및 공백 제거
                raw_place = str(row.get('납품처', '')).strip()
                raw_type = str(row.get('입고타입', '')).strip()
                
                # 2. 입고타입 변환 (HYPER_ 제거)
                converted_type = raw_type.replace('HYPER_', '')
                
                # 3. [수정] 매칭용 키 생성 시 모든 공백 제거
                # 예: "0906 NEW함안상온물
