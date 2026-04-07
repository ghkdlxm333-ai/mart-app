import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드 (바코드 -> ME코드)
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {
            str(r['상품코드']).strip().split('.')[0]: {
                'me': str(r['ME코드']).strip(), 
                'nm': str(r['상품명']).strip()
            } for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])
        }
        
        # 2. 배송코드 로드
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_map = {}
        fallback_map = {} # Q열 타입이 매칭 안될 때 사용할 백업 (이름 기준)
        
        for _, r in df_store.iterrows():
            raw_key = str(r['납품처&타입']).strip()
            clean_key = raw_key.replace(" ", "")
            val = str(r['배송코드']).strip()
            if clean_key and val:
                store_map[clean_key] = val
                # 타입(FLOW/SORTATION 등)을 제외한 순수 이름만 추출하여 백업 맵 생성
                name_only = raw_key.replace("FLOW","").replace("SORTATION","").replace("STOCK","").replace(" ","").strip()
                if name_only:
                    fallback_map[name_only] = val
                    
        return prod_map, store_map, fallback_map, None
    except Exception as e:
        return {}, {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (입고타입 Q열 고정 버전)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, fallback_map, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 1. ordview 데이터 로드 (header=1은 컬럼명이 있는 2번째 줄)
            df_raw = pd.read_excel(uploaded_file, header=1)
            # 낱개수량이 0보다 큰 데이터만 필터링
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [입고타입 및 배송코드 매칭] ---
                # 납품처 (I열 - 보통 9번째)
                raw_place = str(row.get('납품처', '')).strip()
                
                # [지시사항 반영] 입고타입 (Q열 - 17번째 컬럼을 명확히 지정)
                # iloc을 사용하여 컬럼 이름에 상관없이 17번째(인덱스 16) 열 값을 가져옵니다.
                try:
                    raw_type = str(row.iloc[16]).strip() 
                except:
                    raw_type = str(row.get('입고타입', '')).strip()
                
                converted_type = raw_type.replace('
