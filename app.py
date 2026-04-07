import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", layout="wide")

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
        
        # 2. 배송코드 로드
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_map = {}
        fallback_map = {}
        for _, r in df_store.iterrows():
            raw_key = str(r['납품처&타입']).strip()
            clean_key = raw_key.replace(" ", "")
            val = str(r['배송코드']).strip()
            if clean_key and val:
                store_map[clean_key] = val
                name_only = raw_key.replace("FLOW","").replace("SORTATION","").replace("STOCK","").replace(" ","").strip()
                if name_only:
                    fallback_map[name_only] = val
        return prod_map, store_map, fallback_map, None
    except Exception as e:
        return {}, {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (최종본)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, fallback_map, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    # st.file_uploader의 type에 xlsx만 허용하도록 설정
    uploaded_file = st.file_uploader("ordview 엑셀 파일을 업로드하세요", type=['xlsx'])

    if uploaded_file:
        # 파일 이름의 확장자를 체크하여 .xlsx가 아니면 경고창 표시
        if not uploaded_file.name.lower().endswith('.xlsx'):
            st.error("⚠️ .xlsx 형식의 파일만 업로드 가능합니다. 파일을 확인 후 다시 업로드해주세요.")
            st.stop() # 이후 코드 실행 중단

        try:
            # ordview 로드
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # 납품처 I열
                raw_place = str(row.get('납품처', '')).strip()
                
                # 입고타입 Q열 (17번째 컬럼) 고정
                try:
                    raw_type = str(row.iloc[16]).strip()
                except:
                    raw_type = ""
                
                c_type = raw_type.replace('HYPER_', '')
                m_key = (raw_place + c_type).replace(" ", "")
                
                # 배송코드 매칭
                s_code = store_map.get(m_key, "")
                if not s_code:
                    s_code = fallback_map.get(raw_place.replace(" ", ""), "")
