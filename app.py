import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (최종 완성본)", layout="wide")

def clean_text(text):
    """모든 공백 제거 및 대문자 변환 (VLOOKUP 정확도용)"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 시트 로드 (상품코드 -> ME코드 매핑)
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        # 마스터의 '상품코드'를 키로, 'ME코드'와 '상품명'을 가져옴
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

st.title("🛒 홈플러스 수주 자동화 (ME코드 & VLOOKUP 적용)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # ordview 읽기 (2행부터 데이터 시작)
            df_raw = pd.read_excel(uploaded_file, header=1)
            # 컬럼명 공백 제거
            df_raw.columns = [c.strip() if isinstance(c, str) else c for c in df_raw.columns]
            # 낱개수량이 있는 행만 필터링
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [1. 배송코드 매칭 로직] ---
                i_val = str(row.get('납품처', '')).strip() # I열
                q_val = str(row.get('입고타입', '')).strip().upper() # Q열
                
                # 변환 규칙 적용
                q_converted = q_val.replace('HYPER_FLOW', 'FLOW').replace('SORTATION', 'SORTER')
                
                # 결합 키 생성 및 마스터 D:E 매칭
                lookup_key = clean_text(i_val + q_converted)
                shipping_code = vlookup_map.get(lookup_key, "")

                # --- [2. 상품코드 -> ME코드 변환 로직] ---
                p_code_raw = clean_text(row.get('상품코드', ''))
                p_info = prod_dict.get(p_code_raw)
                
                if p_info:
                    final_p_code = p_info['me'] # 마스터의 'ME코드' 사용
                    final_p_name = p_info['nm'] # 마스터의 '상품명' 사용
                else:
                    final_p_
