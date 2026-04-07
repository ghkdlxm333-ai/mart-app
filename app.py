import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (ME코드 매칭)", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드 (C열의 ME코드를 추출하기 위해 dtype=str 설정)
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        # 상품코드(바코드)를 키로, ME코드를 값으로 매칭하는 맵 생성
        # 마스터 파일의 '상품코드'열과 'ME코드'열(C열)을 사용합니다.
        prod_map = {
            str(r['상품코드']).strip(): {
                'me': str(r['ME코드']).strip(), 
                'nm': str(r['상품명']).strip()
            } 
            for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])
        }
        
        # 2. 배송코드 로드
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_list = []
        for _, r in df_store.iterrows():
            name_val = str(r['납품처&타입']).strip() if pd.notna(r['납품처&타입']) else ""
            code_val = str(r['배송코드']).strip() if pd.notna(r['배송코드']) else ""
            if name_val and code_val:
                store_list.append({
                    'name': name_val,
                    'num': name_val[:4], 
                    'code': code_val
                })
        return prod_map, store_list, None
    except Exception as e:
        return {}, [], str(e)

st.title("🛒 홈플러스 수주 자동화 (ME코드 변환 버전)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_list, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # ordview 읽기 (낱개수량이 0보다 큰 데이터만)
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # 1. 배송지 정보 추출
                raw_place = str(row.get('납품처', '')).strip()
                place_num = raw_place[:4]
                in_type = str(row.get('입고타입', '')).strip().replace('HYPER_', '')

                # 2. 배송코드 매칭
                shipping_code = ""
                for item in store_list:
                    if item['num'] == place_num and in_type in item['name']:
                        shipping_code = item['code']
                        break
                if not shipping_code:
                    for item in store_list:
                        if item['num'] == place_num:
                            shipping_code = item['code']
                            break

                # 3. [핵심] 상품코드 -> ME코드 매칭
                p_code = str(row.get('상품코드', '')).strip()
                # 마스터 파일에 코드가 있으면 ME코드로 치환, 없으면 기존 코드 유지
                p_info = prod_dict.get(p_code)
                
                if p_info:
