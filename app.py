import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (HYPER_FLOW 완벽 변환)", layout="wide")

def clean_text(text):
    """모든 공백을 제거하고 대문자로 변환하여 매칭 정확도 극대화"""
    if pd.isna(text): return ""
    # 양끝 공백 제거 후, 중간의 모든 공백 제거
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. Tesco 발주처코드 로드 (D열=납품처&타입, E열=배송코드)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        vlookup_map = {}
        for _, r in df_store.iterrows():
            # D열(index 3) 명칭을 Key로, E열(index 4) 코드를 Value로 저장
            d_val = str(r.iloc[3]).strip()
            e_val = str(r.iloc[4]).strip()
            if pd.notna(d_val) and d_val != "nan":
                # 마스터 파일 D열의 "0906 NEW함안...FLOW" 공백 제거 후 저장
                vlookup_map[clean_text(d_val)] = e_val
        
        return prod_map, vlookup_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (HYPER_FLOW -> FLOW 변환 완료)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # ordview 읽기
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [핵심: VLOOKUP I&Q 로직] ---
                i_val = str(row.get('납품처', '')).strip() # I열
                q_val = str(row.get('입고타입', '')).strip().upper() # Q열
                
                # 지시하신 대로 HYPER_FLOW는 무조건 FLOW로 변환
                # SORTATION은 마스터의 SORTER와 매칭되도록 보정
                q_converted = q_val.replace('HYPER_FLOW', 'FLOW').replace('SORTATION', 'SORTER')
                
                # [납품처 + 변환된 타입] 결합 후 공백 제거 (VLOOKUP용 Key)
                # 예: 0906NEW함안상온물류센터 + FLOW
                lookup_key = clean_text(i_val + q_converted)
                
                # 마스터 D열에서 해당 키로 배송코드(E열) 찾기
                shipping_code = vlookup_map.get(lookup_key, "")

                # 상품 정보 매칭
                p_code = clean_text(row.get('상품코드', ''))
                p_info = prod_dict.get(p_code, {'me': p_code, 'nm': row.get('상품명', '')})

                temp_rows.append({
                    '출고구분': 0,
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                    '발주처코드': '81020000',
                    '발주처': '홈플러스',
                    '배송코드': shipping_code,
                    '배송지': i_val,
                    '상품코드': p_info['me'],
                    '상품명': p_info['nm'],
                    '낱개수량': int(float(row.get('낱개수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                    'Type': '마트'
                })

            if temp_rows:
                df_temp = pd.DataFrame(temp_rows)
                # 배송코드, 상품코드 등이 같으면 낱개수량 합산
