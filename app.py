import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", layout="wide")

def clean_text(text):
    """모든 공백을 제거하고 대문자로 변환하여 매칭 정확도 극대화"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드 (상품코드 -> ME코드 매핑)
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. Tesco 발주처코드 로드 (D열=Key, E열=Value)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        vlookup_map = {}
        for _, r in df_store.iterrows():
            d_val = str(r.iloc[3]).strip() # D열 (납품처&타입)
            e_val = str(r.iloc[4]).strip() # E열 (배송코드)
            if d_val and d_val.lower() != 'nan':
                vlookup_map[clean_text(d_val)] = e_val
        
        return prod_map, vlookup_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (최종 보정본)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 1. ordview 읽기 (2행부터 데이터)
            df_raw = pd.read_excel(uploaded_file, header=1)
            # 컬럼명 정리
            df_raw.columns = [c.strip() if isinstance(c, str) else c for c in df_raw.columns]
            # 수량이 있는 행만 필터링
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [배송코드 매칭] ---
                i_val = str(row.get('납품처', '')).strip()
                q_val = str(row.get('입고타입', '')).strip().upper()
                
                # HYPER_FLOW -> FLOW, SORTATION -> SORTER, SINGLE 유지
                q_converted = q_val.replace('HYPER_FLOW', 'FLOW').replace('SORTATION', 'SORTER')
                
                # I+Q 결합 키 생성 (마스터 D열과 대조용)
                lookup_key = clean_text(i_val + q_converted)
                shipping_code = vlookup_map.get(lookup_key, "")

                # --- [상품코드 -> ME코드 변환] ---
                p_code_raw = clean_text(row.get('상품코드', ''))
                p_info = prod_dict.get(p_code_raw)
                
                if p_info:
                    final_p_code = p_info['me'] # ME코드 할당
                    final_p_name = p_info['nm']
                else:
                    final_p_code = p_code_raw
                    final_p_name = row.get('상품명', '')

                temp_rows.append({
                    '출고구분': 0,
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                    '발주처코드': '81020000',
                    '발주처': '홈플러스',
                    '배송코드': shipping_code,
                    '배송지': i_val,
                    '상품코드': final_p_code, # 변환된 ME코드 삽입
                    '상품명': final_p_name,
                    '낱개수량': int(float(row.get('낱개수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                    'Type': '마트'
                })

            if temp_rows:
                df_temp = pd.DataFrame(temp_rows)
                # 동일 배송코드 및 ME코드 기준 합산
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False).agg({'낱개수량': 'sum'})
                
                # 계산 및 컬럼 정리
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success("데이터 변환 완료 (ME코드 적용)")
                
                # 누락 확인
                missing = df_final[df_final['배송코드'] == ""]
                if not missing.empty:
                    st.error(f"❌ 배송코드 매칭 실패: {missing['배송지'].unique()}")
