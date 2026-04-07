import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (D:E 정밀 매칭)", layout="wide")

def clean_text(text):
    """모든 공백을 제거하고 대문자로 변환하여 비교"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. Tesco 발주처코드 로드 (D열과 E열 매칭 테이블 생성)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        # D열(인덱스 3) '납품처&타입', E열(인덱스 4) '배송코드'
        store_map = {}
        for _, r in df_store.iterrows():
            d_val = str(r.iloc[3]).strip() # D열
            e_val = str(r.iloc[4]).strip() # E열
            if pd.notna(d_val) and d_val != "nan":
                store_map[clean_text(d_val)] = e_val
        
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (D:E열 직접 매칭)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [매칭 키 생성 로직] ---
                raw_place = str(row.get('납품처', '')).strip()
                raw_type = str(row.get('입고타입', '')).strip().upper()
                
                # 1. 타입 변환 (마스터 파일 D열 문구와 일치화)
                if "HYPER_FLOW" in raw_type:
                    converted_type = "FLOW"
                elif "SORT" in raw_type:
                    converted_type = "SORTER" # 마스터 D12셀의 'SORTER'에 맞춤
                else:
                    converted_type = raw_type
                
                # 2. 결합 키 (납품처 + 변환된 타입)
                search_key = clean_text(raw_place + converted_type)
                
                # 3. 마스터 D열(key)에서 E열(value) 찾기
                shipping_code = store_map.get(search_key, "")

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
                    '배송지': raw_place,
                    '상품코드': p_info['me'],
                    '상품명': p_info['nm'],
                    '낱개수량': int(float(row.get('낱개수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                    'Type': '마트'
                })

            if temp_rows:
                df_temp = pd.DataFrame(temp_rows)
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False).agg({'낱개수량': 'sum'})
                
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success("D11:E12 데이터 기반으로 매칭 및 합산이 완료되었습니다.")
                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 결과 엑셀 다운로드", data=output.getvalue(), file_name=f"HP_Final_Fixed_{datetime.now().strftime('%m%d')}.xlsx")
        except Exception as e:
            st.error(f"오류: {e}")
