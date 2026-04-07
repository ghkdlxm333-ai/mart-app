import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (VLOOKUP 방식)", layout="wide")

def clean_text(text):
    """VLOOKUP의 정확도를 높이기 위해 공백만 제거하고 대문자로 변환"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. Tesco 발주처코드 로드 (D열=Key, E열=Value)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        # D열(3번 인덱스): 납품처&타입, E열(4번 인덱스): 배송코드
        vlookup_map = {}
        for _, r in df_store.iterrows():
            d_val = str(r.iloc[3]).strip()
            e_val = str(r.iloc[4]).strip()
            if pd.notna(d_val) and d_val != "nan":
                vlookup_map[clean_text(d_val)] = e_val
        
        return prod_map, vlookup_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (VLOOKUP 로직 적용)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # ordview 읽기 (header=1: 2행부터 데이터 시작)
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [엑셀 VLOOKUP(I&Q, D:E, 2, 0) 로직 구현] ---
                i_val = str(row.get('납품처', '')).strip()
                q_val = str(row.get('입고타입', '')).strip()
                
                # 'HYPER_FLOW'는 마스터의 'FLOW'와 매칭되도록 변환 (VLOOKUP 전처리와 동일)
                q_val_converted = q_val.replace('HYPER_FLOW', 'FLOW').replace('SORTATION', 'SORTER')
                
                # I열 & Q열 결합 (공백 제거)
                lookup_key = clean_text(i_val + q_val_converted)
                
                # 마스터 D:E 범위에서 찾기
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
                # 배송코드+상품코드 기준 합산
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False).agg({'낱개수량': 'sum'})
                
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success("VLOOKUP(I&Q) 방식으로 매칭 및 합산이 완료되었습니다!")
                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 VLOOKUP 결과 다운로드", data=output.getvalue(), file_name=f"HP_VLOOKUP_{datetime.now().strftime('%m%d')}.xlsx")
        except Exception as e:
            st.error(f"오류: {e}")
