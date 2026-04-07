import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", layout="wide")

def clean_text(text):
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 시트 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        p_map = {}
        for _, r in df_prod.iterrows():
            k = clean_text(r.get('상품코드'))
            if k:
                p_map[k] = {'me': str(r.get('ME코드', '')).strip(), 'nm': str(r.get('상품명', '')).strip()}
        
        # 2. Tesco 발주처코드 시트 로드
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        s_map = {}
        for _, r in df_store.iterrows():
            d_val = str(r.iloc[3]).strip() # D열
            e_val = str(r.iloc[4]).strip() # E열
            if d_val and d_val.lower() != 'nan':
                s_map[clean_text(d_val)] = e_val
        return p_map, s_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (최종 수정본)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw.columns = [c.strip() if isinstance(c, str) else c for c in df_raw.columns]
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # 배송코드 매칭 키 생성
                i_val = str(row.get('납품처', '')).strip()
                q_val = str(row.get('입고타입', '')).strip().upper()
                
                # 핵심 변환 로직
                q_converted = q_val.replace('HYPER_FLOW', 'FLOW').replace('SORTATION', 'SORTER')
                lookup_key = clean_text(i_val + q_converted)
                shipping_code = store_map.get(lookup_key, "")

                # 상품코드 -> ME코드 변환
                p_code_raw = clean_text(row.get('상품코드', ''))
                p_info = prod_dict.get(p_code_raw)
                
                if p_info:
                    final_p_code = p_info['me']
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
                    '상품코드': final_p_code,
                    '상품명': final_p_name,
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
                
                cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']
                df_final = df_final[cols]
                
                st.success("매칭 및 ME코드 변환 완료")
                
                missing = df_final[df_final['배송코드'] == ""]
                if not missing.empty:
                    st.error(f"❌ 배송코드 매칭 실패: {missing['배송지'].unique()}")
                
                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 결과 다운로드", data=output.getvalue(), file_name=f"HP_Final_{datetime.now().strftime('%m%d')}.xlsx")
            else:
                st.info("데이터가 없습니다.")
        except Exception as e:
            st.error(f"오류: {e}")
