import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (완벽 매칭)", layout="wide")

def clean_text(text):
    """VLOOKUP 정확도를 위해 공백 제거 및 대문자 변환"""
    if pd.isna(text): return ""
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
            # D열(index 3)과 E열(index 4) 정보를 가져옴
            d_val = str(r.iloc[3]).strip()
            e_val = str(r.iloc[4]).strip()
            if pd.notna(d_val) and d_val != "nan":
                vlookup_map[clean_text(d_val)] = e_val
        
        return prod_map, vlookup_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (함안 코드 매칭 수정본)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [엑셀 VLOOKUP 로직의 핵심 보정] ---
                i_val = str(row.get('납품처', '')).strip()
                q_val = str(row.get('입고타입', '')).strip().upper()
                
                # 마스터 D열(0906...FLOW / SORTER) 형식에 맞게 강제 변환
                if "FLOW" in q_val:
                    q_converted = "FLOW"
                elif "SORT" in q_val:
                    q_converted = "SORTER"
                elif "SINGLE" in q_val:
                    q_converted = "SINGLE"
                else:
                    q_converted = q_val
                
                # I열 & 변환된 Q열 결합
                lookup_key = clean_text(i_val + q_converted)
                
                # 매칭 시도
                shipping_code = vlookup_map.get(lookup_key, "")

                # 만약 실패 시, 'NEW' 단어가 있고 없고의 차이 보정 (최후의 수단)
                if not shipping_code:
                    alt_key = lookup_key.replace("NEW", "")
                    shipping_code = vlookup_map.get(alt_key, "")

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
                
                st.success("매칭 프로세스 완료")
                
                # 미매칭 항목 디버깅용 정보 표시
                missing = df_final[df_final['배송코드'] == ""]
                if not missing.empty:
                    st.error(f"❌ 배송코드 매칭 실패: {missing['배송지'].unique()}")
                    st.info("마스터 파일 D열의 명칭과 ordview의 [납품처+입고타입]이 일치하는지 확인해주세요.")

                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 결과 다운로드", data=output.getvalue(), file_name=f"HP_Final_{datetime.now().strftime('%m%d')}.xlsx")
        except Exception as e:
            st.error(f"오류 발생: {e}")
