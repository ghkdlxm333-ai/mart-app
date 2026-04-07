import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (최종 보정)", layout="wide")

def clean_text(text):
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. 배송코드 로드
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_list = []
        for _, r in df_store.iterrows():
            name_val = str(r['납품처&타입']).strip()
            code_val = str(r['배송코드']).strip()
            if pd.notna(name_val) and pd.notna(code_val):
                store_list.append({
                    'full_name': clean_text(name_val),
                    'num': clean_text(name_val)[:4], # '0906'
                    'code': code_val
                })
        return prod_map, store_list, None
    except Exception as e:
        return {}, [], str(e)

st.title("🛒 홈플러스 수주 자동화 (함안/안성 완벽 구분)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_list, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- [강화된 매칭 로직] ---
                raw_place = str(row.get('납품처', '')).strip()
                place_num = clean_text(raw_place)[:4] # 0906 또는 0982
                
                # Q열 입고타입 파악 및 변환
                raw_type = str(row.get('입고타입', '')).strip().upper()
                if "SINGLE" in raw_type: target_type = "SINGLE"
                elif "SORT" in raw_type: target_type = "SORT" # SORTATION, SORTER 모두 포함
                else: target_type = "FLOW" # HYPER_FLOW 포함 기본값
                
                shipping_code = ""
                # 1단계: 번호(0906)가 맞고, 타입(FLOW/SORT/SINGLE) 키워드가 마스터 명칭에 포함되는지 확인
                for item in store_list:
                    if item['num'] == place_num and target_type in item['full_name']:
                        shipping_code = item['code']
                        break
                
                # 2단계: 실패 시 번호로만 매칭
                if not shipping_code:
                    for item in store_list:
                        if item['num'] == place_num:
                            shipping_code = item['code']
                            break

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
                df_final = df_temp.groupby(group_cols, as_index=False)['낱개수량'].sum()
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success("함안 FLOW/SORTATION 구분이 완료되었습니다.")
                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 최종 엑셀 다운로드", data=output.getvalue(), file_name=f"HP_Final_{datetime.now().strftime('%m%d')}.xlsx")
        except Exception as e:
            st.error(f"오류: {e}")
