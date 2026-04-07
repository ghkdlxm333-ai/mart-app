import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (정밀 결합 매칭)", layout="wide")

def clean_text(text):
    """공백 및 특수문자 제거하여 비교용 텍스트 생성"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text))

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. 배송코드 로드 (납품처&타입 기준 공백제거 매칭 테이블 생성)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_map = {}
        for _, r in df_store.iterrows():
            raw_key = str(r['납품처&타입']).strip()
            if pd.notna(raw_key) and raw_key != "nan":
                # "0906 NEW함안상온물류센터FLOW" -> "0906NEW함안상온물류센터FLOW"
                match_key = clean_text(raw_key)
                store_map[match_key] = str(r['배송코드']).strip()
        
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (결합 키 정밀 매칭)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # ordview 읽기
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- 배송코드 매칭 핵심 로직 ---
                raw_place = str(row.get('납품처', '')).strip()
                raw_type = str(row.get('입고타입', '')).strip()
                
                # 1. [납품처 + 입고타입] 결합 및 공백 제거
                combined_key = clean_text(raw_place + raw_type)
                
                # 2. 마스터 맵에서 배송코드 조회
                shipping_code = store_map.get(combined_key, "")
                
                # 3. (보조 로직) 만약 결합 키로 못 찾았다면, HYPER_ 제거 후 재시도
                if not shipping_code and "HYPER_" in raw_type:
                    alt_key = clean_text(raw_place + raw_type.replace("HYPER_", ""))
                    shipping_code = store_map.get(alt_key, "")

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
                    'Type': '마트',
                    'original_key': combined_key # 디버깅용 키
                })

            if temp_rows:
                df_temp = pd.DataFrame(temp_rows)
                
                # 합산: 배송코드와 상품코드가 같으면 낱개수량 합산
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False)['낱개수량'].sum()
                
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success("배송지+타입 정밀 매칭 및 합산이 완료되었습니다.")
                
                # 배송코드 누락 안내
                missing = df_final[df_final['배송코드'] == ""]
                if not missing.empty:
                    st.warning(f"⚠️ 배송코드를 찾지 못한 항목이 있습니다 (배송지 확인 필요): {missing['배송지'].unique()}")

                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 최종 결과 다운로드", data=output.getvalue(), file_name=f"HP_Final_Matched_{datetime.now().strftime('%m%d')}.xlsx")
        except Exception as e:
            st.error(f"오류 발생: {e}")
else:
    st.error(f"마스터 파일을 읽을 수 없습니다: {error}")
