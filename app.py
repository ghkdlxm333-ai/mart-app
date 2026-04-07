import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (Q열 타입 매칭)", layout="wide")

def clean_text(text):
    """공백 제거 및 문자열 정규화"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 시트 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. Tesco 발주처코드 시트 로드 (매칭 키 생성)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_map = {}
        for _, r in df_store.iterrows():
            raw_key = str(r['납품처&타입']).strip()
            if pd.notna(raw_key) and raw_key != "nan":
                # 마스터의 '납품처&타입'에서 공백 제거 후 키로 저장
                match_key = clean_text(raw_key)
                store_map[match_key] = str(r['배송코드']).strip()
        
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (Q열 & HYPER_FLOW 변환)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # ordview 파일 읽기 (헤더 위치 주의)
            df_raw = pd.read_excel(uploaded_file, header=1)
            # 낱개수량 0보다 큰 데이터만 필터링
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # --- 배송코드 매칭 로직 ---
                # 1. 납품처 가져오기 (보통 I열)
                raw_place = str(row.get('납품처', '')).strip()
                
                # 2. 입고타입 가져오기 (Q열 명시적 지정)
                # 만약 .get('입고타입')이 안되면 iloc로 Q열(16번 인덱스)을 직접 지정할 수도 있습니다.
                raw_type = str(row.get('입고타입', '')).strip().upper()
                
                # 3. 'HYPER_FLOW' -> 'FLOW' 변환
                converted_type = raw_type.replace('HYPER_FLOW', 'FLOW')
                
                # 4. 결합 키 생성 (납품처 + 변환된 타입) 및 공백 제거
                combined_key = clean_text(raw_place + converted_type)
                
                # 5. 매칭
                shipping_code = store_map.get(combined_key, "")
                
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
                
                # 동일 배송코드/상품코드 합산
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False)['낱개수량'].sum()
                
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                # 최종 열 순서 정리
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success("데이터 변환 완료!")
                
                # 배송코드 매칭 실패 시 경고
                missing_codes = df_final[df_final['배송코드'] == ""]
                if not missing_codes.empty:
                    st.warning(f"⚠️ 배송코드를 찾지 못한 항목 (납품처 확인 필요): {missing_codes['배송지'].unique()}")

                st.dataframe(df_final)

                # 엑셀 다운로드 생성
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 결과 엑셀 다운로드", data=output.getvalue(), file_name=f"Homeplus_Order_{datetime.now().strftime('%Y%m%d')}.xlsx")
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
else:
    st.error(f"마스터 파일 로드 실패: {error}")
