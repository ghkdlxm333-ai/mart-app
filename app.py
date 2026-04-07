import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- 1. 기본 설정 및 마스터 파일 경로 (GitHub RAW 링크로 교체 권장) ---
# 예: "https://raw.githubusercontent.com/사용자명/리포지토리/main/Tesco_서식파일_업데이트용.xlsx"
MASTER_FILE_PATH = "Tesco_서식파일_업데이트용.xlsx" 

st.set_page_config(page_title="마트 수주업로드 시스템", page_icon="🛒", layout="wide")

# --- 2. 데이터 클리닝 함수 ---
def clean_val(val):
    if pd.isna(val): return ""
    return str(val).strip().replace('.0', '')

# --- 3. 마스터 데이터 로드 (캐싱 처리) ---
@st.cache_data
def load_master_data(path):
    try:
        xls = pd.ExcelFile(path)
        # 상품코드 매핑: [바코드/상품코드] -> ME코드
        df_prod = pd.read_excel(xls, '상품코드', dtype=str)
        prod_map = {}
        for _, r in df_prod.iterrows():
            code = clean_val(r.get('바코드')) or clean_val(r.get('상품코드'))
            if code:
                prod_map[code] = {'me_code': clean_val(r.get('ME코드')), 'name': clean_val(r.get('상품명'))}

        # 발주처코드 매핑: [납품처&타입] -> [발주처코드, 배송코드, 배송지]
        df_store = pd.read_excel(xls, 'Tesco 발주처코드', dtype=str)
        store_map = {}
        for _, r in df_store.iterrows():
            key = clean_val(r.get('납품처&타입'))
            if key:
                store_map[key] = {
                    '발주처코드': clean_val(r.get('발주처코드')),
                    '발주처': clean_val(r.get('업체명')),
                    '배송코드': clean_val(r.get('배송코드')),
                    '배송지': clean_val(r.get('배송지'))
                }
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

# --- 4. 메인 로직 ---
st.title("🛒 마트(홈플러스) 수주 자동화 시스템")
st.info("ordview 파일을 업로드하면 상품코드(ME)와 배송코드를 자동으로 매칭합니다.")

prod_dict, store_dict, error = load_master_data(MASTER_FILE_PATH)

if error:
    st.error(f"마스터 파일을 불러오지 못했습니다: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        # 파일 판별 및 로드
        if "ordview" in uploaded_file.name.lower():
            st.success("✅ 홈플러스 ordview 파일이 확인되었습니다.")
            df_raw = pd.read_excel(uploaded_file, header=1) # ordview 특성상 header 위치 조정 필요
            
            # 수량 0 제외
            df_raw = df_raw[df_raw['발주수량'] > 0].copy()
            
            final_data = []
            for _, row in df_raw.iterrows():
                # 1) 상품 매칭
                raw_prod_code = clean_val(row.get('상품코드'))
                prod_info = prod_dict.get(raw_prod_code, {'me_code': '', 'name': row.get('상품명')})

                # 2) 배송코드 매칭 (HYPER_FLOW -> FLOW 변환 로직 포함)
                raw_store_name = clean_val(row.get('납품처'))
                raw_type = clean_val(row.get('입고타입')).replace('HYPER_FLOW', 'FLOW')
                mapping_key = f"{raw_store_name}{raw_type}" # 예: "0982 안성ADC물류센터FLOW"
                
                store_info = store_dict.get(mapping_key, {'발주처코드': '', '발주처': '', '배송코드': '', '배송지': ''})

                # 3) 최종 양식 배치
                final_data.append({
                    '출고구분': 0,
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': clean_val(row.get('납품일자')).replace('-', '')[:8],
                    '발주처코드': store_info['발주처코드'],
                    '발주처': store_info['발주처'] or row.get('납품처'),
                    '배송코드': store_info['배송코드'],
                    '배송지': store_info['배송지'] or row.get('배송처'),
                    '상품코드': prod_info['me_code'],
                    '상품명': prod_info['name'],
                    'UNIT수량': int(row.get('발주수량', 0)),
                    'UNIT단가': int(row.get('낱개당 단가', 0)),
                    '금        액': int(row.get('발주금액', 0)),
                    '부  가   세': int(int(row.get('발주금액', 0)) * 0.1),
                    'Type': '마트'
                })

            df_res = pd.DataFrame(final_data)
            
            # 컬럼 순서 고정 및 출력
            cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']
            df_res = df_res[cols]

            st.subheader("📊 변환 결과 미리보기")
            st.dataframe(df_res, use_container_width=True)

            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False, sheet_name='서식업로드')
            
            st.download_button(
                label="📥 통합 수주 양식 다운로드",
                data=output.getvalue(),
                file_name=f"Homeplus_Order_{datetime.now().strftime('%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("파일명에 'ordview'가 포함되어 있지 않습니다. 파일을 확인해 주세요.")
