import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="홈플러스 수주 자동화", page_icon="🛒", layout="wide")

# --- 2. 마스터 데이터 로드 함수 ---
@st.cache_data
def load_master_data(path):
    try:
        # 상품코드 시트: [상품코드] -> [ME코드, 상품명]
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {str(r['상품코드']).strip(): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # Tesco 발주처코드 시트: [납품처&타입] -> [배송코드]
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        # 공백 제거를 통해 매칭 확률을 높임
        store_map = {str(r['납품처&타입']).strip().replace(' ', ''): str(r['배송코드']).strip() 
                     for _, r in df_store.iterrows() if pd.notna(r['납품처&타입'])}
        
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

# --- 3. 메인 화면 구성 ---
st.title("🛒 홈플러스(Tesco) 수주 자동화 시스템")

# 마스터 파일 경로 (같은 폴더에 있어야 함)
MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_dict, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    # --- 핵심: 변수 할당 (여기서 uploaded_file이 정의됩니다) ---
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        if "ordview" in uploaded_file.name.lower():
            try:
                # 파일 읽기 (HTML/Excel 대응)
                try:
                    df_raw = pd.read_excel(uploaded_file, header=1)
                except:
                    uploaded_file.seek(0)
                    df_html = pd.read_html(uploaded_file)
                    df_raw = df_html[0]
                    df_raw.columns = df_raw.iloc[1]; df_raw = df_raw[2:]

                # 발주수량 0보다 큰 것만 필터링
                df_raw = df_raw[pd.to_numeric(df_raw['발주수량'], errors='coerce') > 0].copy()

                final_rows = []
                for _, row in df_raw.iterrows():
                    # 1) 상품 매칭
                    p_code = str(row.get('상품코드', '')).strip()
                    p_info = prod_dict.get(p_code, {'me': '', 'nm': row.get('상품명', '')})

                    # 2) 배송지 및 배송코드 매칭 로직
                    # '납품처' 정보(예: 0906 NEW함안상온물류센터)를 배송지로 사용
                    target_delivery_place = str(row.get('납품처', '')).strip()
                    
                    # 입고타입에서 HYPER_FLOW를 FLOW로 변경
                    in_type = str(row.get('입고타입', '')).strip().replace('HYPER_FLOW', 'FLOW')
                    
                    # 매칭 키 생성 (공백 제거하여 비교)
                    lookup_key = f"{target_delivery_place}{in_type}".replace(' ', '')
                    shipping_code = store_dict.get(lookup_key, "")

                    # 3) 데이터 생성
                    final_rows.append({
                        '출고구분': 0,
                        '수주일자': datetime.now().strftime('%Y%m%d'),
                        '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                        '발주처코드': '81020000',      # 요청: 고정값
                        '발주처': '홈플러스',          # 요청: 고정값
                        '배송코드': shipping_code,     # 매칭 결과
                        '배송지': target_delivery_place, # 요청: 납품처 정보가 배송지로
                        '상품코드': p_info['me'],
                        '상품명': p_info['nm'],
                        'UNIT수량': int(float(row.get('발주수량', 0))),
                        'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                        '금        액': int(float(row.get('발주금액', 0))),
                        '부  가   세': int(float(row.get('발주금액', 0)) * 0.1),
                        'Type': '마트'
                    })

                if not final_rows:
                    st.warning("
