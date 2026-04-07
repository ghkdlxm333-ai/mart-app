import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 마스터 데이터 로드 부분 (기존과 동일하되 배송코드 매핑 강화) ---
@st.cache_data
def load_master_data(path):
    try:
        # 상품코드 매핑
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {str(r['상품코드']).strip(): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 배송코드 매핑 (Tesco 발주처코드 시트 활용)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        # '납품처&타입'을 키로 하여 '배송코드'와 실제 '배송지명'을 가져옴
        store_map = {str(r['납품처&타입']).strip(): str(r['배송코드']).strip() 
                     for _, r in df_store.iterrows() if pd.notna(r['납품처&타입'])}
        
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

# --- 메인 변환 로직 ---
# (파일 업로드 및 ordview 판별부 이후 코드)

if uploaded_file:
    # 엑셀/HTML/CSV 대응 로더 (앞서 드린 에러 방지 로직 포함 권장)
    try:
        df_raw = pd.read_excel(uploaded_file, header=1)
    except:
        uploaded_file.seek(0)
        df_raw = pd.read_html(uploaded_file)[0]
        df_raw.columns = df_raw.iloc[1]; df_raw = df_raw[2:]

    # 1. 수량 0 제외
    df_raw = df_raw[pd.to_numeric(df_raw['발주수량'], errors='coerce') > 0].copy()

    final_rows = []
    for _, row in df_raw.iterrows():
        # [상품 매칭]
        p_code = str(row.get('상품코드', '')).strip()
        p_info = prod_dict.get(p_code, {'me': '', 'nm': row.get('상품명', '')})

        # [배송코드 매칭 핵심 로직]
        target_center = str(row.get('납품처', '')).strip() # 예: 0906 NEW함안상온물류센터
        in_type = str(row.get('입고타입', '')).strip().replace('HYPER_FLOW', 'FLOW') # 타입 변환
        
        # 마스터와 대조할 키 생성 (예: 0906 NEW함안상온물류센터FLOW)
        lookup_key = f"{target_center}{in_type}"
        shipping_code = store_dict.get(lookup_key, "") # 마스터에서 배송코드 조회

        # [데이터 구성]
        final_rows.append({
            '출고구분': 0,
            '수주일자': datetime.now().strftime('%Y%m%d'),
            '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
            '발주처코드': '81020000',  # 요청사항: 고정값
            '발주처': '홈플러스',      # 요청사항: 고정값
            '배송코드': shipping_code, # 마스터에서 찾아온 코드
            '배송지': target_center,   # 요청사항: 센터정보가 배송지로
            '상품코드': p_info['me'],
            '상품명': p_info['nm'],
            'UNIT수량': int(row.get('발주수량', 0)),
            'UNIT단가': int(row.get('낱개당 단가', 0)),
            '금        액': int(row.get('발주금액', 0)),
            '부  가   세': int(int(row.get('발주금액', 0)) * 0.1),
            'Type': '마트'
        })

    df_final = pd.DataFrame(final_rows)
    st.dataframe(df_final) # 결과 확인
