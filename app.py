import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        # 1. 상품코드 시트 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {str(r['상품코드']).strip(): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 2. Tesco 발주처코드 시트 로드
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_list = []
        for _, r in df_store.iterrows():
            if pd.notna(r['납품처&타입']):
                target_name = str(r['납품처&타입']).strip()
                store_list.append({
                    'full_name': target_name,
                    'clean_name': target_name.replace(' ', ''), # 공백 제거 버전
                    'code': str(r['배송코드']).strip()
                })
        return prod_map, store_list, None
    except Exception as e:
        return {}, [], str(e)

st.title("🛒 홈플러스 수주 자동화 (I열 납품처 기준)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_list, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요 (I열 '납품처' 기준 매칭)", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # ordview 파일 읽기 (두 번째 줄이 헤더인 경우 대비)
            df_raw = pd.read_excel(uploaded_file, header=1)
            
            # 발주수량이 있는 데이터만 추출
            df_raw = df_raw[pd.to_numeric(df_raw['발주수량'], errors='coerce') > 0].copy()

            final_rows = []
            for _, row in df_raw.iterrows():
                # --- [핵심] I열 '납품처' 데이터 가져오기 ---
                # ordview의 I열 헤더 이름이 '납품처'인 경우를 찾습니다.
                raw_delivery_place = str(row.get('납품처', '')).strip()
                in_type = str(row.get('입고타입', '')).strip().replace('HYPER_FLOW', 'FLOW')
                
                # 매칭용 키 생성 (예: 0906 NEW함안상온물류센터 + FLOW)
                lookup_key = f"{raw_delivery_place}{in_type}".replace(' ', '')

                # --- 배송코드 매칭 로직 ---
                shipping_code = ""
                # 마스터 리스트에서 공백을 제거한 이름이 포함되어 있는지 확인
                for item in store_list:
                    if item['clean_name'] == lookup_key:
                        shipping_code = item['code']
                        break
                
                # 만약 못 찾았다면, 이름에 'NEW'가 있거나 없어서 생기는 문제 대응
                if not shipping_code:
                    lookup_key_no_new = lookup_key.replace('NEW', '')
                    for item in store_list:
                        if item['clean_name'].replace('NEW', '') == lookup_key_no_new:
                            shipping_code = item['code']
                            break

                # 상품 정보 찾기
                p_code = str(row.get('상품코드', '')).strip()
                p_info = prod_dict.get(p_code, {'me': '', 'nm': row.get('상품명', '')})

                final_rows.append({
                    '출고구분': 0,
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                    '발주처코드': '81020000',
                    '발주처': '홈플러스',
                    '배송코드': shipping_code,
                    '배송지': raw_delivery_place,
                    '상품코드': p_info['me'],
                    '상품명': p_info['nm'],
                    'UNIT수량': int(float(row.get('발주수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                    '금        액': int(float(row.get('발주금액', 0))),
                    '부  가   세': int(float(row.get('발주금액', 0)) * 0.1),
                    'Type': '마트'
                })

            if final_rows:
                df_final = pd.DataFrame(final_rows)
                st.success(f"변환 완료! (총 {len(df_final)}건)")
                
                # 배송코드 매칭 실패 항목 경고
                missing_codes = df_final[df_final['배송코드'] == ""]
                if not missing_codes.empty:
                    st.warning(f"⚠️ {len(missing_codes)}건의 배송코드를 찾지 못했습니다. (배송지: {missing_codes['배송지'].unique()})")

                st.dataframe(df_final)

                # 엑셀 다운로드 생성
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                
                st.download_button(
                    label="📥 변환된 엑셀 다운로드",
                    data=output.getvalue(),
                    file_name=f"Homeplus_Order_{datetime.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"오류 발생: {e}")
