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
        # 상품코드 시트
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {str(r['상품코드']).strip(): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # Tesco 발주처코드 시트
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        
        # [중요] 마스터 파일의 키값에서도 공백을 완전히 제거하여 저장합니다.
        store_map = {}
        for _, r in df_store.iterrows():
            if pd.notna(r['납품처&타입']):
                clean_key = str(r['납품처&타입']).replace(' ', '').strip()
                store_map[clean_key] = str(r['배송코드']).strip()
        
        return prod_map, store_map, None
    except Exception as e:
        return {}, {}, str(e)

# --- 3. 메인 화면 ---
st.title("🛒 홈플러스(Tesco) 수주 자동화")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_dict, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 엑셀 읽기
            try:
                df_raw = pd.read_excel(uploaded_file, header=1)
            except:
                uploaded_file.seek(0)
                df_raw = pd.read_html(uploaded_file)[0]
                df_raw.columns = df_raw.iloc[1]; df_raw = df_raw[2:]

            df_raw = df_raw[pd.to_numeric(df_raw['발주수량'], errors='coerce') > 0].copy()

            final_rows = []
            for _, row in df_raw.iterrows():
                # 상품 정보
                p_code = str(row.get('상품코드', '')).strip()
                p_info = prod_dict.get(p_code, {'me': '', 'nm': row.get('상품명', '')})

                # 배송 정보 (공백 제거 매칭)
                delivery_place = str(row.get('납품처', '')).strip()
                in_type = str(row.get('입고타입', '')).strip().replace('HYPER_FLOW', 'FLOW')
                
                # 매칭용 키 생성 (모든 공백 제거)
                lookup_key = f"{delivery_place}{in_type}".replace(' ', '')
                shipping_code = store_dict.get(lookup_key, "")

                # 데이터 구성
                final_rows.append({
                    '출고구분': 0,
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                    '발주처코드': '81020000', # 고정
                    '발주처': '홈플러스',     # 고정
                    '배송코드': shipping_code,
                    '배송지': delivery_place,
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
                
                # 매칭 실패 확인용 (배송코드가 빈 값인 경우 경고)
                fail_count = df_final['배송코드'].eq("").sum()
                if fail_count > 0:
                    st.warning(f"⚠️ {fail_count}건의 배송코드를 찾지 못했습니다. 마스터 파일의 이름을 확인하세요.")

                st.dataframe(df_final)

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
