import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        # 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {str(r['상품코드']).strip(): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 배송코드 로드 (유연한 매칭을 위해 리스트로 보관)
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_list = []
        for _, r in df_store.iterrows():
            if pd.notna(r['납품처&타입']):
                store_list.append({
                    'key': str(r['납품처&타입']).replace(' ', '').strip(),
                    'code': str(r['배송코드']).strip(),
                    'center_num': str(r['납품처&타입'])[:4] # 앞 4자리 숫자(예: 0906) 추출
                })
        return prod_map, store_list, None
    except Exception as e:
        return {}, [], str(e)

st.title("🛒 홈플러스 수주 자동화 (함안 매칭 보정판)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_list, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['발주수량'], errors='coerce') > 0].copy()

            final_rows = []
            for _, row in df_raw.iterrows():
                p_code = str(row.get('상품코드', '')).strip()
                p_info = prod_dict.get(p_code, {'me': '', 'nm': row.get('상품명', '')})

                # 배송지 정보
                deliv_place = str(row.get('납품처', '')).strip()
                in_type = str(row.get('입고타입', '')).strip().replace('HYPER_FLOW', 'FLOW')
                lookup_key = f"{deliv_place}{in_type}".replace(' ', '')
                center_num = deliv_place[:4] # ordview의 납품처 앞 4자리(예: 0906)

                # [개선된 매칭 로직]
                shipping_code = ""
                # 1단계: 전체 글자로 매칭 시도
                for item in store_list:
                    if item['key'] == lookup_key:
                        shipping_code = item['code']
                        break
                
                # 2단계: 실패 시 센터 번호(0906 등)와 입고타입(FLOW 등)으로 매칭 시도
                if not shipping_code:
                    for item in store_list:
                        if item['center_num'] == center_num and in_type in item['key']:
                            shipping_code = item['code']
                            break

                final_rows.append({
                    '출고구분': 0,
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                    '발주처코드': '81020000',
                    '발주처': '홈플러스',
                    '배송코드': shipping_code,
                    '배송지': deliv_place,
                    '상품코드': p_info['me'],
                    '상품명': p_info['nm'],
                    'UNIT수량': int(float(row.get('발주수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                    '금        액': int(float(row.get('발주금액', 0))),
                    '부  가   세': int(float(row.get('발주금액', 0)) * 0.1),
                    'Type': '마트'
                })

            df_final = pd.DataFrame(final_rows)
            st.success("변환 완료!")
            st.dataframe(df_final)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='서식업로드')
            st.download_button(label="📥 엑셀 다운로드", data=output.getvalue(), file_name="Homeplus_Order.xlsx")
            
        except Exception as e:
            st.error(f"오류: {e}")
