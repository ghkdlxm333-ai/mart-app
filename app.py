import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (SINGLE 추가 버전)", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        # 상품코드 로드
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {str(r['상품코드']).strip(): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        # 배송코드 로드
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_list = []
        for _, r in df_store.iterrows():
            name_val = str(r['납품처&타입']).strip() if pd.notna(r['납품처&타입']) else ""
            code_val = str(r['배송코드']).strip() if pd.notna(r['배송코드']) else ""
            if name_val and code_val:
                store_list.append({'name': name_val, 'num': name_val[:4], 'code': code_val})
        return prod_map, store_list, None
    except Exception as e:
        return {}, [], str(e)

st.title("🛒 홈플러스 수주 자동화 (FLOW/SORT/SINGLE 구분)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_list, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 헤더 위치는 ordview 양식에 따라 header=1 또는 0으로 조정
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                raw_place = str(row.get('납품처', '')).strip()
                place_num = raw_place[:4] # '0906' 등 추출
                in_type = str(row.get('입고타입', '')).strip().upper()

                # --- 타입 판별 보강 (SINGLE 추가) ---
                if "SINGLE" in in_type:
                    target_type = "SINGLE"
                elif "SORT" in in_type:
                    target_type = "SORTATION"
                elif "FLOW" in in_type:
                    target_type = "FLOW"
                else:
                    target_type = in_type
                
                shipping_code = ""
                # 1순위: 센터번호 + 타입(SINGLE/FLOW/SORT) 일치 확인
                for item in store_list:
                    if item['num'] == place_num and target_type in item['name'].upper():
                        shipping_code = item['code']
                        break
                
                # 2순위: 못 찾은 경우 센터번호로만 확인
                if not shipping_code:
                    for item in store_list:
                        if item['num'] == place_num:
                            shipping_code = item['code']
                            break

                p_code = str(row.get('상품코드', '')).strip()
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
                
                # 합산 기준: 배송코드와 상품코드가 모두 같을 때
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False)['낱개수량'].sum()
                
                # 결과 정리
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success(f"변환 완료! (FLOW/SORT/SINGLE 타입별 합산 적용)")
                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                
                st.download_button(
                    label="📥 SINGLE 구분된 결과 다운로드",
                    data=output.getvalue(),
                    file_name=f"HP_Final_TypeMixed_{datetime.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"오류 발생: {e}")
