import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (ME코드 강제 매칭)", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        # 상품코드 시트 로드 (모든 코드는 문자열로 처리)
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        # 딕셔너리 생성 (키: 바코드, 값: ME코드)
        # .str.split('.').str[0] 를 통해 혹시 모를 소수점(.0) 제거
        prod_map = {}
        for _, r in df_prod.iterrows():
            barcode = str(r['상품코드']).strip().split('.')[0]
            me_code = str(r['ME코드']).strip()
            name = str(r['상품명']).strip()
            if barcode:
                prod_map[barcode] = {'me': me_code, 'nm': name}
        
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

st.title("🛒 홈플러스 수주 자동화 (ME코드 치환 버전)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_list, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 1. ordview 데이터 로드 및 전처리
            df_raw = pd.read_excel(uploaded_file, header=1)
            # 낱개수량이 0인 행 제거
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # 배송코드 매칭 로직 (기존과 동일)
                raw_place = str(row.get('납품처', '')).strip()
                place_num = raw_place[:4]
                in_type = str(row.get('입고타입', '')).strip().replace('HYPER_', '')
                shipping_code = ""
                for item in store_list:
                    if item['num'] == place_num and in_type in item['name']:
                        shipping_code = item['code']
                        break
                if not shipping_code:
                    for item in store_list:
                        if item['num'] == place_num:
                            shipping_code = item['code']
                            break

                # 2. [중요] 상품코드(바코드) -> ME코드 치환
                # ordview의 상품코드에서 소수점(.0) 제거 후 매칭
                raw_p_code = str(row.get('상품코드', '')).strip().split('.')[0]
                p_info = prod_dict.get(raw_p_code)

                if p_info:
                    final_p_code = p_info['me']  # ME코드로 변경
                    final_p_name = p_info['nm']  # 마스터의 상품명으로 변경
                else:
                    # 마스터에 없을 경우 식별을 위해 유지 (혹은 경고 표시)
                    final_p_code = raw_p_code
                    final_p_name = row.get('상품명', '')

                temp_rows.append({
                    '출고구분': 0,
                    '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                    '발주처코드': '81020000',
                    '발주처': '홈플러스',
                    '배송코드': shipping_code,
                    '배송지': raw_place,
                    '상품코드': final_p_code,  # 여기서 ME코드가 들어감
                    '상품명': final_p_name,
                    '낱개수량': int(float(row.get('낱개수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                    'Type': '마트'
                })

            if temp_rows:
                df_temp = pd.DataFrame(temp_rows)
                
                # 3. 합산 처리
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False)['낱개수량'].sum()
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                
                # 금액 재계산
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                # 양식 순서 정리
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]

                st.success("✅ ME코드 변환 및 합산이 완료되었습니다.")
                st.dataframe(df_final)

                # 4. 엑셀 출력 (상품코드가 숫자로 변하지 않도록 처리)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                    # 엑셀 시트에서 '상품코드' 열을 텍스트 형식으로 강제 지정 가능하지만, 
                    # 기본적으로 pandas가 문자열은 텍스트로 저장합니다.
                
                st.download_button(
                    label="📥 ME코드 결과 파일 다운로드",
                    data=output.getvalue(),
                    file_name=f"Homeplus_ME_Final_{datetime.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("변환할 데이터가 없습니다.")
        except Exception as e:
            st.error(f"처리 중 오류 발생: {e}")
