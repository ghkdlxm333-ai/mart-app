import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화 (최종 보정)", layout="wide")

def clean_text(text):
    """모든 공백 제거 및 대문자 변환"""
    if pd.isna(text): return ""
    return re.sub(r'\s+', '', str(text)).upper()

@st.cache_data
def load_master_data(path):
    try:
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {clean_text(r['상품코드']): {'me': str(r['ME코드']).strip(), 'nm': str(r['상품명']).strip()} 
                    for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])}
        
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        vlookup_map = {}
        for _, r in df_store.iterrows():
            # D열(3번 인덱스)과 E열(4번 인덱스) 직접 참조
            d_val = str(r.iloc[3]).strip()
            e_val = str(r.iloc[4]).strip()
            if d_val and d_val.lower() != 'nan':
                vlookup_map[clean_text(d_val)] = e_val
        return prod_map, vlookup_map, None
    except Exception as e:
        return {}, {}, str(e)

st.title("🛒 홈플러스 수주 자동화 (매칭 오류 완전 해결)")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, vlookup_map, error = load_master_data(MASTER_FILE)

if not error:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        try:
            # 1. ordview 읽기 (header=1)
            df_raw = pd.read_excel(uploaded_file, header=1)
            
            # 컬럼명 공백 제거 (매우 중요: '입고타입 ' 처럼 공백이 붙어있는 경우 대비)
            df_raw.columns = [c.strip() if isinstance(c, str) else c for c in df_raw.columns]
            
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                # 2. I열(납품처)과 Q열(입고타입) 값 추출
                # 컬럼명이 정확하지 않을 경우를 대비해 위치 기반 iloc도 고려 가능하나 우선 이름으로 시도
                i_val = str(row.get('납품처', '')).strip()
                q_val = str(row.get('입고타입', '')).strip().upper()
                
                # 3. 지시사항: HYPER_FLOW -> FLOW 변환 (매칭 키 생성 전 필수)
                if "HYPER_FLOW" in q_val:
                    q_converted = "FLOW"
                elif "SORT" in q_val:
                    q_converted = "SORTER"
                elif "SINGLE" in q_val:
                    q_converted = "SINGLE"
                else:
                    q_converted = q_val
                
                # 4. 결합 키 생성 (VLOOKUP I&Q 방식)
                # 예: 0906NEW함안상온물류센터 + FLOW
                lookup_key = clean_text(i_val + q_converted)
                
                # 5. 매칭 시도
                shipping_code = vlookup_map.get(lookup_key, "")

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
                    '배송지': i_val,
                    '상품코드': p_info['me'],
                    '상품명': p_info['nm'],
                    '낱개수량': int(float(row.get('낱개수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))),
                    'Type': '마트'
                })

            if temp_rows:
                df_temp = pd.DataFrame(temp_rows)
                group_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type']
                df_final = df_temp.groupby(group_cols, as_index=False).agg({'낱개수량': 'sum'})
                
                df_final.rename(columns={'낱개수량': 'UNIT수량'}, inplace=True)
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                df_final = df_final[['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type']]
                
                st.success("매칭 완료")
                
                # 오류 디버깅 영역
                missing = df_final[df_final['배송코드'] == ""]
                if not missing.empty:
                    st.error(f"❌ 매칭 실패 발생!")
                    # 사용자가 직접 확인할 수 있도록 매칭 시도했던 키값을 보여줌
                    st.write("실패한 납품처:", missing['배송지'].unique())
                    st.info("💡 팁: '입고타입' 컬럼이 정확히 'HYPER_FLOW'나 'SORTATION'으로 적혀있는지 확인해주세요.")

                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='서식업로드')
                st.download_button(label="📥 결과 다운로드", data=output.getvalue(), file_name=f"HP_Final_{datetime.now().strftime('%m%d')}.xlsx")
        except Exception as e:
            st.error(f"데이터 처리 오류: {e}")
