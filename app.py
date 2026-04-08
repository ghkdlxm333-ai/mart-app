import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="홈플러스 수주 자동화", page_icon="🔵", layout="wide")

@st.cache_data
def load_master_data(path):
    try:
        df_prod = pd.read_excel(path, sheet_name='상품코드', dtype=str)
        prod_map = {
            str(r['상품코드']).strip().split('.')[0]: {
                'me': str(r['ME코드']).strip(), 
                'nm': str(r['상품명']).strip()
            } for _, r in df_prod.iterrows() if pd.notna(r['상품코드'])
        }
        
        df_store = pd.read_excel(path, sheet_name='Tesco 발주처코드', dtype=str)
        store_map = {}
        fallback_map = {}
        for _, r in df_store.iterrows():
            raw_key = str(r['납품처&타입']).strip()
            clean_key = raw_key.replace(" ", "")
            val = str(r['배송코드']).strip()
            if clean_key and val:
                store_map[clean_key] = val
                name_only = raw_key.replace("FLOW","").replace("SORTATION","").replace("STOCK","").replace(" ","").strip()
                if name_only:
                    fallback_map[name_only] = val
        return prod_map, store_map, fallback_map, None
    except Exception as e:
        return {}, {}, {}, str(e)

# --- [스타일 함수: SINGLE인 경우 노란색 표시] ---
def highlight_single(row):
    # '입고타입' 열이 'SINGLE'인 경우 해당 행의 배경색을 노란색으로 설정
    if row['입고타입'] == 'SINGLE':
        return ['background-color: yellow'] * len(row)
    return [''] * len(row)

st.title("🛒🔵 홈플러스 수주 자동화")

MASTER_FILE = "Tesco_서식파일_업데이트용.xlsx"
prod_dict, store_map, fallback_map, error = load_master_data(MASTER_FILE)

if error:
    st.error(f"마스터 파일 로드 실패: {error}")
else:
    # --- [추가된 안내 문구 섹션] ---
    st.markdown("### ※ 업로드 전 확인사항")
    st.info("💡 **엑셀파일 확장자를 .xlsx로 변환 후 업로드해주세요.** (xls, csv 파일은 변환이 필요합니다)")
    # -----------------------------------

    uploaded_file = st.file_uploader("본문(헤더분할) 엑셀파일로 첨부해주세요.", type=['xlsx'])
    
    if uploaded_file:
        if not uploaded_file.name.lower().endswith('.xlsx'):
            st.error("⚠️ .xlsx 형식의 파일만 업로드 가능합니다.")
            st.stop()

        try:
            df_raw = pd.read_excel(uploaded_file, header=1)
            df_raw = df_raw[pd.to_numeric(df_raw['낱개수량'], errors='coerce') > 0].copy()

            temp_rows = []
            for _, row in df_raw.iterrows():
                raw_place = str(row.get('납품처', '')).strip()
                
                # Q열(인덱스 16)에서 입고타입 추출
                try:
                    raw_type = str(row.iloc[16]).strip()
                except:
                    raw_type = ""
                
                c_type = raw_type.replace('HYPER_', '')
                m_key = (raw_place + c_type).replace(" ", "")
                
                s_code = store_map.get(m_key, "")
                if not s_code:
                    s_code = fallback_map.get(raw_place.replace(" ", ""), "")

                p_code = str(row.get('상품코드', '')).strip().split('.')[0]
                p_info = prod_dict.get(p_code)
                f_p_code = p_info['me'] if p_info else p_code
                f_p_name = p_info['nm'] if p_info else row.get('상품명', '')

                temp_rows.append({
                    '출고구분': 0, '수주일자': datetime.now().strftime('%Y%m%d'),
                    '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
                    '발주처코드': '81020000', '발주처': '홈플러스', '배송코드': s_code,
                    '배송지': raw_place, '상품코드': f_p_code, '상품명': f_p_name,
                    'UNIT수량': int(float(row.get('낱개수량', 0))),
                    'UNIT단가': int(float(row.get('낱개당 단가', 0))), 
                    'Type': '마트',
                    '입고타입': raw_type  # 강조 표시를 위해 임시 열 추가
                })

            if temp_rows:
                df_temp = pd.DataFrame(temp_rows)
                # 합산 시 '입고타입'도 포함하여 그룹화 (강조 표시 유지용)
                grp = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT단가', 'Type', '입고타입']
                df_final = df_temp.groupby(grp, as_index=False)['UNIT수량'].sum()
                
                df_final['금        액'] = df_final['UNIT수량'] * df_final['UNIT단가']
                df_final['부  가   세'] = (df_final['금        액'] * 0.1).astype(int)
                
                # 순서 정렬 (입고타입은 화면 확인용으로 포함)
                cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'Type', '입고타입']
                df_final = df_final[cols]

                st.success("✅ 분석 완료! 'SINGLE' 타입은 노란색으로 표시됩니다.")

                # --- [화면 출력 시 스타일 적용] ---
                styled_df = df_final.style.apply(highlight_single, axis=1)
                st.dataframe(styled_df, use_container_width=True)

                # 다운로드용 파일에서는 '입고타입' 열 제외 (양식 준수)
                df_download = df_final.drop(columns=['입고타입'])
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_download.to_excel(writer, index=False, sheet_name='서식업로드')
                
                st.download_button(label="📥 결과 다운로드", data=output.getvalue(), file_name=f"Homeplus_{datetime.now().strftime('%m%d')}.xlsx")
            else:
                st.info("데이터가 없습니다.")
        except Exception as e:
            st.error(f"오류: {e}")
