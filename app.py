import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 1. 설정 ---
st.set_page_config(page_title="Tesco 수주 관리 시스템", layout="wide")
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

MASTER_FILE_NAME = "Tesco_서식파일_업데이트용.xlsx"

@st.cache_data
def load_master_data():
    if not os.path.exists(MASTER_FILE_NAME):
        return None, None, None, "마스터 파일(Tesco_서식파일_업데이트용.xlsx)을 찾을 수 없습니다."
    
    try:
        # 상품코드: 바코드를 문자열 숫자로 정제하여 로드
        df_item = pd.read_excel(MASTER_FILE_NAME, sheet_name='상품코드', dtype=str)
        df_item['바코드'] = df_item['바코드'].str.strip()
        item_map = df_item.dropna(subset=['바코드']).drop_duplicates('바코드').set_index('바코드')[['ME코드', '상품명']].to_dict('index')
        
        # 발주처/배송코드: Summary 시트의 기준이 되는 매핑 데이터
        df_store = pd.read_excel(MASTER_FILE_NAME, sheet_name='Tesco 발주처코드', dtype=str)
        store_map = df_store.dropna(subset=['납품처&타입']).drop_duplicates('납품처&타입').set_index('납품처&타입')[['발주처코드', '배송코드', '배송지']].to_dict('index')
        
        # Summary 시트 양식 (헤더용)
        df_summary = pd.read_excel(MASTER_FILE_NAME, sheet_name='Summary', dtype=str).head(0)
        
        return item_map, store_map, df_summary, None
    except Exception as e:
        return None, None, None, f"마스터 로드 오류: {e}"

# --- 2. 데이터 처리 ---
def process_data(uploaded_file, item_map, store_map):
    # 1. Raw 데이터 읽기 및 FLOW 변환
    raw_df = pd.read_excel(uploaded_file, header=1, dtype=str)
    raw_df.columns = raw_df.columns.str.strip()
    raw_df = raw_df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # 2. [중요] 상품코드를 숫자로 변환 (지수 표기 제거 및 정수화)
    def clean_code(x):
        try:
            return str(int(float(x)))
        except:
            return str(x).strip()
    
    raw_df['상품코드'] = raw_df['상품코드'].apply(clean_code)
    
    # 3. 수량/금액 숫자화 및 합산
    raw_df['낱개수량'] = pd.to_numeric(raw_df['낱개수량'], errors='coerce').fillna(0)
    raw_df['낱개당 단가'] = pd.to_numeric(raw_df['낱개당 단가'], errors='coerce').fillna(0)
    raw_df['발주금액'] = pd.to_numeric(raw_df['발주금액'], errors='coerce').fillna(0)
    
    # 4. 발주처 & 상품코드별 합계 계산 (Summary 로직)
    summary_grouped = raw_df.groupby([
        '납품처코드', '납품처', '배송처코드', '배송처', '상품코드', '입고타입', '납품일자'
    ], as_index=False).agg({
        '낱개수량': 'sum',
        '발주금액': 'sum',
        '낱개당 단가': 'mean' # 단가는 동일하다고 가정하거나 평균값 사용
    })
    
    # 5. 최종 결과물 구성 (마스터 Summary 시트 값 우선)
    final_rows = []
    for _, row in summary_grouped.iterrows():
        if row['낱개수량'] <= 0: continue
        
        m_key = f"{row['납품처코드']} {row['납품처']}{row['입고타입']}"
        s_info = store_map.get(m_key, {})
        i_info = item_map.get(row['상품코드'], {})
        
        final_rows.append({
            '출고구분': '0',
            '수주일자': today_str,
            '납품일자': str(row['납품일자']).replace('-', '')[:8],
            '발주처코드': s_info.get('발주처코드', row['납품처코드']),
            '발주처': row['납품처'],
            '배송코드': s_info.get('배송코드', row['배송처코드']),
            '배송지': s_info.get('배송지', row['배송처']),
            '상품코드': i_info.get('ME코드', row['상품코드']),
            '상품명': i_info.get('상품명', row['상품명']),
            'UNIT수량': int(row['낱개수량']),
            'UNIT단가': int(row['낱개당 단가']),
            '금        액': int(row['발주금액']),
            '부  가   세': int(row['발주금액'] * 0.1),
            'LOT': '',
            '특이사항_1': '',
            'Type': row['입고타입'],
            '특이사항_2': ''
        })
        
    final_df = pd.DataFrame(final_rows)
    headers = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항']
    
    return raw_df, final_df, headers

# --- 3. 실행 UI ---
st.title("🛒 Tesco 수주 데이터 통합 시스템")

item_map, store_map, _, err = load_master_data()

if err:
    st.error(err)
else:
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx'])
    
    if uploaded_file:
        raw_data, final_data, headers = process_data(uploaded_file, item_map, store_map)
        
        st.subheader("✅ 수량 합산 결과 (Summary 기준)")
        st.write(f"총 {len(final_data)}개의 합산 항목이 생성되었습니다.")
        st.dataframe(final_data)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 원본 데이터를 '원본 데이터' 시트로 저장
            raw_data.to_excel(writer, index=False, sheet_name='원본 데이터')
            # 결과 데이터를 '서식' 시트로 저장 (중복 컬럼명 처리)
            final_data.columns = headers
            final_data.to_excel(writer, index=False, sheet_name='서식')
            
        st.download_button("📥 통합 수주 파일 다운로드", output.getvalue(), f"Tesco_Result_{today_str}.xlsx")
