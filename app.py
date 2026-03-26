import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 1. 설정 및 KST 시간 ---
st.set_page_config(page_title="Tesco 수주 Summary 합산기", layout="wide")
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

# --- 2. 마스터 파일 경로 ---
MASTER_FILE_NAME = "Tesco_서식파일_업데이트용.xlsx"

@st.cache_data
def load_fixed_master():
    """GitHub의 마스터 파일을 로드하여 Summary 양식과 매핑 데이터를 가져옵니다."""
    if not os.path.exists(MASTER_FILE_NAME):
        return None, None, None, f"❌ '{MASTER_FILE_NAME}' 파일을 찾을 수 없습니다."
    
    try:
        # 1. 상품코드 매핑 (바코드 -> ME코드)
        df_item = pd.read_excel(MASTER_FILE_NAME, sheet_name='상품코드', dtype=str)
        df_item = df_item.dropna(subset=['바코드']).drop_duplicates(subset=['바코드'])
        item_map = df_item.set_index('바코드')['ME코드'].to_dict()
        
        # 2. 발주처코드 매핑 (납품처&타입 -> 발주처/배송코드)
        # 'Tesco 발주처코드' 시트의 데이터를 기준 정보로 사용
        df_store = pd.read_excel(MASTER_FILE_NAME, sheet_name='Tesco 발주처코드', dtype=str)
        df_store = df_store.dropna(subset=['납품처&타입']).drop_duplicates(subset=['납품처&타입'])
        store_map = df_store.set_index('납품처&타입')[['발주처코드', '배송코드', '배송지']].to_dict('index')
        
        # 3. Summary 시트의 컬럼 양식 가져오기
        df_summary_template = pd.read_excel(MASTER_FILE_NAME, sheet_name='Summary', dtype=str)
        
        return item_map, store_map, df_summary_template, None
    except Exception as e:
        return None, None, None, f"❌ 마스터 로드 오류: {e}"

# --- 3. 데이터 처리 로직 ---
def process_mart_data(ordview_file, item_map, store_map):
    # ordview 읽기 (헤더 2행)
    raw_df = pd.read_excel(ordview_file, header=1, dtype=str)
    raw_df.columns = raw_df.columns.str.strip()
    
    # 'HYPER_FLOW' -> 'FLOW' 변환 (전체 데이터 대상)
    raw_df = raw_df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # 숫자 변환 (합산을 위해 필수)
    raw_df['낱개수량'] = pd.to_numeric(raw_df['낱개수량'], errors='coerce').fillna(0)
    raw_df['낱개당 단가'] = pd.to_numeric(raw_df['낱개당 단가'], errors='coerce').fillna(0)
    raw_df['발주금액'] = pd.to_numeric(raw_df['발주금액'], errors='coerce').fillna(0)
    
    # [핵심] 발주처(납품처) + 상품코드 기준 합산
    # 같은 곳에 같은 물건이 들어온 경우 하나로 합칩니다.
    summary_grouped = raw_df.groupby([
        '납품처코드', '납품처', '배송처코드', '배송처', '상품코드', '상품명', '납품일자', '낱개당 단가', '입고타입'
    ], as_index=False).agg({
        '낱개수량': 'sum',
        '발주금액': 'sum'
    })
    
    # 수량 0인 항목 제외
    summary_grouped = summary_grouped[summary_grouped['낱개수량'] > 0]
    
    # 결과 리스트 생성
    res_list = []
    for _, row in summary_grouped.iterrows():
        # 마스터 매핑 키 생성 (예: 0982 안성ADC물류센터FLOW)
        m_key = f"{row['납품처코드']} {row['납품처']}{row['입고타입']}"
        store_info = store_map.get(m_key, {})
        
        # 사내 상품코드 변환
        internal_code = item_map.get(row['상품코드'], row['상품코드'])
        
        # 'Summary' 시트 기준 데이터 구성
        res_list.append({
            '출고구분': '0',
            '수주일자': today_str,
            '납품일자': str(row['납품일자']).replace('-', '')[:8],
            '발주처코드': store_info.get('발주처코드', row['납품처코드']),
            '발주처': row['납품처'],
            '배송코드': store_info.get('배송코드', row['배송처코드']),
            '배송지': store_info.get('배송지', row['배송처']),
            '상품코드': internal_code,
            '상품명': row['상품명'],
            'UNIT수량': int(row['낱개수량']),
            'UNIT단가': int(row['낱개당 단가']),
            '금        액': int(row['발주금액']),
            '부  가   세': int(row['발주금액'] * 0.1),
            'LOT': '',
            '특이사항_1': '', # 중복 방지용 임시 키
            'Type': row['입고타입'],
            '특이사항_2': ''  # 중복 방지용 임시 키
        })
    
    result_df = pd.DataFrame(res_list)
    
    # 최종 엑셀용 컬럼명 (중복 '특이사항' 허용을 위해 리스트 준비)
    final_output_headers = [
        '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
        '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
        'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
    ]
    
    return raw_df, result_df, final_output_headers

# --- 4. 메인 UI ---
st.title("🛒 Tesco 수주 통합 자동화 (Summary 기준)")

# 마스터 자동 로드
item_map, store_map, template, error_msg = load_fixed_master()

if error_msg:
    st.error(error_msg)
else:
    st.success("✅ 마스터(Tesco_서식파일_업데이트용.xlsx) 로드 완료")
    
    ordview_file = st.file_uploader("ordview(Raw) 파일을 업로드하세요", type=['xlsx', 'xls'])

    if ordview_file:
        try:
            with st.spinner("데이터 처리 중..."):
                raw_transformed, final_df, final_headers = process_mart_data(ordview_file, item_map, store_map)
            
            st.divider()
            st.subheader("📊 수주 합산 결과 (Summary 기준)")
            st.dataframe(final_df, use_container_width=True)
            
            # 엑셀 다운로드 파일 생성
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 1. 결과 데이터 (중복 컬럼명 처리)
                # 데이터프레임 컬럼을 강제로 리스트 헤더로 덮어씌워 저장
                final_df.columns = final_headers
                final_df.to_excel(writer, index=False, sheet_name='서식')
                
                # 2. 원본 데이터 시트 (FLOW 변환본)
                raw_transformed.to_excel(writer, index=False, sheet_name='원본 데이터')
            
            st.download_button(
                label="📥 통합 수주 결과 다운로드 (Excel)",
                data=output.getvalue(),
                file_name=f"Tesco_통합수주_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
