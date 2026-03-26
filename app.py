import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta, timezone

# --- 1. 페이지 및 설정 ---
st.set_page_config(page_title="Tesco 마트 수주 시스템", layout="wide")
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

st.title("🛒 Tesco(홈플러스) 수주 자동화 시스템")

# --- 2. 상수 및 컬럼 정의 ---
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금액', '부가세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

# --- 3. 데이터 프로세싱 함수 ---
def process_mart_data(ordview_file, master_file):
    # A. 마스터 데이터 로드 (상품코드 및 발주처 매핑)
    # 상품코드 시트에서 바코드별 ME코드 매핑
    df_item = pd.read_excel(master_file, sheet_name='상품코드', dtype=str)
    item_map = df_item.set_index('바코드')['ME코드'].to_dict()
    
    # 발주처코드 시트에서 납품처/타입별 배송코드 매핑
    df_store = pd.read_excel(master_file, sheet_name='Tesco 발주처코드', dtype=str)
    # '납품처&타입' 컬럼을 키로 사용하여 배송코드와 발주처코드를 가져옴
    store_map = df_store.set_index('납품처&타입')[['발주처코드', '배송코드', '배송지']].to_dict('index')

    # B. ordview 파일 읽기 (헤더 위치 수정: header=1)
    # 첫 줄이 '주문서LIST'이므로 두 번째 줄을 컬럼명으로 인식하게 함
    df = pd.read_excel(ordview_file, header=1, dtype=str)
    df.columns = df.columns.str.strip()
    
    # C. 'HYPER_FLOW' -> 'FLOW' 변환
    df = df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # D. 데이터 전처리 (숫자 변환)
    df['낱개수량'] = pd.to_numeric(df['낱개수량'], errors='coerce').fillna(0)
    df['낱개당 단가'] = pd.to_numeric(df['낱개당 단가'], errors='coerce').fillna(0)
    df['발주금액'] = pd.to_numeric(df['발주금액'], errors='coerce').fillna(0)
    
    # E. Summary 생성 (발주처/배송처/상품별 묶기)
    summary = df.groupby([
        '납품처코드', '납품처', '배송처코드', '배송처', '상품코드', '상품명', '납품일자', '낱개당 단가', '입고타입'
    ], as_index=False).agg({'낱개수량': 'sum', '발주금액': 'sum'})
    
    # 수량 0 제외
    summary = summary[summary['낱개수량'] > 0]
    
    # F. 최종 통합 양식 구성
    final_df = pd.DataFrame(columns=FINAL_COLUMNS)
    
    # 매핑 키 생성 (예: 0982 안성ADC물류센터FLOW)
    summary['MappingKey'] = summary['납품처코드'] + " " + summary['납품처'] + summary['입고타입']
    
    # 데이터 매핑 및 채우기
    res_list = []
    for _, row in summary.iterrows():
        m_key = row['MappingKey']
        store_info = store_map.get(m_key, {})
        
        res_list.append({
            '출고구분': '0',
            '수주일자': today_str,
            '납품일자': str(row['납품일자']).replace('-', '')[:8],
            '발주처코드': store_info.get('발주처코드', row['납품처코드']),
            '발주처': row['납품처'],
            '배송코드': store_info.get('배송코드', row['배송처코드']),
            '배송지': store_info.get('배송지', row['배송처']),
            '상품코드': item_map.get(row['상품코드'], row['상품코드']), # 마스터에 없으면 원본 코드 유지
            '상품명': row['상품명'],
            'UNIT수량': int(row['낱개수량']),
            'UNIT단가': int(row['낱개당 단가']),
            '금액': int(row['발주금액']),
            '부가세': int(row['발주금액'] * 0.1),
            'Type': row['입고타입']
        })
    
    final_result = pd.DataFrame(res_list)
    return df, final_result

# --- 4. UI 구성 ---
st.info("💡 마스터 파일(Tesco_서식파일_업데이트용.xlsx)과 ordview 파일을 모두 업로드하세요.")

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("1. 마스터 서식 파일 업로드", type=['xlsx'])
with col2:
    ordview_file = st.file_uploader("2. ordview(Raw) 파일 업로드", type=['xlsx', 'xls'])

if master_file and ordview_file:
    try:
        raw_transformed, final_df = process_mart_data(ordview_file, master_file)
        
        st.divider()
        st.success(f"변환 성공! 총 {len(final_df)}건의 유효 수주가 생성되었습니다.")
        
        # 결과 표시
        st.subheader("📊 통합 수주업로드 결과 미리보기")
        st.dataframe(final_df, use_container_width=True)
        
        # 다운로드 버튼
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='서식')
            raw_transformed.to_excel(writer, index=False, sheet_name='원본데이터_FLOW변환')
            
        st.download_button(
            label="📥 통합 수주 파일 다운로드 (Excel)",
            data=output.getvalue(),
            file_name=f"Tesco_통합수주_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"처리 중 오류 발생: {e}")
        st.warning("팁: ordview 파일의 '낱개수량' 컬럼이 정확히 두 번째 줄에 있는지 확인하세요.")
