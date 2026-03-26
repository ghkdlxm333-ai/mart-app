import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="마트 수주 자동화 시스템", layout="wide")
st.title("🛒 마트(Tesco) 수주 통합 업로드 시스템")
st.markdown("로우데이터(ordview)를 업로드하면 마스터 파일의 기준에 따라 통합 수주 양식으로 변환합니다.")

# --- 2. 설정 변수 ---
# 최종 출력 양식 컬럼 (사용자 요청 기반)
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금액', '부가세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

# --- 3. 데이터 처리 로직 ---
def process_mart_order(raw_file, master_file):
    # 1. 마스터 파일의 각 시트 로드 (기준 정보)
    # 실제 배포 시에는 github url이나 로컬 경로에서 불러오도록 설정 가능
    xls_master = pd.ExcelFile(master_file)
    df_product_master = pd.read_excel(xls_master, sheet_name='상품코드', dtype=str)
    df_store_master = pd.read_excel(xls_master, sheet_name='Tesco 발주처코드', dtype=str)
    
    # 2. 로우 데이터(ordview) 로드
    df_raw = pd.read_excel(raw_file, header=1) # ordview는 보통 2행부터 데이터 시작
    
    # 3. 데이터 변환 및 매핑 (Summary 시트 로직 구현)
    # ordview의 정보를 바탕으로 필요한 값 추출
    processed_data = []
    
    for _, row in df_raw.iterrows():
        # 수량이 0인 경우 제외
        qty = row.get('발주수량', 0)
        if qty == 0 or pd.isna(qty):
            continue
            
        # 데이터 매핑 (마스터 파일의 Summary 시트 로직 모사)
        order_info = {
            '출고구분': 0,
            '수주일자': datetime.now().strftime('%Y%m%d'),
            '납품일자': str(row.get('납품일자', '')).replace('-', '')[:8],
            '발주처코드': row.get('납품처코드', ''),
            '발주처': row.get('납품처', ''),
            '배송코드': row.get('배송처코드', ''),
            '배송지': row.get('배송처', ''),
            '상품코드': row.get('상품코드', ''), # 바코드 기반 매핑 필요시 추가 로직 구현
            '상품명': row.get('상품명', ''),
            'UNIT수량': int(qty),
            'UNIT단가': int(row.get('낱개당 단가', 0)),
            '금액': int(row.get('발주금액', 0)),
            '부가세': int(int(row.get('발주금액', 0)) * 0.1),
            'LOT': '',
            '특이사항1': row.get('입고타입', ''),
            'Type': '마트',
            '특이사항2': ''
        }
        processed_data.append(order_info)
    
    return pd.DataFrame(processed_data)

# --- 4. 파일 업로드 영역 ---
st.sidebar.header("📁 마스터 파일 설정")
# 실제 운영시 GitHub의 RAW 링크를 사용하거나 서버 파일 경로 지정
master_upload = st.sidebar.file_uploader("마스터 파일(Tesco_서식파일) 업로드", type=['xlsx'])

st.subheader("📥 로우 데이터 업로드")
raw_upload = st.file_uploader("ordview 파일을 업로드하세요 (Excel)", type=['xlsx', 'xls'])

# --- 5. 실행 및 결과 출력 ---
if raw_upload and master_upload:
    try:
        with st.spinner("데이터 매핑 및 변환 중..."):
            df_result = process_mart_order(raw_upload, master_upload)
            
            # 컬럼 순서 맞추기 및 누락 컬럼 생성
            for col in FINAL_COLUMNS:
                if col not in df_result.columns:
                    df_result[col] = ''
            
            df_final = df_result[FINAL_COLUMNS]
            
        st.success(f"✅ 변환 완료! (총 {len(df_final)}건)")
        
        # 미리보기
        st.dataframe(df_final, use_container_width=True)
        
        # 엑셀 다운로드 버튼
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='통합수주업로드')
        
        st.download_button(
            label="📥 통합 수주 양식 다운로드",
            data=output.getvalue(),
            file_name=f"Mart_Upload_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ 처리 중 오류 발생: {e}")
        st.info("파일 형식이 마스터 파일의 '원본 데이터' 및 'Summary' 시트 구조와 일치하는지 확인해주세요.")
else:
    st.info("왼쪽 사이드바에 마스터 파일을 넣고, 중앙에 ordview 파일을 업로드해 주세요.")
