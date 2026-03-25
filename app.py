import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from datetime import datetime, timedelta, timezone

# --- 설정 ---
st.set_page_config(page_title="Tesco 수주 자동 변환기", layout="wide")

# 고정된 서식 파일 경로 (GitHub에 업로드한 파일명과 일치해야 함)
TEMPLATE_PATH = "Tesco_Template.xlsx" 

# 사내 표준 양식 컬럼
FINAL_COLUMNS = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금액', '부가세', 'LOT', '특이사항1', 'Type', '특이사항2']

def clean_val(val):
    return str(val).strip() if pd.notna(val) else ""

@st.cache_data
def process_ordview(file):
    # 1. ordview 파일 읽기 (데이터는 2행부터 시작하는 경우가 많음)
    # 첫 행이 '주문서LIST'인 경우 header=1로 설정하여 실제 컬럼명을 읽음
    df_raw = pd.read_excel(file, header=1)
    
    # 발주번호가 있는 행만 유효 데이터로 간주
    df_raw = df_raw.dropna(subset=['발주번호'])
    
    # 2. 통합 양식 매핑 로직
    kst = timezone(timedelta(hours=9))
    today_str = datetime.now(kst).strftime("%Y%m%d")
    
    processed_df = pd.DataFrame()
    processed_df['출고구분'] = 0
    processed_df['수주일자'] = today_str
    # 날짜 형식 정리 (2026-03-25 -> 20260325)
    processed_df['납품일자'] = df_raw['납품일자'].astype(str).str.replace(r'\D', '', regex=True).str[:8]
    processed_df['발주처코드'] = df_raw['납품처코드'].apply(clean_val)
    processed_df['발주처'] = df_raw['납품처'].apply(clean_val)
    processed_df['배송코드'] = df_raw['배송처코드'].apply(clean_val)
    processed_df['배송지'] = df_raw['배송처'].apply(clean_val)
    processed_df['상품코드'] = df_raw['상품코드'].apply(clean_val)
    processed_df['상품명'] = df_raw['상품명'].apply(clean_val)
    processed_df['UNIT수량'] = pd.to_numeric(df_raw['발주수량'], errors='coerce').fillna(0).astype(int)
    processed_df['UNIT단가'] = pd.to_numeric(df_raw['낱개당 단가'], errors='coerce').fillna(0).astype(int)
    processed_df['금액'] = pd.to_numeric(df_raw['발주금액'], errors='coerce').fillna(0).astype(int)
    processed_df['부가세'] = (processed_df['금액'] * 0.1).astype(int)
    
    for col in FINAL_COLUMNS:
        if col not in processed_df.columns:
            processed_df[col] = ""
            
    return processed_df[FINAL_COLUMNS], df_raw

# --- 메인 UI ---
st.title("🏪 Tesco(홈플러스) 수주 통합 시스템")

# 서식 파일 존재 확인
if not os.path.exists(TEMPLATE_PATH):
    st.error(f"❌ 서식 파일({TEMPLATE_PATH})을 찾을 수 없습니다. GitHub에 파일을 업로드했는지 확인해주세요.")
else:
    st.success("✅ Tesco 기준 서식 로드 완료")
    
    uploaded_file = st.file_uploader("홈플러스 ordview 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            with st.spinner("데이터 변환 및 서식 적용 중..."):
                final_df, raw_df = process_ordview(uploaded_file)
                
                # --- 엑셀 시트 교체 (오류 방지 로직) ---
                # 1. 기존 워크북 로드
                wb = openpyxl.load_workbook(TEMPLATE_PATH)
                
                # 2. '원본 데이터' 시트 업데이트
                # 만약 시트가 있으면 삭제하고 새로 생성 (가장 안전한 방법)
                target_sheet_name = '원본 데이터'
                if target_sheet_name in wb.sheetnames:
                    # 시트 삭제 전, 다른 시트가 반드시 하나 이상 존재해야 함 (openpyxl 규칙)
                    del wb[target_sheet_name]
                
                new_ws = wb.create_sheet(target_sheet_name)
                
                # 3. 데이터 쓰기 (Pandas df를 openpyxl 시트로 전달)
                # 헤더 쓰기
                for j, col_name in enumerate(raw_df.columns, 1):
                    new_ws.cell(row=1, column=j, value=col_name)
                # 데이터 쓰기
                for i, row in enumerate(raw_df.values, 2):
                    for j, value in enumerate(row, 1):
                        new_ws.cell(row=i, column=j, value=value)
                
                # 4. 파일 저장 메모리 버퍼
                output = io.BytesIO()
                wb.save(output)
                processed_data = output.getvalue()

                st.balloons()
                st.subheader("📊 변환 결과 미리보기")
                st.dataframe(final_df, use_container_width=True)

                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 서식 적용 파일 다운로드",
                        data=processed_data,
                        file_name=f"Tesco_Order_{datetime.now().strftime('%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                with col2:
                    # 사내 ERP 업로드용 (단순 엑셀)
                    final_output = io.BytesIO()
                    final_df.to_excel(final_output, index=False)
                    st.download_button(
                        label="📥 통합 수주양식만 다운로드",
                        data=final_output.getvalue(),
                        file_name=f"통합수주_{datetime.now().strftime('%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"⚠️ 오류 발생: {e}")
            st.info("Tip: 업로드한 ordview 파일의 첫 번째 행이 '주문서LIST'로 시작하는지 확인해주세요.")

st.divider()
st.caption("Developed by Jay | SCM Outbound Operations Support")
