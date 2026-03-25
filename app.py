import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 설정 ---
st.set_page_config(page_title="Tesco Summary 변환 시스템", layout="wide")

# 사내 표준 양식 컬럼 (사용자 제공 양식 기준)
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

# 실제 엑셀 저장 시 사용할 컬럼명 (공백 유지)
REAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
]

def clean_numeric(val):
    """숫자 데이터 정제 (NaN 처리 및 정수화)"""
    try:
        return int(pd.to_numeric(val, errors='coerce'))
    except:
        return 0

# --- 메인 UI ---
st.title("🏪 Tesco 'Summary' 데이터 통합 변환기")
st.info("업로드한 파일의 'Summary' 시트에서 수량이 있는 데이터만 추출하여 통합 양식으로 변환합니다.")

# 파일 업로더 (rawdata 업로드)
uploaded_file = st.file_uploader("Tesco 서식파일(rawdata)을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        with st.spinner("Summary 시트 분석 중..."):
            # 1. 'Summary' 시트만 읽어오기
            # Summary 시트의 헤더 위치가 보통 1행이 아닐 수 있으므로 데이터 확인 후 조정 필요
            # 여기서는 헤더가 1행에 있다고 가정하고 읽은 뒤 수량 컬럼을 체크합니다.
            df_summary = pd.read_excel(uploaded_file, sheet_name='Summary', dtype=str)
            
            # 컬럼명 공백 제거
            df_summary.columns = df_summary.columns.str.strip()
            
            # 2. 수량 0인 자료 제외 로직
            # '수량' 컬럼 숫자로 변환
            df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
            df_filtered = df_summary[df_summary['수량'] > 0].copy()
            
            if df_filtered.empty:
                st.error("❌ 'Summary' 시트에 수량이 0보다 큰 데이터가 없습니다.")
            else:
                # 3. 통합 수주업로드 양식으로 매핑
                kst = timezone(timedelta(hours=9))
                today_str = datetime.now(kst).strftime("%Y%m%d")
                
                final_df = pd.DataFrame(columns=FINAL_COLUMNS)
                
                final_df['출고구분'] = 0
                final_df['수주일자'] = today_str
                # 납품일자는 보통 파일명이나 특정 셀에 있으나, Summary 시트에 없다면 오늘 날짜 등으로 대체 가능
                # 여기서는 빈값으로 두거나 필요시 데이터에서 추출하는 로직 추가 가능
                final_df['납품일자'] = today_str 
                
                final_df['발주처코드'] = df_filtered['발주코드'].fillna('')
                final_df['발주처'] = "홈플러스" # 혹은 데이터 내 '발주처' 컬럼 매핑
                final_df['배송코드'] = df_filtered['배송코드'].fillna('')
                final_df['배송지'] = "" # Summary에 배송지 명칭이 있다면 매핑
                final_df['상품코드'] = df_filtered['상품코드'].fillna('')
                final_df['상품명'] = df_filtered['상품명'].fillna('')
                final_df['UNIT수량'] = df_filtered['수량'].astype(int)
                final_df['UNIT단가'] = pd.to_numeric(df_filtered['UNIT단가'], errors='coerce').fillna(0).astype(int)
                
                # 금액 및 부가세 계산
                final_df['금        액'] = final_df['UNIT수량'] * final_df['UNIT단가']
                final_df['부  가   세'] = (final_df['금        액'] * 0.1).astype(int)
                
                # 나머지 빈 컬럼 채우기
                final_df.fillna('', inplace=True)
                
                # 결과 출력
                st.success(f"✅ 총 {len(final_df)}건의 유효 수주 데이터를 추출했습니다.")
                st.dataframe(final_df, use_container_width=True)
                
                # 4. 다운로드 파일 생성 (Excel)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # 사용자 요청 양식과 동일한 컬럼명으로 저장
                    df_export = final_df.copy()
                    df_export.columns = REAL_COLUMNS
                    df_export.to_excel(writer, index=False, sheet_name='서식')
                
                st.download_button(
                    label="📥 통합 수주업로드 파일 다운로드",
                    data=output.getvalue(),
                    file_name=f"통합_수주업로드_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
        st.info("'Summary' 시트가 파일에 존재하는지, 컬럼명(수량, 상품코드 등)이 일치하는지 확인해주세요.")
