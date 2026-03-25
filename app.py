import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from datetime import datetime, timedelta, timezone

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="Tesco 수주 자동 변환 시스템", layout="wide")

# 고정된 서식 파일 이름
MASTER_TEMPLATE = "Tesco 서식파일_업데이트용.xlsx"

# 최종 통합 수주업로드 양식 컬럼 (공백 포함)
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

EXPORT_HEADERS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
]

# --- 2. 메인 UI ---
st.title("📂 Tesco 수주 데이터 처리 (FLOW 자동 변환)")
st.info("M, N, X열이 삭제된 ordview 파일을 업로드하세요. 'HYPER_FLOW'는 자동으로 'FLOW'로 변환됩니다.")

if not os.path.exists(MASTER_TEMPLATE):
    st.error(f"❌ 서버에서 '{MASTER_TEMPLATE}' 파일을 찾을 수 없습니다.")
else:
    # 1. 파일 업로드
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            with st.spinner("데이터 변환 및 Summary 추출 중..."):
                # (1) 업로드된 파일 읽기
                df_input = pd.read_excel(uploaded_file)
                
                # (2) 'HYPER_FLOW' -> 'FLOW' 변환 로직 추가
                df_input = df_input.replace('HYPER_FLOW', 'FLOW', regex=True)
                
                # (3) 서식 파일 로드 및 '원본 데이터' 시트 업데이트
                wb = openpyxl.load_workbook(MASTER_TEMPLATE)
                if '원본 데이터' in wb.sheetnames:
                    ws_source = wb['원본 데이터']
                    
                    # 기존 데이터 삭제 (헤더 제외)
                    for row in ws_source.iter_rows(min_row=2):
                        for cell in row:
                            cell.value = None
                    
                    # 변환된 데이터 쓰기
                    for r_idx, row in enumerate(df_input.values, 2):
                        for c_idx, value in enumerate(row, 1):
                            ws_source.cell(row=r_idx, column=c_idx, value=value)
                
                # (4) 수식 계산 결과를 읽기 위해 임시 저장 후 재로드
                tmp_buffer = io.BytesIO()
                wb.save(tmp_buffer)
                tmp_buffer.seek(0)
                
                # data_only=True로 로드하여 Summary 시트의 수식 결과값만 가져옴
                wb_eval = openpyxl.load_workbook(tmp_buffer, data_only=True)
                ws_summary = wb_eval['Summary']
                
                # Summary 시트 읽기
                data = ws_summary.values
                cols = next(data)
                df_summary = pd.DataFrame(data, columns=cols)
                
                # (5) 수량 0 제외 필터링 및 통합 양식 생성
                df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
                df_valid = df_summary[df_summary['수량'] > 0].copy()

                if df_valid.empty:
                    st.warning("⚠️ 'Summary' 시트에 유효한 수량 데이터가 없습니다.")
                else:
                    # 현재 날짜 설정
                    kst = timezone(timedelta(hours=9))
                    today = datetime.now(kst).strftime("%Y%m%d")
                    
                    # 통합 수주업로드 데이터프레임 구성
                    res_df = pd.DataFrame(columns=FINAL_COLUMNS)
                    res_df['출고구분'] = 0
                    res_df['수주일자'] = today
                    res_df['납품일자'] = today 
                    res_df['발주처코드'] = df_valid['발주코드'].fillna('')
                    res_df['발주처'] = "홈플러스"
                    res_df['배송코드'] = df_valid['배송코드'].fillna('')
                    res_df['상품코드'] = df_valid['상품코드'].fillna('')
                    res_df['상품명'] = df_valid['상품명'].fillna('')
                    res_df['UNIT수량'] = df_valid['수량'].astype(int)
                    res_df['UNIT단가'] = pd.to_numeric(df_valid['UNIT단가'], errors='coerce').fillna(0).astype(int)
                    res_df['금        액'] = res_df['UNIT수량'] * res_df['UNIT단가']
                    res_df['부  가   세'] = (res_df['금        액'] * 0.1).astype(int)
                    res_df.fillna('', inplace=True)

                    # 결과 화면 표시
                    st.success(f"✅ 변환 완료: 'HYPER_FLOW'를 'FLOW'로 변경하고 {len(res_df)}건을 추출했습니다.")
                    st.dataframe(res_df, use_container_width=True)
                    
                    # 다운로드 파일 생성
                    final_output = io.BytesIO()
                    with pd.ExcelWriter(final_output, engine='xlsxwriter') as writer:
                        export_df = res_df.copy()
                        export_df.columns = EXPORT_HEADERS
                        export_df.to_excel(writer, index=False, sheet_name='서식')
                    
                    st.download_button(
                        label="📥 통합 수주업로드 파일 다운로드",
                        data=final_output.getvalue(),
                        file_name=f"통합수주_{today}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.info("서식 파일의 시트 이름이나 컬럼 항목이 일치하는지 확인해 주세요.")
