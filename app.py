import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from datetime import datetime, timedelta, timezone

st.set_page_config(page_title="Tesco 수주 자동 변환기", layout="wide")

# 1. 파일 이름 설정 (본인의 실제 파일명과 반드시 일치해야 함)
MASTER_TEMPLATE = "Tesco_서식파일_업데이트용.xlsx"

st.title("📂 Tesco 수주 데이터 처리")

# 파일이 없는 경우를 대비한 안전 장치
if not os.path.exists(MASTER_TEMPLATE):
    st.error(f"⚠️ 에러: '{MASTER_TEMPLATE}' 파일을 찾을 수 없습니다.")
    st.info("해결 방법: GitHub 저장소에 'Tesco_서식파일_업데이트용.xlsx' 파일을 업로드해 주세요.")
else:
    uploaded_file = st.file_uploader("M, N, X열을 삭제한 ordview 파일을 업로드하세요", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            # 데이터 읽기
            df_input = pd.read_excel(uploaded_file)
            
            # HYPER_FLOW -> FLOW 변환
            df_input = df_input.astype(str).replace('HYPER_FLOW', 'FLOW', regex=True)
            
            # 서식 파일 작업
            wb = openpyxl.load_workbook(MASTER_TEMPLATE)
            
            if '원본 데이터' not in wb.sheetnames:
                st.error("❌ 서식 파일 내에 '원본 데이터' 시트가 없습니다.")
            else:
                ws_source = wb['원본 데이터']
                # 기존 데이터 삭제 (2행부터)
                if ws_source.max_row > 1:
                    ws_source.delete_rows(2, ws_source.max_row)
                
                # 데이터 입력
                for r_idx, row in enumerate(df_input.values, 2):
                    for c_idx, value in enumerate(row, 1):
                        ws_source.cell(row=r_idx, column=c_idx, value=value)
                
                # 메모리에서 수식 계산 결과 추출
                tmp_buffer = io.BytesIO()
                wb.save(tmp_buffer)
                tmp_buffer.seek(0)
                
                wb_eval = openpyxl.load_workbook(tmp_buffer, data_only=True)
                
                if 'Summary' not in wb_eval.sheetnames:
                    st.error("❌ 서식 파일 내에 'Summary' 시트가 없습니다.")
                else:
                    ws_summary = wb_eval['Summary']
                    data = ws_summary.values
                    cols = next(data)
                    df_summary = pd.DataFrame(data, columns=cols)
                    
                    # 수량 0 제외
                    df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
                    df_valid = df_summary[df_summary['수량'] > 0].copy()

                    if not df_valid.empty:
                        # 최종 양식 정리
                        today = datetime.now(timezone(timedelta(hours=9))).strftime("%Y%m%d")
                        
                        # 화면 표시 및 다운로드 로직 (기존과 동일)
                        st.success("✅ 변환 성공!")
                        st.dataframe(df_valid)
                        
                        # (중략: 다운로드 버튼 코드...)
                        csv = df_valid.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 결과 다운로드 (CSV)", csv, f"Tesco_Result_{today}.csv", "text/csv")
                    else:
                        st.warning("⚠️ Summary 시트에 수량이 있는 데이터가 없습니다.")
        except Exception as e:
            st.error(f"🚨 상세 에러 발생: {e}")
