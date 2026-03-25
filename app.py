import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from datetime import datetime, timedelta, timezone

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="Tesco 수주 자동 변환기", layout="wide")

# 고정된 서식 파일 이름 (GitHub에 업로드된 파일명과 일치해야 함)
MASTER_TEMPLATE = "Tesco_서식파일_업데이트용.xlsx"

# 최종 통합 수주업로드 양식 컬럼 순서
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

# --- 2. 메인 UI ---
st.title("📂 Tesco 수주 자동화 시스템")
st.markdown("### 로우데이터(ordview) 업로드만으로 전처리와 Summary 추출을 한 번에 처리합니다.")

# 서식 파일 체크
if not os.path.exists(MASTER_TEMPLATE):
    st.error(f"❌ 서버에서 서식 파일을 찾을 수 없습니다: {MASTER_TEMPLATE}")
    st.info("GitHub 저장소에 해당 엑셀 파일을 업로드했는지 확인해주세요.")
else:
    # 1. 로우데이터 업로드 (전처리 전의 순수 ordview 파일)
    uploaded_file = st.file_uploader("ordview 파일을 업로드하세요", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            with st.spinner("데이터 처리 중... (HYPER_FLOW 변환 및 Summary 추출)"):
                # (1) 로우데이터 읽기
                # ordview 파일은 보통 1번 행(인덱스 0)부터 데이터가 시작되거나 상단에 제목이 있을 수 있음
                # 여기서는 '배달일시'가 포함된 행을 헤더로 찾도록 구성
                df_raw = pd.read_excel(uploaded_file)
                
                # (2) 전처리: 'HYPER_FLOW' -> 'FLOW' 변환
                # 모든 열에서 해당 글자를 찾아 변환합니다.
                df_processed = df_raw.replace('HYPER_FLOW', 'FLOW', regex=True)
                
                # (3) 서식 파일의 '원본 데이터' 시트에 복사
                wb = openpyxl.load_workbook(MASTER_TEMPLATE)
                if '원본 데이터' not in wb.sheetnames:
                    st.error("❌ 서식 파일에 '원본 데이터' 시트가 없습니다.")
                else:
                    ws_source = wb['원본 데이터']
                    # 기존 데이터 삭제 (헤더인 1행 제외하고 2행부터 삭제)
                    if ws_source.max_row > 1:
                        ws_source.delete_rows(2, ws_source.max_row + 100) # 여유있게 삭제
                    
                    # 전처리된 데이터 입력 (2행부터)
                    for r_idx, row in enumerate(df_processed.values, 2):
                        for c_idx, value in enumerate(row, 1):
                            ws_source.cell(row=r_idx, column=c_idx, value=value)
                
                # (4) 엑셀 내부 수식 결과 추출을 위해 메모리에 저장 후 재로드
                tmp_buffer = io.BytesIO()
                wb.save(tmp_buffer)
                tmp_buffer.seek(0)
                
                # data_only=True: 수식이 아닌 계산된 '값'을 가져옴
                wb_eval = openpyxl.load_workbook(tmp_buffer, data_only=True)
                
                if 'Summary' not in wb_eval.sheetnames:
                    st.error("❌ 서식 파일에 'Summary' 시트가 없습니다.")
                else:
                    ws_summary = wb_eval['Summary']
                    summary_data = ws_summary.values
                    summary_cols = next(summary_data)
                    df_summary = pd.DataFrame(summary_data, columns=summary_cols)
                    
                    # (5) 수량 0 제외 및 통합 양식 생성
                    # '수량' 컬럼 숫자로 변환
                    df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
                    df_valid = df_summary[df_summary['수량'] > 0].copy()

                    if not df_valid.empty:
                        # 오늘/내일 날짜 설정
                        kst = timezone(timedelta(hours=9))
                        today = datetime.now(kst).strftime("%Y%m%d")
                        
                        # 결과 데이터프레임 구성
                        res_df = pd.DataFrame(columns=FINAL_COLUMNS)
                        res_df['출고구분'] = 0
                        res_df['수주일자'] = today
                        res_df['납품일자'] = today 
                        res_df['발주처코드'] = df_valid.get('발주코드', '')
                        res_df['발주처'] = "홈플러스"
                        res_df['배송코드'] = df_valid.get('배송코드', '')
                        res_df['상품코드'] = df_valid.get('상품코드', '')
                        res_df['상품명'] = df_valid.get('상품명', '')
                        res_df['UNIT수량'] = df_valid['수량'].astype(int)
                        
                        # 단가 및 금액 계산
                        u_price = pd.to_numeric(df_valid.get('UNIT단가', 0), errors='coerce').fillna(0)
                        res_df['UNIT단가'] = u_price.astype(int)
                        res_df['금        액'] = (res_df['UNIT수량'] * res_df['UNIT단가']).astype(int)
                        res_df['부  가   세'] = (res_df['금        액'] * 0.1).astype(int)
                        res_df.fillna('', inplace=True)

                        # 결과 화면 출력
                        st.success(f"✅ 처리가 완료되었습니다. (추출 데이터: {len(res_df)}건)")
                        st.dataframe(res_df, use_container_width=True)
                        
                        # 엑셀 파일로 변환하여 다운로드 버튼 생성
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            res_df.to_excel(writer, index=False, sheet_name='서식')
                        
                        st.download_button(
                            label="📥 통합 수주업로드 파일 다운로드",
                            data=output.getvalue(),
                            file_name=f"통합수주_{today}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                    else:
                        st.warning("⚠️ 'Summary' 시트에 수량이 0보다 큰 데이터가 없습니다. 서식파일의 수식을 확인해 주세요.")

        except Exception as e:
            st.error(f"🚨 오류 발생: {e}")
            st.info("파일의 시트 이름이 '원본 데이터', 'Summary'가 맞는지, 그리고 로우데이터의 형식이 이전과 동일한지 확인해 주세요.")
