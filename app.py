import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from datetime import datetime, timedelta, timezone

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="Tesco 수주 자동 변환기", layout="wide")

# 서버(GitHub)에 저장된 서식 파일 이름
MASTER_TEMPLATE = "Tesco_서식파일_업데이트용.xlsx"

# 최종 결과 양식 컬럼
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

st.title("📂 Tesco 수주 자동화 시스템")

# [파일 존재 여부 확인]
if not os.path.exists(MASTER_TEMPLATE):
    st.error(f"❌ 서버에서 '{MASTER_TEMPLATE}' 파일을 찾을 수 없습니다. GitHub에 파일을 업로드했는지 확인하세요.")
else:
    # 1. ordview 로우데이터 업로드
    uploaded_file = st.file_uploader("ordview 원본 파일을 업로드하세요 (수기 수정 불필요)", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            with st.spinner("데이터 변환 및 수식 계산 중..."):
                # (1) 로우데이터 읽기 (상단 불필요한 행 자동 스킵)
                # 데이터가 실제 시작되는 '배달일시' 행을 찾기 위해 시트 전체를 읽음
                df_raw = pd.read_excel(uploaded_file)
                
                # '배달일시'라는 글자가 포함된 행을 찾아 헤더로 재설정
                header_idx = None
                for i, row in df_raw.iterrows():
                    if "배달일시" in row.values:
                        header_idx = i
                        break
                
                if header_idx is not None:
                    df_raw = pd.read_excel(uploaded_file, skiprows=header_idx + 1)
                
                # (2) 'HYPER_FLOW' -> 'FLOW' 변환
                df_processed = df_raw.astype(str).replace('HYPER_FLOW', 'FLOW', regex=True)
                
                # (3) 서식 파일 로드 및 데이터 주입
                wb = openpyxl.load_workbook(MASTER_TEMPLATE)
                if '원본 데이터' in wb.sheetnames:
                    ws_source = wb['원본 데이터']
                    # 기존 데이터 삭제 (2행부터 끝까지)
                    if ws_source.max_row > 1:
                        ws_source.delete_rows(2, ws_source.max_row + 500)
                    
                    # 변환된 데이터 쓰기
                    for r_idx, row in enumerate(df_processed.values, 2):
                        for c_idx, value in enumerate(row, 1):
                            # 숫자로 변환 가능한 데이터는 숫자로 입력 (수식 계산용)
                            try:
                                if "." in str(value): ws_source.cell(row=r_idx, column=c_idx, value=float(value))
                                else: ws_source.cell(row=r_idx, column=c_idx, value=int(value))
                            except:
                                ws_source.cell(row=r_idx, column=c_idx, value=value)
                
                # (4) 수식 계산 결과를 읽기 위해 임시 저장 및 재로드
                tmp_buffer = io.BytesIO()
                wb.save(tmp_buffer)
                tmp_buffer.seek(0)
                
                # data_only=True로 로드 (Summary 시트의 결과값만 가져옴)
                wb_eval = openpyxl.load_workbook(tmp_buffer, data_only=True)
                
                if 'Summary' in wb_eval.sheetnames:
                    ws_summary = wb_eval['Summary']
                    summary_data = ws_summary.values
                    summary_cols = next(summary_data)
                    df_summary = pd.DataFrame(summary_data, columns=summary_cols)
                    
                    # (5) 수량 0 제외 필터링
                    if '수량' in df_summary.columns:
                        df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
                        df_valid = df_summary[df_summary['수량'] > 0].copy()
                        
                        if not df_valid.empty:
                            today = datetime.now(timezone(timedelta(hours=9))).strftime("%Y%m%d")
                            
                            # 최종 결과 양식 구성
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
                            
                            u_price = pd.to_numeric(df_valid.get('UNIT단가', 0), errors='coerce').fillna(0)
                            res_df['UNIT단가'] = u_price.astype(int)
                            res_df['금        액'] = (res_df['UNIT수량'] * res_df['UNIT단가']).astype(int)
                            res_df['부  가   세'] = (res_df['금        액'] * 0.1).astype(int)
                            res_df.fillna('', inplace=True)

                            st.success(f"✅ 변환 완료! ({len(res_df)}건 추출)")
                            st.dataframe(res_df)

                            # 엑셀 다운로드 버튼
                            out_excel = io.BytesIO()
                            with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
                                res_df.to_excel(writer, index=False, sheet_name='서식')
                            
                            st.download_button(
                                label="📥 통합 수주업로드 파일 다운로드",
                                data=out_excel.getvalue(),
                                file_name=f"Tesco_Upload_{today}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                        else:
                            st.warning("⚠️ Summary 시트에 수량이 있는 행이 없습니다.")
                    else:
                        st.error("❌ Summary 시트에서 '수량' 열을 찾을 수 없습니다.")
                else:
                    st.error("❌ 서식 파일에 'Summary' 시트가 없습니다.")
                    
        except Exception as e:
            st.error(f"🚨 상세 오류: {e}")
