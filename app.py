import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from datetime import datetime, timedelta, timezone

# --- 페이지 설정 ---
st.set_page_config(page_title="Tesco 수주 자동 변환기", layout="wide")

# GitHub 저장소 내 파일명 (파일명이 다르면 여기서 수정)
MASTER_TEMPLATE = "Tesco_서식파일_업데이트용.xlsx"

# 최종 결과 양식 컬럼 순서
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

st.title("📂 Tesco 수주 자동화 시스템")

# 서식 파일 존재 확인
if not os.path.exists(MASTER_TEMPLATE):
    st.error(f"❌ '{MASTER_TEMPLATE}' 파일을 찾을 수 없습니다. GitHub 업로드 상태를 확인하세요.")
else:
    uploaded_file = st.file_uploader("ordview 원본 파일을 업로드하세요", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            with st.spinner("데이터 처리 중..."):
                # 1. ordview 읽기 (헤더 행 자동 찾기)
                # '배달일시'가 있는 행을 찾을 때까지 최대 10행을 검사합니다.
                df_temp = pd.read_excel(uploaded_file, header=None, nrows=10)
                header_row = 0
                for i, row in df_temp.iterrows():
                    if "배달일시" in row.values:
                        header_row = i
                        break
                
                # 실제 데이터 로드
                df_raw = pd.read_excel(uploaded_file, skiprows=header_row)
                
                # 2. 'HYPER_FLOW' -> 'FLOW' 변환
                df_processed = df_raw.copy()
                df_processed = df_processed.astype(str).replace('HYPER_FLOW', 'FLOW', regex=True)
                
                # 3. 서식 파일 작업
                wb = openpyxl.load_workbook(MASTER_TEMPLATE)
                
                if '원본 데이터' in wb.sheetnames:
                    ws_source = wb['원본 데이터']
                    # 데이터 영역(2행부터) 초기화
                    if ws_source.max_row >= 2:
                        ws_source.delete_rows(2, ws_source.max_row)
                    
                    # 처리된 데이터 쓰기
                    for r_idx, row in enumerate(df_processed.values, 2):
                        for c_idx, value in enumerate(row, 1):
                            # 숫자로 변환 가능한 값은 숫자로 넣어 수식 연산 지원
                            try:
                                if "." in str(value): ws_source.cell(row=r_idx, column=c_idx, value=float(value))
                                else: ws_source.cell(row=r_idx, column=c_idx, value=int(value))
                            except:
                                ws_source.cell(row=r_idx, column=c_idx, value=value)
                
                # 4. 수식 계산 결과 추출 (임시 저장 후 재로드)
                tmp_buffer = io.BytesIO()
                wb.save(tmp_buffer)
                tmp_buffer.seek(0)
                
                # data_only=True로 계산된 값만 가져옴
                wb_eval = openpyxl.load_workbook(tmp_buffer, data_only=True)
                
                if 'Summary' in wb_eval.sheetnames:
                    ws_summary = wb_eval['Summary']
                    data = list(ws_summary.values)
                    
                    # Summary 시트가 비어있지 않은지 확인
                    if len(data) > 1:
                        summary_cols = data[0]
                        df_summary = pd.DataFrame(data[1:], columns=summary_cols)
                        
                        # 5. 수량 0 제외 필터링
                        if '수량' in df_summary.columns:
                            df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
                            df_valid = df_summary[df_summary['수량'] > 0].copy()
                            
                            if not df_valid.empty:
                                kst = timezone(timedelta(hours=9))
                                today = datetime.now(kst).strftime("%Y%m%d")
                                
                                # 최종 양식 조립
                                res_df = pd.DataFrame(columns=FINAL_COLUMNS)
                                res_df['출고구분'] = 0
                                res_df['수주일자'] = today
                                res_df['납품일자'] = (datetime.now(kst) + timedelta(days=1)).strftime("%Y%m%d")
                                res_df['발주처코드'] = df_valid.get('발주코드', '').fillna('')
                                res_df['발주처'] = "홈플러스"
                                res_df['배송코드'] = df_valid.get('배송코드', '').fillna('')
                                res_df['상품코드'] = df_valid.get('상품코드', '').fillna('')
                                res_df['상품명'] = df_valid.get('상품명', '').fillna('')
                                res_df['UNIT수량'] = df_valid['수량'].astype(int)
                                
                                u_price = pd.to_numeric(df_valid.get('UNIT단가', 0), errors='coerce').fillna(0)
                                res_df['UNIT단가'] = u_price.astype(int
