import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from datetime import datetime, timedelta, timezone

# --- 1. 페이지 및 환경 설정 ---
st.set_page_config(page_title="Tesco 수주 자동 변환 시스템", layout="wide")

# 고정된 서식 파일 이름
MASTER_TEMPLATE = "Tesco 서식파일_업데이트용.xlsx"

# 최종 통합 수주업로드 양식 컬럼 (사용자 제공 양식 기준)
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

# 엑셀 저장 시 실제 헤더명 (중복 방지 및 공백 유지)
EXPORT_HEADERS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
]

# --- 2. 데이터 처리 함수 ---
def process_data(uploaded_ordview):
    # (1) ordview 파일 읽기 (헤더 위치 반영)
    df = pd.read_excel(uploaded_ordview, header=1)
    
    # (2) M, N, X열 삭제 (0부터 시작하는 인덱스 기준: M=12, N=13, X=23)
    # 컬럼 이름이 유동적일 수 있으므로 인덱스로 삭제하는 것이 안전합니다.
    cols_to_drop = [12, 13, 23]
    df.drop(df.columns[cols_to_drop], axis=1, inplace=True)
    
    # (3) 'HYPER_FLOW' -> 'FLOW' 변환
    df = df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # (4) Tesco 서식파일_업데이트용 로드
    wb = openpyxl.load_workbook(MASTER_TEMPLATE)
    
    # (5) '원본 데이터' 시트에 가공된 데이터 복붙
    if '원본 데이터' in wb.sheetnames:
        ws_source = wb['원본 데이터']
        # 기존 내용 삭제 (헤더 제외)
        for row in ws_source.iter_rows(min_row=2):
            for cell in row:
                cell.value = None
        
        # 새 데이터 쓰기
        for r_idx, row in enumerate(df.values, 2):
            for c_idx, value in enumerate(row, 1):
                ws_source.cell(row=r_idx, column=c_idx, value=value)
    
    # (6) 파일 저장 후 다시 읽어 Summary 시트 데이터 추출
    # 수식이 계산된 결과를 가져오기 위해 임시 버퍼 사용
    tmp_buffer = io.BytesIO()
    wb.save(tmp_buffer)
    tmp_buffer.seek(0)
    
    # 수식 결과값을 읽기 위해 data_only=True로 로드
    wb_result = openpyxl.load_workbook(tmp_buffer, data_only=True)
    ws_summary = wb_result['Summary']
    
    # Summary 시트를 데이터프레임으로 변환
    data = ws_summary.values
    cols = next(data)
    df_summary = pd.DataFrame(data, columns=cols)
    
    # (7) 수량이 0인 행 제외 및 통합 양식 정리
    df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
    df_final_data = df_summary[df_summary['수량'] > 0].copy()
    
    return df_final_data

# --- 3. 메인 UI ---
st.title("🚀 Tesco 수주 업무 자동화 (Summary 기반)")

if not os.path.exists(MASTER_TEMPLATE):
    st.error(f"❌ '{MASTER_TEMPLATE}' 파일이 서버에 없습니다. GitHub에 업로드해주세요.")
else:
    ordview_file = st.file_uploader("홈플러스 ordview 파일을 업로드하세요", type=['xlsx', 'xls'])

    if ordview_file:
        try:
            with st.spinner("처리 중... (열 삭제 및 FLOW 변환 적용)"):
                summary_data = process_data(ordview_file)
                
                if summary_data.empty:
                    st.warning("⚠️ Summary 시트에 수량이 있는 데이터가 없습니다.")
                else:
                    # 통합 수주업로드 파일 양식 생성
                    kst = timezone(timedelta(hours=9))
                    today = datetime.now(kst).strftime("%Y%m%d")
                    
                    res_df = pd.DataFrame(columns=FINAL_COLUMNS)
                    res_df['출고구분'] = 0
                    res_df['수주일자'] = today
                    res_df['납품일자'] = today # 필요시 데이터에서 추출
                    res_df['발주처코드'] = summary_data['발주코드']
                    res_df['발주처'] = "홈플러스"
                    res_df['배송코드'] = summary_data['배송코드']
                    res_df['상품코드'] = summary_data['상품코드']
                    res_df['상품명'] = summary_data['상품명']
                    res_df['UNIT수량'] = summary_data['수량'].astype(int)
                    res_df['UNIT단가'] = pd.to_numeric(summary_data['UNIT단가'], errors='coerce').fillna(0).astype(int)
                    res_df['금        액'] = res_df['UNIT수량'] * res_df['UNIT단가']
                    res_df['부  가   세'] = (res_df['금        액'] * 0.1).astype(int)
                    res_df.fillna('', inplace=True)
                    
                    st.success(f"✅ 변환 완료! (총 {len(res_df)}건)")
                    st.dataframe(res_df, use_container_width=True)
                    
                    # 최종 엑셀 다운로드 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        export_df = res_df.copy()
                        export_df.columns = EXPORT_HEADERS
                        export_df.to_excel(writer, index=False, sheet_name='서식')
                    
                    st.download_button(
                        label="📥 통합 수주업로드 파일 다운로드",
                        data=output.getvalue(),
                        file_name=f"통합_수주업로드_{today}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
        except Exception as e:
            st.error(f"오류 발생: {e}")

st.caption("Developed by Jay | SCM Outbound Process Automation")
