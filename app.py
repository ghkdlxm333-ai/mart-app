import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="Tesco 수주 통합 변환 시스템", layout="wide")

# 고정된 서식 파일 이름 (GitHub에 업로드한 파일명과 일치해야 함)
TEMPLATE_FILE = "Tesco 서식파일(914)_260325납품(싱글타입확인)"

# 사용자가 요청한 '통합 수주업로드' 최종 컬럼 양식 (공백까지 정확히 일치)
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

# 엑셀 저장 시 실제 출력될 헤더 (특이사항 중복 처리 대응)
REAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
]

# --- 2. 메인 UI ---
st.title("🏪 Tesco Summary → 통합 수주업로드 변환")
st.markdown("`ordview` 파일을 업로드하면, 서식파일의 **Summary 시트** 내용을 바탕으로 통합 업로드 파일을 생성합니다.")

# 서식 파일 존재 확인
if not os.path.exists(TEMPLATE_FILE):
    st.error(f"❌ 서버에서 '{TEMPLATE_FILE}' 파일을 찾을 수 없습니다. GitHub에 파일을 업로드해 주세요.")
else:
    # 파일 업로더
    uploaded_file = st.file_uploader("홈플러스 ordview 엑셀 파일을 업로드하세요", type=['xlsx', 'xls'])

    if uploaded_file:
        try:
            with st.spinner("Summary 시트 데이터를 분석 중입니다..."):
                # 1. 고정된 서식 파일의 'Summary' 시트 읽기
                # (주의: 사용자가 업로드한 ordview가 아니라, 
                # ordview 업로드 시점에 서버에 있는 'Tesco_Template.xlsx'의 Summary를 읽음)
                df_summary = pd.read_excel(TEMPLATE_FILE, sheet_name='Summary', dtype=str)
                
                # 컬럼명 앞뒤 공백 제거
                df_summary.columns = df_summary.columns.str.strip()
                
                # 2. 수량 데이터 숫자 변환 및 0인 행 제외
                df_summary['수량'] = pd.to_numeric(df_summary['수량'], errors='coerce').fillna(0)
                df_filtered = df_summary[df_summary['수량'] > 0].copy()
                
                if df_filtered.empty:
                    st.warning("⚠️ 'Summary' 시트에 수량이 0보다 큰 데이터가 없습니다. 서식 파일의 수식을 확인해 주세요.")
                else:
                    # 3. 통합 수주업로드 형식으로 데이터 매핑
                    kst = timezone(timedelta(hours=9))
                    today_str = datetime.now(kst).strftime("%Y%m%d")
                    
                    # 납품일자 추출 (Summary 시트에 납품일 컬럼이 있다면 사용, 없다면 오늘 날짜)
                    # 파일 데이터의 특성에 따라 '납품일' 또는 '납품일자' 컬럼명을 확인해야 함
                    deliv_date = today_str
                    if '납품일' in df_filtered.columns:
                        deliv_date = df_filtered['납품일'].iloc[0]
                    
                    res_df = pd.DataFrame(index=df_filtered.index, columns=FINAL_COLUMNS)
                    
                    res_df['출고구분'] = 0
                    res_df['수주일자'] = today_str
                    res_df['납품일자'] = deliv_date
                    res_df['발주처코드'] = df_filtered['발주코드'].fillna('')
                    res_df['발주처'] = "홈플러스" # 필요 시 '발주처' 컬럼에서 가져오도록 수정 가능
                    res_df['배송코드'] = df_filtered['배송코드'].fillna('')
                    res_df['배송지'] = df_filtered.get('배송지', '') # 배송지 컬럼이 있으면 가져옴
                    res_df['상품코드'] = df_filtered['상품코드'].fillna('')
                    res_df['상품명'] = df_filtered['상품명'].fillna('')
                    res_df['UNIT수량'] = df_filtered['수량'].astype(int)
                    res_df['UNIT단가'] = pd.to_numeric(df_filtered['UNIT단가'], errors='coerce').fillna(0).astype(int)
                    
                    # 금액 및 부가세 계산
                    res_df['금        액'] = res_df['UNIT수량'] * res_df['UNIT단가']
                    res_df['부  가   세'] = (res_df['금        액'] * 0.1).astype(int)
                    
                    # 나머지 빈 값 처리
                    res_df.fillna('', inplace=True)
                    
                    # 4. 화면 표시 및 다운로드
                    st.success(f"✅ Summary 시트에서 {len(res_df)}건의 유효 데이터를 추출했습니다.")
                    st.dataframe(res_df, use_container_width=True)
                    
                    # 엑셀 파일 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # 최종 출력은 REAL_COLUMNS 헤더 사용
                        export_df = res_df.copy()
                        export_df.columns = REAL_COLUMNS
                        export_df.to_excel(writer, index=False, sheet_name='서식')
                    
                    st.download_button(
                        label="📥 통합 수주업로드용 파일 다운로드",
                        data=output.getvalue(),
                        file_name=f"Tesco_통합수주_{today_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.info("서식 파일의 'Summary' 시트 명칭과 컬럼명(상품코드, 수량, UNIT단가 등)을 확인해 주세요.")

st.divider()
st.caption("Developed by Jay | SCM Outbound Automation")
