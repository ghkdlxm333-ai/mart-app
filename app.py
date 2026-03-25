import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, timedelta, timezone
import openpyxl

# --- 설정 ---
st.set_page_config(page_title="홈플러스(Tesco) 수주 변환 시스템", layout="wide")

# 사내 표준 양식 컬럼 (기존 코드 유지)
FINAL_COLUMNS = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 'UNIT단가', '금액', '부가세', 'LOT', '특이사항1', 'Type', '특이사항2']

def clean_val(val):
    return str(val).strip() if pd.notna(val) else ""

def process_ordview(file):
    # 1. ordview 파일 읽기 (헤더 위치 자동 탐색)
    # 보통 2번째 줄(index 1)에 '배달일시'나 '발주번호'가 시작됨
    df_raw = pd.read_excel(file, header=1) 
    
    # 2. 데이터 클렌징 (불필요한 행 제거)
    df_raw = df_raw.dropna(subset=['발주번호'])
    
    # 3. 통합 양식에 맞게 매핑
    kst = timezone(timedelta(hours=9))
    today_str = datetime.now(kst).strftime("%Y%m%d")
    
    processed_df = pd.DataFrame()
    processed_df['출고구분'] = [0] * len(df_raw)
    processed_df['수주일자'] = today_str
    # 납품일자 (YYYY-MM-DD -> YYYYMMdd)
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
    
    # 부족한 컬럼 채우기
    for col in FINAL_COLUMNS:
        if col not in processed_df.columns:
            processed_df[col] = ""
            
    return processed_df[FINAL_COLUMNS], df_raw

# --- UI ---
st.title("🏪 홈플러스(Tesco) 발주서 자동 변환기")
st.markdown("`ordview` 파일을 업로드하면 `Tesco 서식파일` 양식으로 자동 정리합니다.")

# 1. 마스터 서식파일 업로드 (서버에 고정해둘 수도 있음)
with st.expander("🛠️ 기준 서식파일 설정", expanded=False):
    master_template = st.file_uploader("Tesco 서식파일(914).xlsx 원본을 업로드하세요", type=['xlsx'])

# 2. 작업 파일 업로드
st.subheader("📥 발주서 업로드")
ordview_file = st.file_uploader("홈플러스에서 다운받은 ordview 파일을 업로드하세요", type=['xlsx', 'xls'])

if ordview_file and master_template:
    try:
        with st.spinner("데이터 변환 중..."):
            # 데이터 처리
            final_df, raw_content = process_ordview(ordview_file)
            
            # --- 엑셀 파일 생성 (기존 서식 유지하며 시트만 교체) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 기존 서식파일의 시트들을 복사하기 위해 로드
                template_workbook = openpyxl.load_workbook(master_template)
                writer.book = template_workbook
                
                # '원본 데이터' 시트가 있으면 지우고 새로 쓰기
                if '원본 데이터' in writer.book.sheetnames:
                    std = writer.book['원본 데이터']
                    writer.book.remove(std)
                
                # ordview에서 추출한 Raw Data를 '원본 데이터' 시트에 기록
                raw_content.to_excel(writer, sheet_name='원본 데이터', index=False)
                
            st.success("✅ 변환 완료!")
            
            # --- 결과 미리보기 및 다운로드 ---
            st.subheader("📊 변환 데이터 미리보기 (통합양식)")
            st.dataframe(final_df, use_container_width=True)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 변환된 Tesco 서식파일 다운로드",
                    data=output.getvalue(),
                    file_name=f"Tesco_서식적용_{datetime.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            
            with col2:
                # 사내 통합 업로드용 CSV/Excel 별도 생성 가능
                output_final = io.BytesIO()
                final_df.to_excel(output_final, index=False)
                st.download_button(
                    label="📥 통합 수주업로드용 파일 다운로드",
                    data=output_final.getvalue(),
                    file_name=f"통합수주_{datetime.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"오류 발생: {e}")
        st.info("파일의 헤더 양식이 일치하는지 확인해 주세요.")
else:
    st.info("좌측 상단의 서식파일과 중앙의 ordview 파일을 모두 업로드해주세요.")
