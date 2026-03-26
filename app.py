import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta, timezone

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="마트(Tesco) 수주 변환 시스템", layout="wide")

st.title("🛒 마트(Tesco) 수주업로드 자동화")
st.markdown("`ordview` 파일을 업로드하면 'FLOW' 변환 및 통합 양식으로 재구성합니다.")

# --- 2. 상수 설정 ---
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금액', '부가세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

# 한국 시간 설정
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

# --- 3. 데이터 처리 함수 ---
def process_tesco_orders(uploaded_file):
    # 1. 파일 읽기 (엑셀/CSV 대응)
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, dtype=str)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)
    
    # 공백 제거
    df.columns = df.columns.str.strip()
    
    # 2. 'HYPER_FLOW' -> 'FLOW' 변환
    # 모든 컬럼에 대해 변환을 수행하거나 특정 컬럼(예: 입고타입, 센터구분) 지정 가능
    df = df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # 3. 데이터 요약 (발주처/상품코드별 수량 및 단가 합계)
    # 수량과 단가는 숫자로 변환
    df['낱개수량'] = pd.to_numeric(df['낱개수량'], errors='coerce').fillna(0)
    df['낱개당 단가'] = pd.to_numeric(df['낱개당 단가'], errors='coerce').fillna(0)
    df['발주금액'] = pd.to_numeric(df['발주금액'], errors='coerce').fillna(0)
    
    # 요약 기준: 납품처코드, 납품처, 배송처코드, 배송처, 상품코드, 상품명, 납품일자
    # (Tesco 서식의 Summary 로직 적용)
    summary_df = df.groupby([
        '납품처코드', '납품처', '배송처코드', '배송처', '상품코드', '상품명', '납품일자', '낱개당 단가'
    ], as_index=False).agg({
        '낱개수량': 'sum',
        '발주금액': 'sum'
    })
    
    # 4. 수량 0 제외
    summary_df = summary_df[summary_df['낱개수량'] > 0]
    
    # 5. 통합 수주업로드 형식으로 변환
    output_df = pd.DataFrame(columns=FINAL_COLUMNS)
    output_df['출고구분'] = '0'
    output_df['수주일자'] = today_str
    output_df['납품일자'] = summary_df['납품일자'].str.replace(r'\D', '', regex=True).str[:8]
    output_df['발주처코드'] = summary_df['납품처코드']
    output_df['발주처'] = summary_df['납품처']
    output_df['배송코드'] = summary_df['배송처코드']
    output_df['배송지'] = summary_df['배송처']
    output_df['상품코드'] = summary_df['상품코드']
    output_df['상품명'] = summary_df['상품명']
    output_df['UNIT수량'] = summary_df['낱개수량'].astype(int)
    output_df['UNIT단가'] = summary_df['낱개당 단가'].astype(int)
    output_df['금액'] = summary_df['발주금액'].astype(int)
    output_df['부가세'] = (output_df['금액'] * 0.1).astype(int)
    
    return df, output_df

# --- 4. 메인 화면 ---
uploaded_file = st.file_uploader("ordview (Raw 데이터) 파일을 업로드하세요", type=['xlsx', 'xls', 'csv'])

if uploaded_file:
    try:
        with st.spinner("데이터를 변환하고 요약하는 중..."):
            raw_transformed, final_result = process_tesco_orders(uploaded_file)
        
        st.success("✅ 변환 완료!")
        
        # 미리보기
        tab1, tab2 = st.tabs(["📊 통합 업로드 양식 (결과)", "📝 원본 데이터 (FLOW 변환본)"])
        
        with tab1:
            st.subheader(f"총 {len(final_result)}건의 수주 데이터")
            st.dataframe(final_result, use_container_width=True)
            
            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_result.to_excel(writer, index=False, sheet_name='서식')
            
            st.download_button(
                label="📥 통합 수주업로드 파일 다운로드",
                data=output.getvalue(),
                file_name=f"Tesco_통합수주_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with tab2:
            st.caption("HYPER_FLOW가 FLOW로 변환된 원본 데이터입니다.")
            st.dataframe(raw_transformed.head(100)) # 상위 100건만 출력

    except Exception as e:
        st.error(f"파일 처리 중 오류가 발생했습니다. 파일 형식을 확인해주세요. \n 에러 메시지: {e}")

# --- 5. 가이드 ---
st.divider()
with st.expander("💡 사용 가이드"):
    st.markdown("""
    1. 마트 사이트에서 다운로드한 `ordview` 파일을 그대로 업로드합니다.
    2. 프로그램이 자동으로 'HYPER_FLOW' 단어를 'FLOW'로 바꿉니다.
    3. 동일한 배송지/상품코드/단가를 가진 항목들을 하나로 합칩니다(Summary).
    4. 수량이 0인 데이터는 자동으로 필터링됩니다.
    5. '통합 수주업로드' 버튼을 눌러 엑셀을 받은 뒤 사내 시스템에 붙여넣으세요.
    """)
