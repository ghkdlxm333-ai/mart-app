import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 1. 설정 및 KST 시간 ---
st.set_page_config(page_title="Tesco 수주 자동화", layout="wide")
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

# --- 2. 마스터 파일 경로 설정 ---
# GitHub에 업로드된 파일 이름을 아래와 똑같이 맞춰주세요.
MASTER_FILE_NAME = "Tesco_서식파일_업데이트용.xlsx"

@st.cache_data
def load_fixed_master():
    """GitHub 레포지토리에 포함된 마스터 파일을 자동으로 로드합니다."""
    if not os.path.exists(MASTER_FILE_NAME):
        return None, None, f"❌ '{MASTER_FILE_NAME}' 파일을 찾을 수 없습니다. GitHub 레포지토리에 파일을 올려주세요."
    
    try:
        # 상품코드 매핑 (중복 제거 로직 추가)
        df_item = pd.read_excel(MASTER_FILE_NAME, sheet_name='상품코드', dtype=str)
        df_item = df_item.dropna(subset=['바코드']).drop_duplicates(subset=['바코드'], keep='first')
        item_map = df_item.set_index('바코드')['ME코드'].to_dict()
        
        # 발주처코드 매핑 (중복 제거 로직 추가 - 오류 발생 지점 해결)
        df_store = pd.read_excel(MASTER_FILE_NAME, sheet_name='Tesco 발주처코드', dtype=str)
        df_store = df_store.dropna(subset=['납품처&타입']).drop_duplicates(subset=['납품처&타입'], keep='first')
        store_map = df_store.set_index('납품처&타입')[['발주처코드', '배송코드', '배송지']].to_dict('index')
        
        return item_map, store_map, None
    except Exception as e:
        return None, None, f"❌ 마스터 파일 로드 중 오류: {e}"

# --- 3. 데이터 처리 로직 ---
def process_mart_data(ordview_file, item_map, store_map):
    # ordview 읽기 (첫 줄 '주문서LIST' 제외)
    df = pd.read_excel(ordview_file, header=1, dtype=str)
    df.columns = df.columns.str.strip()
    
    # 'HYPER_FLOW' -> 'FLOW' 변환
    df = df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # 숫자 변환 및 전처리
    df['낱개수량'] = pd.to_numeric(df['낱개수량'], errors='coerce').fillna(0)
    df['낱개당 단가'] = pd.to_numeric(df['낱개당 단가'], errors='coerce').fillna(0)
    df['발주금액'] = pd.to_numeric(df['발주금액'], errors='coerce').fillna(0)
    
    # Summary (중복 합계)
    summary = df.groupby([
        '납품처코드', '납품처', '배송처코드', '배송처', '상품코드', '상품명', '납품일자', '낱개당 단가', '입고타입'
    ], as_index=False).agg({'낱개수량': 'sum', '발주금액': 'sum'})
    
    summary = summary[summary['낱개수량'] > 0]
    
    # 통합 양식 생성
    res_list = []
    for _, row in summary.iterrows():
        # 매핑 키 생성
        m_key = f"{row['납품처코드']} {row['납품처']}{row['입고타입']}"
        store_info = store_map.get(m_key, {})
        
        res_list.append({
            '출고구분': '0',
            '수주일자': today_str,
            '납품일자': str(row['납품일자']).replace('-', '')[:8],
            '발주처코드': store_info.get('발주처코드', row['납품처코드']),
            '발주처': row['납품처'],
            '배송코드': store_info.get('배송코드', row['배송처코드']),
            '배송지': store_info.get('배송지', row['배송처']),
            '상품코드': item_map.get(row['상품코드'], row['상품코드']),
            '상품명': row['상품명'],
            'UNIT수량': int(row['낱개수량']),
            'UNIT단가': int(row['낱개당 단가']),
            '금액': int(row['발주금액']),
            '부가세': int(row['발주금액'] * 0.1),
            'Type': row['입고타입']
        })
    
    return df, pd.DataFrame(res_list)

# --- 4. 메인 UI ---
st.title("🏪 Tesco(홈플러스) 수주 자동화")

# 마스터 파일 자동 로드
item_map, store_map, error_msg = load_fixed_master()

if error_msg:
    st.error(error_msg)
    st.info("💡 GitHub 레포지토리에 'Tesco_서식파일_업데이트용.xlsx' 파일이 있는지 확인해 주세요.")
else:
    st.success("✅ 마스터 서식(상품/발주처 코드) 자동 로드 완료")
    
    ordview_file = st.file_uploader("ordview(Raw) 파일을 업로드하세요", type=['xlsx', 'xls'])

    if ordview_file:
        try:
            raw_transformed, final_df = process_mart_data(ordview_file, item_map, store_map)
            
            st.divider()
            st.subheader("📊 통합 수주업로드 결과")
            st.dataframe(final_df, use_container_width=True)
            
            # 엑셀 다운로드 파일 생성
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='서식')
                raw_transformed.to_excel(writer, index=False, sheet_name='원본데이터_FLOW변환')
            
            st.download_button(
                label="📥 통합 수주 파일 다운로드",
                data=output.getvalue(),
                file_name=f"Tesco_통합수주_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ 데이터 처리 중 오류가 발생했습니다: {e}")
