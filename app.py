import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 1. 설정 및 KST 시간 ---
st.set_page_config(page_title="Tesco 수주 합산 시스템", layout="wide")
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

# --- 2. 상수 설정 (제공해주신 파일 1행 양식과 동일하게 설정) ---
# '금        액', '부  가   세'의 공백까지 일치시켰습니다.
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
]

MASTER_FILE_NAME = "Tesco_서식파일_업데이트용.xlsx"

@st.cache_data
def load_fixed_master():
    """GitHub 레포지토리의 마스터 파일을 로드하고 중복을 제거합니다."""
    if not os.path.exists(MASTER_FILE_NAME):
        return None, None, f"❌ '{MASTER_FILE_NAME}' 파일을 찾을 수 없습니다."
    
    try:
        # 상품코드 매핑 (중복 제거)
        df_item = pd.read_excel(MASTER_FILE_NAME, sheet_name='상품코드', dtype=str)
        df_item = df_item.dropna(subset=['바코드']).drop_duplicates(subset=['바코드'])
        item_map = df_item.set_index('바코드')['ME코드'].to_dict()
        
        # 발주처코드 매핑 (중복 제거 - Index Unique 오류 방지)
        df_store = pd.read_excel(MASTER_FILE_NAME, sheet_name='Tesco 발주처코드', dtype=str)
        df_store = df_store.dropna(subset=['납품처&타입']).drop_duplicates(subset=['납품처&타입'])
        store_map = df_store.set_index('납품처&타입')[['발주처코드', '배송코드', '배송지']].to_dict('index')
        
        return item_map, store_map, None
    except Exception as e:
        return None, None, f"❌ 마스터 로드 오류: {e}"

# --- 3. 데이터 처리 로직 ---
def process_mart_data(ordview_file, item_map, store_map):
    # ordview 읽기 (헤더는 2행에 있음)
    df = pd.read_excel(ordview_file, header=1, dtype=str)
    df.columns = df.columns.str.strip()
    
    # 'HYPER_FLOW' -> 'FLOW' 변환
    df = df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # 숫자 변환
    df['낱개수량'] = pd.to_numeric(df['낱개수량'], errors='coerce').fillna(0)
    df['낱개당 단가'] = pd.to_numeric(df['낱개당 단가'], errors='coerce').fillna(0)
    df['발주금액'] = pd.to_numeric(df['발주금액'], errors='coerce').fillna(0)
    
    # [핵심] 같은 발주처 & 같은 상품코드 & 같은 납품일자일 경우 합계 계산
    # 요약 기준에 '낱개당 단가'를 포함시켜 단가가 다른 경우 분리되도록 함
    summary = df.groupby([
        '납품처코드', '납품처', '배송처코드', '배송처', '상품코드', '상품명', '납품일자', '낱개당 단가', '입고타입'
    ], as_index=False).agg({
        '낱개수량': 'sum',
        '발주금액': 'sum'
    })
    
    # 수량 0 제외
    summary = summary[summary['낱개수량'] > 0]
    
    res_list = []
    for _, row in summary.iterrows():
        # 마스터 매핑 키
