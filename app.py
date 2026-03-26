import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 1. 설정 및 KST 시간 ---
st.set_page_config(page_title="Tesco 수주 마스터 시스템", layout="wide")
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

# GitHub에 저장된 마스터 파일명
MASTER_FILE_NAME = "Tesco_서식파일_업데이트용.xlsx"

@st.cache_data
def load_fixed_master():
    """마스터 파일의 각 시트 데이터를 로드합니다."""
    if not os.path.exists(MASTER_FILE_NAME):
        return None, None, f"❌ '{MASTER_FILE_NAME}' 파일을 찾을 수 없습니다."
    
    try:
        # 상품코드 매핑 (바코드 숫자를 키로 사용)
        df_item = pd.read_excel(MASTER_FILE_NAME, sheet_name='상품코드', dtype={'바코드': str, 'ME코드': str})
        df_item = df_item.dropna(subset=['바코드']).drop_duplicates(subset=['바코드'])
        item_map = df_item.set_index('바코드')['ME코드'].to_dict()
        
        # 발주처코드 매핑
        df_store = pd.read_excel(MASTER_FILE_NAME, sheet_name='Tesco 발주처코드', dtype=str)
        df_store = df_store.dropna(subset=['납품처&타입']).drop_duplicates(subset=['납품처&타입'])
        store_map = df_store.set_index('납품처&타입')[['발주처코드', '배송코드', '배송지']].to_dict('index')
        
        return item_map, store_map, None
    except Exception as e:
        return None, None, f"❌ 마스터 로드 오류: {e}"

# --- 2. 데이터 처리 함수 ---
def process_full_cycle(ordview_file, item_map, store_map):
    # [Step 1] Raw 데이터 읽기 및 변환
    raw_df = pd.read_excel(ordview_file, header=1, dtype=str)
    raw_df.columns = raw_df.columns.str.strip()
    
    # 'HYPER_FLOW' -> 'FLOW' 변환
    raw_df = raw_df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # 상품코드 숫자형 변환 (바코드 매핑을 위해)
    raw_df['상품코드'] = pd.to_numeric(raw_df['상품코드'], errors='coerce').astype(str).str.split('.').str[0]
    
    # 수치 데이터 변환
    raw_df['낱개수량'] = pd.to_numeric(raw_df['낱개수량'], errors='coerce').fillna(0)
    raw_df['낱개당 단가'] = pd.to_numeric(raw_df['낱개당 단가'], errors='coerce').fillna(0)
    raw_df['발주금액'] = pd.to_numeric(raw_df['발주금액'], errors='coerce').fillna(0)

    # [Step 2] 합산 로직 (Summary 시트 산출물 생성)
    # 발주처 + 상품코드 + 단가별로 그룹화하여 수량 합산
    summary_grouped = raw_df.groupby([
        '납품처코드', '납품처', '배송처코드', '배송처', '상품코드', '상품명', '납품일자
