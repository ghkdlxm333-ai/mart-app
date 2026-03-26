import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. 수주 업로드 최종 양식 (통합 수주 1행 기준)
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
]

def process_data(raw_df):
    # '원본 데이터' 시트에 복붙한 것과 동일한 효과를 내기 위한 전처리
    # 수량 0인 로우 제외
    df = raw_df[raw_df['발주수량'] > 0].copy()
    
    # Summary 시트의 로직을 기반으로 값 배치
    res_df = pd.DataFrame()
    res_df['출고구분'] = [0] * len(df)
    res_df['수주일자'] = datetime.now().strftime('%Y%m%d')
    res_df['납품일자'] = df['납품일자'].astype(str).str.replace('-', '').str[:8]
    res_df['발주처코드'] = df['납품처코드']
    res_df['발주처'] = df['납품처']
    res_df['배송코드'] = df['배송처코드']
    res_df['배송지'] = df['배송처']
    res_df['상품코드'] = df['상품코드']
    res_df['상품명'] = df['상품명']
    res_df['UNIT수량'] = df['발주수량'].astype(int)
    res_df['UNIT단가'] = df['낱개당 단가'].astype(int) # 이 부분의 오류를 수정했습니다.
    res_df['금        액'] = df['발주금액'].astype(int)
    res_df['부  가   세'] = (df['발주금액'] * 0.1).astype(int)
    res_df['Type'] = '마트'
    
    # 누락된 컬럼 빈값으로 채우기
    for col in FINAL_COLUMNS:
        if col not in res_df.columns:
            res_df[col] = ''
            
    return res_df[FINAL_COLUMNS]

# Streamlit UI 부분은 기존과 동일하게 유지...
