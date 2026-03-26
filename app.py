import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, timedelta, timezone

# --- 1. 설정 및 KST 시간 ---
st.set_page_config(page_title="Tesco 수주 합산 시스템", layout="wide")
kst = timezone(timedelta(hours=9))
today_str = datetime.now(kst).strftime("%Y%m%d")

# --- 2. 마스터 파일 및 양식 설정 ---
MASTER_FILE_NAME = "Tesco_서식파일_업데이트용.xlsx"

# 중복 오류 해결: '특이사항' 뒤에 번호를 붙여 중복을 피함 (나중에 출력 시에는 조정 가능)
FINAL_COLUMNS = [
    '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
    '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
    'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항1', 'Type', '특이사항2'
]

@st.cache_data
def load_fixed_master():
    """GitHub의 마스터 파일을 로드하고 매핑 데이터를 생성합니다."""
    if not os.path.exists(MASTER_FILE_NAME):
        return None, None, None, f"❌ '{MASTER_FILE_NAME}' 파일을 찾을 수 없습니다."
    
    try:
        # 1. 상품코드 매핑
        df_item = pd.read_excel(MASTER_FILE_NAME, sheet_name='상품코드', dtype=str)
        df_item = df_item.dropna(subset=['바코드']).drop_duplicates(subset=['바코드'])
        item_map = df_item.set_index('바코드')['ME코드'].to_dict()
        
        # 2. 발주처코드 매핑
        df_store = pd.read_excel(MASTER_FILE_NAME, sheet_name='Tesco 발주처코드', dtype=str)
        df_store = df_store.dropna(subset=['납품처&타입']).drop_duplicates(subset=['납품처&타입'])
        store_map = df_store.set_index('납품처&타입')[['발주처코드', '배송코드', '배송지']].to_dict('index')
        
        # 3. Summary 시트 양식 확인 (컬럼 구조 파악용)
        df_summary_template = pd.read_excel(MASTER_FILE_NAME, sheet_name='Summary', dtype=str).head(0)
        
        return item_map, store_map, df_summary_template, None
    except Exception as e:
        return None, None, None, f"❌ 마스터 로드 오류: {e}"

# --- 3. 데이터 처리 로직 ---
def process_mart_data(ordview_file, item_map, store_map):
    # ordview 읽기 (헤더 2행)
    df = pd.read_excel(ordview_file, header=1, dtype=str)
    df.columns = df.columns.str.strip()
    
    # 'HYPER_FLOW' -> 'FLOW' 변환
    df = df.replace('HYPER_FLOW', 'FLOW', regex=True)
    
    # 숫자 변환
    df['낱개수량'] = pd.to_numeric(df['낱개수량'], errors='coerce').fillna(0)
    df['낱개당 단가'] = pd.to_numeric(df['낱개당 단가'], errors='coerce').fillna(0)
    df['발주금액'] = pd.to_numeric(df['발주금액'], errors='coerce').fillna(0)
    
    # [핵심] 발주처 + 상품코드 + 납품일자 기준 합산
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
        m_key = f"{row['납품처코드']} {row['납품처']}{row['입고타입']}"
        store_info = store_map.get(m_key, {})
        
        # 사내 상품코드 매핑
        internal_code = item_map.get(row['상품코드'], row['상품코드'])
        
        res_list.append({
            '출고구분': '0',
            '수주일자': today_str,
            '납품일자': str(row['납품일자']).replace('-', '')[:8],
            '발주처코드': store_info.get('발주처코드', row['납품처코드']),
            '발주처': row['납품처'],
            '배송코드': store_info.get('배송코드', row['배송처코드']),
            '배송지': store_info.get('배송지', row['배송처']),
            '상품코드': internal_code,
            '상품명': row['상품명'],
            'UNIT수량': int(row['낱개수량']),
            'UNIT단가': int(row['낱개당 단가']),
            '금        액': int(row['발주금액']),
            '부  가   세': int(row['발주금액'] * 0.1),
            'LOT': '',
            '특이사항1': '',
            'Type': row['입고타입'],
            '특이사항2': ''
        })
    
    result_df = pd.DataFrame(res_list)
    
    # 최종 결과물의 컬럼명을 서식과 동일하게 복구 (중복 이름 허용을 위해 리스트로 재설정)
    final_output_cols = [
        '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
        '배송코드', '배송지', '상품코드', '상품명', 'UNIT수량', 
        'UNIT단가', '금        액', '부  가   세', 'LOT', '특이사항', 'Type', '특이사항'
    ]
    
    return df, result_df, final_output_cols

# --- 4. 메인 UI ---
st.title("🛒 Tesco 수주 데이터 Summary 합산기")

item_map, store_map, _, error_msg = load_fixed_master()

if error_msg:
    st.error(error_msg)
else:
    st.success("✅ 마스터(Summary 양식 포함) 로드 완료")
    
    ordview_file = st.file_uploader("ordview(Raw) 파일을 업로드하세요", type=['xlsx', 'xls'])

    if ordview_file:
        try:
            raw_transformed, final_df, final_cols = process_mart_data(ordview_file, item_map, store_map)
            
            st.divider()
            st.subheader("📊 Summary 기준 합산 결과")
            st.dataframe(final_df, use_container_width=True)
            
            # 엑셀 다운로드 (중복 컬럼명을 위해 전처리)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 데이터프레임의 컬럼을 강제로 서식용 컬럼명으로 덮어씌움
                final_df.columns = final_cols 
                final_df.to_excel(writer, index=False, sheet_name='서식')
                raw_transformed.to_excel(writer, index=False, sheet_name='FLOW변환원본')
            
            st.download_button(
                label="📥 통합 수주 파일 다운로드",
                data=output.getvalue(),
                file_name=f"Tesco_Summary_합산_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ 처리 오류: {e}")
