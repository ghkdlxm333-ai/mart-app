import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import csv
from datetime import datetime, date
from PIL import Image

# ==========================================
# ⚙️ 페이지 및 기본 설정 (Wide Layout & 탭 아이콘)
# ==========================================
try:
    img = Image.open("logo2.png")
except FileNotFoundError:
    img = "🌿"

st.set_page_config(
    page_title="멘소래담 통합 수주업로드", 
    page_icon=img, 
    layout="wide"
)

# ==========================================
# 🎨 B2B SaaS 스타일 커스텀 CSS (스트림릿 느낌 지우기)
# ==========================================
st.markdown("""
<style>
    /* 폰트 적용 (Pretendard 등 모던 폰트 룩) */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"]  {
        font-family: 'Pretendard', sans-serif !important;
    }

    /* 💡 핵심 수정: 우측 상단의 Share, Star, GitHub 등 클라우드 툴바 완벽 숨기기 */
    [data-testid="stHeaderActionElements"] {display: none !important;}
    [data-testid="stToolbar"] {display: none !important;}
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    .stDeployButton {display: none !important;}
    
    /* 💡 헤더 배경은 투명하게 해서 좌측의 사이드바 여는 버튼( > )은 보이게 유지 */
    [data-testid="stHeader"] {
        background-color: transparent !important;
    }
    
    /* 전체 배경색과 사이드바 배경색을 완벽한 흰색(#ffffff)으로 통일 */
    .stApp {
        background-color: #ffffff;
    }
    
    /* 사이드바 배경색 흰색 통일 및 아주 연한 경계선만 남김 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #f1f5f9;
    }
    
    /* 탭 디자인 변경 (하단 밑줄 스타일로 자연스럽게) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 0px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        border: none;
        border-radius: 0;
        padding: 12px 4px;
        font-weight: 500;
        color: #64748b;
        box-shadow: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: transparent;
        border: none;
        border-bottom: 3px solid #3b82f6; /* 파란색 밑줄 포인트 */
        font-weight: 700;
        color: #1e293b;
        box-shadow: none;
    }
    
    /* 메트릭 카드(통계 요약) 디자인 */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
    }
    
    /* 다운로드 버튼 스타일링 (그라데이션 & 강조) */
    .stDownloadButton button {
        width: 100%;
        border-radius: 8px;
        font-weight: 700;
        letter-spacing: 0.5px;
        background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
        color: white;
        border: none;
        padding: 12px 0;
        box-shadow: 0 4px 6px rgba(59, 130, 246, 0.25);
        transition: all 0.3s ease;
    }
    .stDownloadButton button:hover {
        background: linear-gradient(135deg, #4338ca 0%, #2563eb 100%);
        box-shadow: 0 6px 8px rgba(59, 130, 246, 0.35);
        transform: translateY(-1px);
        color: white;
    }
    
    /* 파일 업로더 디자인 개선 */
    [data-testid="stFileUploadDropzone"] {
        border-radius: 12px;
        border: 2px dashed #94a3b8;
        background-color: #ffffff;
        padding: 30px;
        transition: all 0.2s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #3b82f6;
        background-color: #f8fafc;
    }
</style>
""", unsafe_allow_html=True)
# 모든 날짜 형식을 하이픈 없이 YYYYMMDD로 통일
today_str = datetime.today().strftime("%Y%m%d")

# ==========================================
# 🎨 좌측 사이드바 (Sidebar) - 로고 단독 배치
# ==========================================
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except FileNotFoundError:
        pass
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(f"📅 **시스템 기준일:** `{today_str}`")

# ==========================================
# 📝 메인 화면 타이틀 (깔끔하게 텍스트만)
# ==========================================
st.title("통합 마트 수주 자동 변환 대시보드")
st.markdown("> **Tesco, 이마트 계열(TRD/노브랜드), 롯데마트**의 수주 데이터를 하나의 표준 양식으로 자동 병합·변환합니다.")
st.markdown("<br>", unsafe_allow_html=True)

# ⭐ 최종 통일 양식 컬럼 리스트 (구분 열 추가)
FINAL_COLUMNS = [
    '구분', '수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 
    'ME코드', '상품명', '수량', '단가', 'Total Amount'
]

def to_excel_unified(df, sheet_name="통합_수주업로드"):
    """데이터프레임을 엑셀 파일(메모리)로 변환하고 숫자 서식을 지정합니다."""
    numeric_cols = ['수량', '단가', 'Total Amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        num_format = workbook.add_format({'num_format': '#,##0'})
        center_format = workbook.add_format({'align': 'center'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#1e293b', 'font_color': 'white', 'border': 1, 'align': 'center'})
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        for col_idx, col_name in enumerate(df.columns):
            if col_name in ['수량', '단가', 'Total Amount']:
                worksheet.set_column(col_idx, col_idx, 12, num_format)
            elif col_name in ['구분', '수주날짜', '납품일자', '발주코드', '배송코드']:
                worksheet.set_column(col_idx, col_idx, 14, center_format)
            elif col_name in ['상품명', '배송처']:
                worksheet.set_column(col_idx, col_idx, 30)
            else:
                worksheet.set_column(col_idx, col_idx, 15)
    return output.getvalue()

# ==========================================
# 🗂️ 대시보드 탭 분리
# ==========================================
tab_tesco, tab_emart, tab_lotte = st.tabs(["Tesco", "이마트 (TRD/노브랜드 포함)", "롯데마트"])

# =====================================================================
# 🔴 [TAB 1] TESCO 로직
# =====================================================================
with tab_tesco:
    st.markdown("### Tesco 발주 데이터 업로드")
    
    FULL_PRODUCT_MAP = {
        8809020342310: 'ME90521CLA', 8809020342211: 'ME90521CLL', 8809020342419: 'ME90521CLS',
        8809020340804: 'ME90521MC1', 8809020340774: 'ME90521LP2', 8809020348992: 'ME90521E18',
        8809020340279: 'ME90521LR1', 8809020344444: 'ME90521EL9', 8809020344451: 'ME90521EL8',
        8809020344468: 'ME90521EL7', 8809020344192: 'ME90521EL6', 8809020344048: 'ME90521EL4',
        8809020344123: 'ME90521EL0', 8809020344239: 'ME90521E13', 8809020349821: 'ME90521CC4',
        8809020349814: 'ME90521CC2', 8809020349807: 'ME90521CC1', 8809020345212: 'ME00421186',
        8809020345236: 'ME00421183', 8809020345229: 'ME00421301', 8809020348978: 'ME00421151',
        8809020349661: 'ME90621CPS', 8809020349654: 'ME90621CPM', 8809020346516: 'ME90621AT2',
        8809020340286: 'ME00621AB5', 8809020340293: 'ME00621C21', 8809020346561: 'ME00621AT6',
        8809020346585: 'ME90621NA7', 8809020346592: 'ME90621ADI', 8809020346660: 'ME90621A07',
        8809020349425: 'ME00621A08', 8809020349685: 'ME00621AS1', 8809020349692: 'ME00621AL1',
        8809020349708: 'ME00621AR1', 8809020349715: 'ME00621AG1', 8809020349722: 'ME00621AF9',
        8809020349371: 'ME90621GK3', 8809020349418: 'ME90621GK2', 8809020349388: 'ME90621GL3',
        8809020349050: 'ME90621GLO', 8809020349067: 'ME90621GM4', 8809020349074: 'ME90621GE1',
        8809020349203: 'ME90621HCR', 8809020349098: 'ME90621HSL', 8809020349104: 'ME90621SM4',
        8809020349210: 'ME90621SCM', 8809020349166: 'ME90621GO8', 8809020349906: 'ME90621GLL',
        8809020349944: 'ME90621FGC', 8809020340200: 'ME00621H37', 8809020340217: 'ME00621H38',
        8809020340170: 'ME00621C15', 8809020340187: 'ME00621S24', 8809020340194: 'ME00621AS3',
        8809020340606: 'ME00621C22', 8809020340590: 'ME00621H44', 8809020340712: 'ME90621TC1',
        8809020341627: 'ME00621FMC', 8809020341634: 'ME00621FMR', 8809020341641: 'ME00621FBR',
        8809020341207: 'ME80421DR2', 8809020341061: 'ME81921SLL', 8809020341054: 'ME81921SVV',
        8809020341801: 'ME81921SL1', 8809020342501: 'ME90521LD9', 8809020342518: 'ME90521GT2',
        8809020342495: 'ME90521GS2', 8809020349036: 'ME00621CM5', 8809020346509: 'ME90621AFE',
        8809020349968: 'ME00621H41', 8809020342433: 'ME90621AC4', 8809020343478: 'ME00621ABN',
        8809020342525: 'ME80421DCH', 8809020343683: 'ME90521WC4', 8809020343690: 'ME90521WC5',
        8809020343706: 'ME90521WC6', 8809020344338: 'ME00621FHH', 8809020344321: 'ME90621MAM'
    }

    RAW_STORE_MAP = {
        '0903목천물류서비스센터SORTATION': 81020901, '0903목천물류서비스센터FLOW': 81020902,
        '0903목천물류서비스센터STOCK': 81020903, '0982안성ADC물류센터STOCK': 81020982,
        '0907밀양EXP센터FLOW': 81021903, '0967일죽물류서비스센터FLOW': 81021904,
        '0905기흥물류서비스센터FLOW': 81021907, '0961밀양물류센터FLOW': 81040912,
        '0961밀양물류센터STOCK': 81040913, '0906NEW함안상온물류센터FLOW': 81040912,
        '0906NEW함안상온물류센터SORTATION': 81040913, '0906NEW함안상온물류센터SORTER': 81040913,
        '0982안성ADC물류센터SORTATION': 81020980, '0982안성ADC물류센터FLOW': 81020981,
        '0970함안EXP물류센터SORTATION': 89029018, '0970함안EXP물류센터FLOW': 81040913,
        '0982안성ADC물류센터SINGLE': 81020981, '0906NEW함안상온물류센터SINGLE': 81040912,
        '0968365용인DSCDSD': 81040904, '0969남양주EXP물류센터FLOW': 81040905,
        '0968365용인DSCSTOCK': 81040904, '0969남양주EXP물류센터STOCK': 81040905,
        '0931덕평EXP물류센터FLOW': 81040906, '0934오산Exp물류센터FLOW': 81040907,
        '0935오산365물류센터STOCK': 81040908, '2001BH)영통점DSD': 81020192,
        '2002BH)강서점DSD': 81020191, '2003BH)인천송도점DSD': 81020190,
        '0934오산EXP물류센터SORTATION': 81040907, '0907밀양EXP센터SORTATION': 81021903,
        '0905기흥물류서비스센터SORTATION': 81021901, '0051강서점DSD': 81020191
    }

    NORMALIZED_STORE_MAP = {re.sub(r'^\d+', '', k).replace(" ", "").upper(): v for k, v in RAW_STORE_MAP.items()}

    file_tesco = st.file_uploader("📂 드래그 앤 드롭으로 파일을 업로드하세요 (csv/xlsx)", type=['xlsx', 'xls', 'csv'], key="tesco")

    if file_tesco:
        try:
            with st.spinner("🔄 Tesco 데이터 통합 변환 중입니다..."):
                all_rows = []
                if file_tesco.name.endswith('.csv'):
                    content = file_tesco.getvalue()
                    try: text = content.decode('utf-8-sig')
                    except: text = content.decode('cp949')
                    reader = csv.reader(io.StringIO(text))
                    all_rows = [row for row in reader]
                else:
                    df_temp = pd.read_excel(file_tesco, header=None, engine='openpyxl')
                    all_rows = df_temp.fillna('').astype(str).values.tolist()

                parsed_data = []
                col_map = {}
                for row in all_rows:
                    row_strs = [str(x).strip() for x in row]
                    if '상품코드' in row_strs and ('발주금액' in row_strs or '낱개수량' in row_strs):
                        col_map = {
                            '상품명': row_strs.index('상품명') if '상품명' in row_strs else -1,
                            '상품코드': row_strs.index('상품코드'),
                            '입고타입': row_strs.index('입고타입') if '입고타입' in row_strs else -1,
                            '수량': row_strs.index('낱개수량') if '낱개수량' in row_strs else -1,
                            '단가': row_strs.index('낱개당 단가') if '낱개당 단가' in row_strs else -1,
                            '금액': row_strs.index('발주금액') if '발주금액' in row_strs else -1,
                            '납품처': row_strs.index('납품처') if '납품처' in row_strs else -1,
                            '납품일자': row_strs.index('납품일자') if '납품일자' in row_strs else -1
                        }
                        continue
                    
                    if not col_map: continue
                    try:
                        b_idx = col_map['상품코드']
                        if b_idx >= len(row_strs): continue
                        b_str = re.sub(r'[^\d]', '', row_strs[b_idx])
                        if not b_str: continue
                        barcode = int(b_str)
                        
                        if barcode in FULL_PRODUCT_MAP:
                            def get_val(k):
                                i = col_map[k]
                                if i != -1 and i < len(row_strs):
                                    v = re.sub(r'[^\d.]', '', row_strs[i])
                                    return float(v) if v else 0.0
                                return 0.0
                            def get_str(k):
                                i = col_map[k]
                                return row_strs[i] if i != -1 and i < len(row_strs) else ''

                            parsed_data.append({
                                '상품명': get_str('상품명'), '바코드': barcode, '입고타입': get_str('입고타입'),
                                '수량': get_val('수량'), '단가': get_val('단가'), '금액': get_val('금액'),
                                '납품처': get_str('납품처'), '납품일자': get_str('납품일자')
                            })
                    except Exception: pass

                df = pd.DataFrame(parsed_data)
                df['상품코드'] = df['바코드'].map(FULL_PRODUCT_MAP)
                
                def get_store_code(row):
                    s = str(row['납품처']).replace(' ', '').upper()
                    t = str(row['입고타입']).replace(' ', '').upper()
                    if 'HYPER_FLOW' in t: t = 'FLOW'
                    elif 'MIX' in t: t = 'SORTATION'
                    s = re.sub(r'^\d+', '', s)
                    key = s + t
                    if key in NORMALIZED_STORE_MAP: return NORMALIZED_STORE_MAP[key]
                    for norm_k, code in NORMALIZED_STORE_MAP.items():
                        if norm_k in key or key in norm_k: return code
                    return 81040913
                
                df['배송코드'] = df.apply(get_store_code, axis=1)
                df['발주코드'] = 81020000
                df = df[df['수량'] > 0]
                
                groupby_cols = ['발주코드', '배송코드', '납품처', '상품코드', '상품명', '단가', '납품일자']
                df_grouped = df.groupby(groupby_cols, as_index=False).agg({'수량': 'sum', '금액': 'sum'})
                
                df_grouped['구분'] = "0"
                df_grouped['수주날짜'] = today_str
                df_grouped['납품일자'] = pd.to_datetime(df_grouped['납품일자'], errors='coerce').dt.strftime('%Y%m%d')
                df_grouped['발주처'] = 'Tesco'
                df_grouped.rename(columns={'납품처': '배송처', '상품코드': 'ME코드', '금액': 'Total Amount'}, inplace=True)
                df_final = df_grouped[FINAL_COLUMNS].copy()
                
                st.success("✨ Tesco 데이터 정제 및 병합이 완료되었습니다!")
                c1, c2, c3 = st.columns(3)
                c1.metric("📦 총 처리 건수", f"{len(df_final):,} 건")
                c2.metric("🔢 총 납품 수량", f"{df_final['수량'].sum():,.0f} 개")
                c3.metric("💰 총 납품 금액", f"{df_final['Total Amount'].sum():,.0f} 원")

                with st.expander("👀 변환된 상세 데이터 미리보기 (약 20~30줄 표시)", expanded=True):
                    st.dataframe(df_final, use_container_width=True, height=500)
                
                st.download_button(
                    label="📥 통일 양식 다운로드 (Tesco)", 
                    data=to_excel_unified(df_final), 
                    file_name=f"수주통합본_Tesco_{today_str}.xlsx", 
                    mime="application/vnd.ms-excel", key="dl_tesco",
                )
        except Exception as e:
            st.error(f"오류 발생: {e}")

# =====================================================================
# 🟡 [TAB 2] 이마트 (이마트 / 트레이더스 / 노브랜드) 로직
# =====================================================================
with tab_emart:
    st.markdown("### 이마트 (이마트/TRD/노브랜드) 발주 데이터 업로드")
    
    TEMPLATE_FILES = [
        "NEW 이마트 서식파일_20260420납품.xlsx",
        "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx",
        "NEW 노브랜드_20260409납품.xlsx"
    ]

    @st.cache_data
    def load_emart_master():
        appended = []
        for fn in TEMPLATE_FILES:
            if os.path.exists(fn):
                xls = pd.ExcelFile(fn)
                prod_sheets = [s for s in xls.sheet_names if '제품명' in s]
                if prod_sheets:
                    d = pd.read_excel(xls, sheet_name=prod_sheets[0])
                    d.columns = d.columns.str.strip()
                    appended.append(d)
        if not appended: return None
        md = pd.concat(appended, ignore_index=True)
        if '바코드' in md.columns:
            md['바코드'] = md['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            md = md.drop_duplicates(subset=['바코드'], keep='first')
        return md

    prod_df = load_emart_master()
    
    if prod_df is None or prod_df.empty:
        st.warning("⚠️ 서버에 이마트 마스터 파일(서식파일 3종)이 없습니다. 관리자에게 문의하세요.")
    else:
        file_emart = st.file_uploader("📂 드래그 앤 드롭으로 파일을 업로드하세요 (xlsx/csv)", type=['xlsx', 'xls', 'csv'], key="emart")
        
        if file_emart:
            try:
                with st.spinner("🔄 이마트 데이터 통합 변환 중입니다..."):
                    if file_emart.name.endswith('.csv'):
                        try:
                            raw_df = pd.read_csv(file_emart, encoding='utf-8-sig')
                        except:
                            file_emart.seek(0)
                            raw_df = pd.read_csv(file_emart, encoding='cp949')
                    else:
                        xls_raw = pd.ExcelFile(file_emart)
                        t_sheet = xls_raw.sheet_names[0]
                        for s in xls_raw.sheet_names:
                            temp = pd.read_excel(xls_raw, sheet_name=s, nrows=3)
                            if '점포코드' in temp.columns:
                                t_sheet = s
                                break
                        raw_df = pd.read_excel(xls_raw, sheet_name=t_sheet)

                    raw_df = raw_df.dropna(subset=['점포코드'])
                    raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)
                    raw_df['센터코드'] = raw_df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False).str.strip()
                    raw_df['수량'] = pd.to_numeric(raw_df.get('수량', 0), errors='coerce').fillna(0)
                    
                    date_col = '센터입하일자' if '센터입하일자' in raw_df.columns else ('센터입하일' if '센터입하일' in raw_df.columns else '점입점일자')
                    raw_df['배송일자'] = raw_df.get(date_col, '').astype(str).str.replace('.0', '', regex=False).str.replace('-', '', regex=False).str.strip()
                    
                    raw_df = raw_df[raw_df['수량'] > 0].copy() 

                    emart_map_dict = {
                        'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
                        'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
                        'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
                    }

                    def process_emart(row):
                        code = row['점포코드']
                        center = str(row['센터코드'])
                        if (1000 <= code <= 1999) or code >= 9000: cust = 'E-mart'
                        elif 2000 <= code <= 2999: cust = 'E-mart(TRD)'
                        elif 3000 <= code <= 3999: cust = 'E-mart(노브랜드)'
                        else: cust = 'Unknown'
                        
                        mapped_code = emart_map_dict.get(cust, {}).get(center, center)
                        return pd.Series([cust, mapped_code])

                    raw_df[['Customer', '배송코드']] = raw_df.apply(process_emart, axis=1)
                    raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
                    name_col = '상품명(기획)' if '상품명(기획)' in prod_df.columns else '상품명'
                    
                    merged_df = pd.merge(raw_df, prod_df[['바코드', '상품코드(기획)', name_col]], left_on='상품코드', right_on='바코드', how='left')
                    merged_df['최종_상품코드'] = merged_df['상품코드(기획)'].fillna(merged_df['상품코드'])
                    merged_df['최종_상품명'] = merged_df[name_col].fillna(merged_df.get('상품명', ''))

                    delivery_name_map = {
                        '81010901': '이마트 백암물류센터', 
                        '81010902': '이마트 시화물류센터', 
                        '81010903': '이마트 대구물류센터',
                        '81010905': '이마트 여주물류센터', 
                        '81010906': '이마트 광주물류센터',
                        '81010904': '이마트 노브랜드 여주2물류센터', 
                        '81010968': '이마트 노브랜드 여주물류센터',
                        '81010969': '이마트 노브랜드 시화물류센터', 
                        '89011175': '이마트 노브랜드 대구물류(신규)',
                        '81033036': '이마트 트레이더스 평택물류',
                        '89011174': '이마트 트레이더스 대구물류', 
                        '81011012': '이마트 트레이더스 여주물류',
                        '81011010': '이마트 트레이더스 시화물류'
                    }

                    merged_df['발주코드'] = '81010000'
                    merged_df['날짜'] = today_str
                    
                    merged_df['배송처'] = merged_df['배송코드'].astype(str).map(delivery_name_map).fillna(merged_df['배송코드'])
                    
                    subset_df = merged_df[[
                        '날짜', '배송일자', '발주코드', 'Customer', '배송코드', '배송처', 
                        '최종_상품코드', '최종_상품명', '수량', '발주원가', '발주금액'
                    ]].copy()
                    
                    subset_df.rename(columns={
                        '날짜': '수주날짜', '배송일자': '납품일자', 'Customer': '발주처', 
                        '최종_상품코드': 'ME코드', '최종_상품명': '상품명', '발주원가': '단가', '발주금액': 'Total Amount'
                    }, inplace=True)

                    group_cols = ['수주날짜', '납품일자', '발주코드', '발주처', '배송코드', '배송처', 'ME코드', '상품명', '단가']
                    grouped_df = subset_df.groupby(group_cols, dropna=False, as_index=False)[['수량', 'Total Amount']].sum()
                    
                    grouped_df['구분'] = "0" 
                    df_final = grouped_df[FINAL_COLUMNS].copy()
                    
                    df_final = df_final.sort_values(by=['발주처', '배송처', '상품명']).reset_index(drop=True)
                    
                    st.success("✨ 이마트 데이터 정제 및 병합이 완료되었습니다!")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("📦 총 처리 건수", f"{len(df_final):,} 건")
                    c2.metric("🔢 총 납품 수량", f"{df_final['수량'].sum():,.0f} 개")
                    c3.metric("💰 총 납품 금액", f"{df_final['Total Amount'].sum():,.0f} 원")

                    with st.expander("👀 변환된 상세 데이터 미리보기 (약 20~30줄 표시)", expanded=True):
                        st.dataframe(df_final, use_container_width=True, height=500)
                        
                    st.download_button(
                        label="📥 통일 양식 다운로드 (이마트)", 
                        data=to_excel_unified(df_final), 
                        file_name=f"수주통합본_Emart_{today_str}.xlsx", 
                        mime="application/vnd.ms-excel", key="dl_emart",
                    )
            except Exception as e:
                st.error(f"오류 발생: {e}")

# =====================================================================
# 🟢 [TAB 3] 롯데마트 로직
# =====================================================================
with tab_lotte:
    st.markdown("### 롯데마트 EDI 발주 데이터 업로드")
    
    LOTTE_TEMPLATE = '2022 롯데마트 서식파일 260417납품.xlsx'
    CENTER_CODE_MAP = {'오산센터': '81030907', '김해센터': '81030908'}

    def clean_lotte_code(val):
        s = str(val).strip()
        if s.endswith('.0'): s = s[:-2]
        return s
    
    def clean_lotte_number(val):
        s = str(val).replace(',', '').strip()
        if s.endswith('.0'): s = s[:-2]
        s = re.sub(r'[^0-9]', '', s)
        return int(s) if s else 0

    file_lotte = st.file_uploader("📂 드래그 앤 드롭으로 파일을 업로드하세요 (xls/csv)", type=['xlsx', 'csv'], key="lotte")
    
    if file_lotte:
        try:
            with st.spinner("🔄 롯데마트 데이터 통합 변환 중입니다..."):
                if file_lotte.name.endswith('.csv'): df_edi = pd.read_csv(file_lotte, header=None)
                else: df_edi = pd.read_excel(file_lotte, header=None)
                df_edi = df_edi.dropna(how='all')
                
                parsed_list, curr_center, curr_doc_no, curr_delivery_date = [], "", "", ""
                
                for i, row in df_edi.iterrows():
                    r = [str(x).strip() for x in row.tolist()]
                    if r[0] == 'ORDERS':
                        curr_doc_no = clean_lotte_code(r[1])
                        name = str(r[5]).strip()
                        curr_center = re.sub(r'상온센타|상온센터|센타', '센터', name).replace('센터센터', '센터')
                        
                        curr_delivery_date = re.sub(r'[^0-9]', '', str(r[7]) if len(r) > 7 else "") 
                        continue
                    
                    barcode = clean_lotte_code(r[1])
                    if barcode.startswith('880'):
                        qty = clean_lotte_number(r[6])
                        ipsu = clean_lotte_number(r[5]) or 1
                        u_qty = qty * ipsu
                        if u_qty > 0:
                            edi_price = clean_lotte_number(r[7] if len(r) > 7 else 0)
                            parsed_list.append({
                                '발주번호': curr_doc_no, '센터': curr_center, '납품일자': curr_delivery_date,
                                '바코드': barcode, 'EDI_품명': r[2], 'UNIT수량': u_qty, 'EDI_단가': edi_price
                            })
                            
                if not parsed_list:
                    st.warning("⚠️ 유효한 롯데마트 발주 내역이 없습니다.")
                else:
                    df_parsed = pd.DataFrame(parsed_list)
                    
                    if os.path.exists(LOTTE_TEMPLATE):
                        df_map_sheet = pd.read_excel(LOTTE_TEMPLATE, sheet_name=0)
                        df_price_sheet = pd.read_excel(LOTTE_TEMPLATE, sheet_name=1)
                        
                        m_dict = df_map_sheet[[df_map_sheet.columns[3], df_map_sheet.columns[13]]].copy()
                        m_dict.columns = ['바코드', 'ME코드']
                        m_dict['바코드'] = m_dict['바코드'].apply(clean_lotte_code)
                        m_dict = m_dict.drop_duplicates(subset=['바코드'])

                        c_me = [c for c in df_price_sheet.columns if '상품코드' in str(c) or 'ME' in str(c).upper()][0]
                        c_name = [c for c in df_price_sheet.columns if '품명' in str(c) or '상품명' in str(c)][0]
                        c_price = [c for c in df_price_sheet.columns if '단가' in str(c)][0]
                        
                        p_dict = df_price_sheet[[c_me, c_name, c_price]].dropna(subset=[c_me]).copy()
                        p_dict.columns = ['ME코드', '마스터_품명', '마스터_단가']
                        p_dict['ME코드'] = p_dict['ME코드'].apply(clean_lotte_code)
                        p_dict['마스터_단가'] = p_dict['마스터_단가'].apply(clean_lotte_number)
                        p_dict = p_dict.drop_duplicates(subset=['ME코드'])

                        df_final = pd.merge(df_parsed, m_dict, on='바코드', how='left')
                        df_final['ME코드'] = df_final['ME코드'].fillna(df_final['바코드'])
                        df_final = pd.merge(df_final, p_dict, on='ME코드', how='left')
                        df_final['품명'] = df_final['마스터_품명'].fillna(df_final['EDI_품명'])
                        df_final['UNIT단가'] = df_final['마스터_단가'].fillna(df_final['EDI_단가'])
                    else:
                        st.warning("⚠️ 롯데마트 서식파일을 찾을 수 없어 원본 EDI 단가/품명으로 산출합니다.")
                        df_final = df_parsed.copy()
                        df_final['ME코드'] = df_final['바코드']
                        df_final['품명'] = df_final['EDI_품명']
                        df_final['UNIT단가'] = df_final['EDI_단가']

                    # 롯데마트 특정 바코드 수동 맵핑
                    LOTTE_MANUAL_MAP = {'8809020342075': 'ME90621GKK', '8809020342105' : 'ME90621LL5', '8809020345229' : 'ME00421301', '8809020342037' : 'ME90621GMM',
                                       '8809020342044':'ME90621LLL', '8809020342464':'ME00621AB8'}
                    df_final['ME코드'] = df_final['바코드'].astype(str).map(LOTTE_MANUAL_MAP).fillna(df_final['ME코드'])

                    # 발주번호, 센터 등 원본을 묶어 수량 합산
                    df_grouped = df_final.groupby(['발주번호', '센터', '납품일자', 'ME코드'], as_index=False).agg({'품명': 'first', 'UNIT단가': 'first', 'UNIT수량': 'sum'})
                    
                    # 배송코드 설정
                    df_grouped['배송코드'] = df_grouped['센터'].map(CENTER_CODE_MAP).fillna(df_grouped['발주번호'])
                    
                    # 발주코드 = 배송코드
                    df_grouped['발주코드'] = df_grouped['배송코드']
                    
                    df_grouped['Total Amount'] = df_grouped['UNIT수량'] * df_grouped['UNIT단가']
                    df_grouped['구분'] = "0" 
                    df_grouped['수주날짜'] = today_str
                    
                    df_grouped.rename(columns={
                        '센터': '배송처', '품명': '상품명', 'UNIT수량': '수량', 'UNIT단가': '단가'
                    }, inplace=True)

                    # 배송처 이름으로 발주처 컬럼 통일 적용
                    df_grouped['발주처'] = df_grouped['배송처']

                    df_final = df_grouped[FINAL_COLUMNS].copy()
                    
                    st.success("✨ 롯데마트 데이터 정제 및 병합이 완료되었습니다!")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("📦 총 처리 건수", f"{len(df_final):,} 건")
                    c2.metric("🔢 총 납품 수량", f"{df_final['수량'].sum():,.0f} 개")
                    c3.metric("💰 총 납품 금액", f"{df_final['Total Amount'].sum():,.0f} 원")

                    with st.expander("👀 변환된 상세 데이터 미리보기 (약 20~30줄 표시)", expanded=True):
                        st.dataframe(df_final, use_container_width=True, height=500)
                        
                    st.download_button(
                        label="📥 통일 양식 다운로드 (롯데마트)", 
                        data=to_excel_unified(df_final), 
                        file_name=f"수주통합본_Lotte_{today_str}.xlsx", 
                        mime="application/vnd.ms-excel", key="dl_lotte",
                    )
        except Exception as e:
            st.error(f"오류 발생: {e}")
