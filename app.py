# load_master_data 함수 내부 수정
# 키를 생성할 때 공백을 모두 없애고 저장합니다.
store_map = {}
for _, r in df_store.iterrows():
    key = str(r['납품처&타입']).replace(" ", "").strip() # 공백 제거
    val = str(r['배송코드']).strip()
    if key and val:
        store_map[key] = val

# ... (중간 생략) ...

# 배송코드 매칭 로직 수정
raw_place = str(row.get('납품처', '')).strip()
raw_type = str(row.get('입고타입', '')).strip()

# 1. 타입 변환
converted_type = raw_type.replace('HYPER_', '')

# 2. 매칭용 키 생성 (양쪽 모두 공백을 제거하여 비교)
# 예: "0906NEW함안상온물류센터FLOW" (공백 없음)
matching_key = (raw_place + converted_type).replace(" ", "")

# 3. 배송코드 찾기
shipping_code = store_map.get(matching_key, "")
