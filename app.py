# [배송 정보 처리 부분 수정]
target_delivery_place = str(row.get('납품처', '')).strip()
in_type = str(row.get('입고타입', '')).strip().replace('HYPER_FLOW', 'FLOW')

# 매칭 키 생성 (공백 제거)
lookup_key = f"{target_delivery_place}{in_type}".replace(' ', '')

# 만약 매칭이 안 된다면 'NEW'를 제거하고도 찾아봄 (유연한 매칭)
shipping_code = store_dict.get(lookup_key, "")
if not shipping_code:
    flexible_key = lookup_key.replace('NEW', '') # NEW를 지우고 재시도
    shipping_code = store_dict.get(flexible_key, "")

# 그래도 없다면 화면에 경고 표시 (어떤 키가 문제인지 확인용)
if not shipping_code and target_delivery_place:
    st.warning(f"⚠️ 배송코드 매칭 실패: [{lookup_key}] - 마스터 파일을 확인하세요.")
