import os
import json
import pandas as pd
import folium
import requests
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from branca.colormap import linear
import base64
from io import BytesIO
import matplotlib.pyplot as plt

# === 경로 설정 ===
FILE_PATH = r"C:\ESG_Project1\file\generator_file\HOME_발전설비_발전기별.xlsx"
OUTPUT_HTML = r"C:\ESG_Project1\map\generator_interactive_chart.html"
CACHE_FILE = r"C:\ESG_Project1\map\coords_cache.json"

KAKAO_API_KEY = "93c089f75a2730af2f15c01838e892d3"

# === 데이터 불러오기 ===
df = pd.read_excel(FILE_PATH)
region_col = '광역지역'
subregion_col = '세부지역'
df['시도'] = df[region_col].astype(str).str.extract(r'^(.*?[시도])')[0]
df['시도'] = df['시도'].fillna('알수없음').astype(str).str.replace(" ", "")

# === 시도별 통계 ===
sido_counts = df['시도'].value_counts().reset_index()
sido_counts.columns = ['시도', '발전소수']
total_count = sido_counts['발전소수'].sum()
max_count = sido_counts['발전소수'].max()

# === 캐시 불러오기 ===
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        coords_cache = json.load(f)
else:
    coords_cache = {}

# === 고유 주소 목록 추출 ===
df['주소'] = (df[region_col].astype(str) + " " + df[subregion_col].astype(str)).str.strip()
unique_addresses = df['주소'].dropna().unique()
print(f"총 {len(unique_addresses)}개의 고유 주소 중 {len(coords_cache)}개 캐시됨.")

# === Kakao 좌표 조회 함수 ===
def get_coords_kakao(address):
    if address in coords_cache:
        return address, coords_cache[address]

    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    params = {"query": address}

    try:
        response = requests.get(url, headers=headers, params=params, timeout=5)
        result = response.json()
        x = float(result['documents'][0]['x'])
        y = float(result['documents'][0]['y'])
        coords_cache[address] = [y, x]
    except Exception:
        coords_cache[address] = [36.5, 127.8]  # 기본 좌표
    return address, coords_cache[address]

# === 병렬 처리 (멀티스레드) ===
addresses_to_fetch = [a for a in unique_addresses if a not in coords_cache]
print(f"API 요청 대상: {len(addresses_to_fetch)}건")

with ThreadPoolExecutor(max_workers=5) as executor:
    futures = [executor.submit(get_coords_kakao, addr) for addr in addresses_to_fetch]
    for i, f in enumerate(as_completed(futures), 1):
        addr, coord = f.result()
        if i % 50 == 0:
            print(f"진행률: {i}/{len(addresses_to_fetch)} ({i/len(addresses_to_fetch)*100:.1f}%)")
        time.sleep(0.1)  # API rate 제한 방지

# === 캐시 저장 ===
with open(CACHE_FILE, "w", encoding="utf-8") as f:
    json.dump(coords_cache, f, ensure_ascii=False, indent=2)
print(f"좌표 캐시 저장 완료 ({len(coords_cache)}개)")

# === 지도 초기화 ===
m = folium.Map(location=[36.5, 127.8], zoom_start=7)
colormap = linear.YlOrRd_09.scale(0, max_count)
colormap.add_to(m)

# === 파이 차트 함수 ===
def make_pie_chart(sido_name):
    row = sido_counts[sido_counts['시도'] == sido_name].iloc[0]
    labels = ['해당 시도', '나머지']
    sizes = [row['발전소수'], total_count - row['발전소수']]
    fig, ax = plt.subplots(figsize=(2,2))
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=['#ff9999','#dddddd'])
    ax.axis('equal')
    buf = BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode('utf-8')

# === Marker 추가 ===
for _, row in df.iterrows():
    address = row['주소']
    lat, lon = coords_cache.get(address, [36.5, 127.8])
    sido_name = str(row['시도']).replace(" ", "")
    count = int(sido_counts[sido_counts['시도'] == sido_name]['발전소수'].values[0])
    radius = 5 + (count / max_count) * 25
    color = colormap(count)
    img_base64 = make_pie_chart(sido_name)
    html = f"""
        <h4>{sido_name}</h4>
        발전소 수: {count}개<br>
        전체 대비: {count/total_count*100:.2f}%<br>
        <img src="data:image/png;base64,{img_base64}" width="150">
    """
    folium.CircleMarker(
        location=[lat, lon],
        radius=radius,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.7,
        popup=folium.Popup(html, max_width=300)
    ).add_to(m)

m.save(OUTPUT_HTML)
print(f"✅ 지도 생성 완료: {OUTPUT_HTML}")