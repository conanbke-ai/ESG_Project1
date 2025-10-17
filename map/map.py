import os
import pandas as pd
import folium
import requests
import time
from branca.colormap import linear
import base64
from io import BytesIO
import matplotlib.pyplot as plt
from tqdm import tqdm

# === 1️⃣ 경로 설정 ===
FILE_PATH = r"C:\ESG_Project1\file\generator_file\HOME_발전설비_발전기별.xlsx"
OUTPUT_HTML = r"C:\ESG_Project1\map\generator_interactive_chart.html"

# === 2️⃣ Kakao API Key ===
KAKAO_API_KEY = "93c089f75a2730af2f15c01838e892d3"

# === 3️⃣ 엑셀 불러오기 ===
df = pd.read_excel(FILE_PATH)
region_col = '광역지역'
subregion_col = '세부지역'

# === 4️⃣ 시도 단위 추출 및 NaN 처리 ===
df['시도'] = df[region_col].astype(str).str.strip().str.extract(r'^(.*?[시도])')[0]
df['시도'] = df['시도'].fillna('알수없음').astype(str).str.replace(" ", "")

# === 5️⃣ 시도별 발전소 수 집계 ===
sido_counts = df['시도'].value_counts().reset_index()
sido_counts.columns = ['시도', '발전소수']
total_count = sido_counts['발전소수'].sum()
max_count = sido_counts['발전소수'].max()

# === 6️⃣ 지도 초기화 ===
m = folium.Map(location=[36.5, 127.8], zoom_start=7)

# === 7️⃣ 컬러맵 생성 ===
colormap = linear.YlOrRd_09.scale(0, max_count)
colormap.caption = '시도별 발전소 수'
colormap.add_to(m)

# === 8️⃣ 시도별 파이 차트 이미지 생성 함수 ===
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

# === 9️⃣ Kakao 좌표 조회 함수 ===
def get_coords_kakao(address):
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    params = {"query": address}
    response = requests.get(url, headers=headers, params=params)
    result = response.json()
    try:
        x = float(result['documents'][0]['x'])
        y = float(result['documents'][0]['y'])
        return float(y), float(x)  # folium은 (lat, lon)
    except:
        return 36.5, 127.8  # 기본값

# === 10️⃣ 발전소별 Marker 추가 ===
coords_cache = {}

for idx, row in tqdm(df.iterrows(), total=len(df), desc="발전소 좌표 조회"):
    sido_name = str(row['시도']).replace(" ", "")
    
    # 주소 생성 (광역지역 + 세부지역)
    region_str = str(row[region_col]) if pd.notna(row[region_col]) else ""
    subregion_str = str(row[subregion_col]) if pd.notna(row[subregion_col]) else ""
    address = f"{region_str} {subregion_str}".strip()
    
    # 좌표 조회 (캐싱)
    if address in coords_cache:
        lat, lon = coords_cache[address]
    else:
        lat, lon = get_coords_kakao(address)
        coords_cache[address] = (lat, lon)
        time.sleep(0.2)

    # Marker 색상 & 크기
    count = int(sido_counts[sido_counts['시도'] == sido_name]['발전소수'].values[0])
    radius = 5 + (count / max_count) * 25
    color = colormap(count)

    # 파이 차트 HTML
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

# === 11️⃣ 결과 저장 ===
m.save(OUTPUT_HTML)
print(f"✅ Kakao 기반 지도 + 파이 차트 팝업 생성 완료: {OUTPUT_HTML}")