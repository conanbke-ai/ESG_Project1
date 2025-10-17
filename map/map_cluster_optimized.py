import os
import json
import pandas as pd
import folium
from folium.plugins import MarkerCluster
import requests
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from branca.colormap import linear
import base64
from io import BytesIO
import matplotlib.pyplot as plt
from tqdm import tqdm

# === 경로 설정 ===
FILE_PATH = r"C:\ESG_Project1\file\generator_file\HOME_발전설비_발전기별.xlsx"
OUTPUT_HTML = r"C:\ESG_Project1\map\generator_map.html"
CACHE_FILE = r"C:\ESG_Project1\map\coords_cache.json"
KAKAO_API_KEY = "93c089f75a2730af2f15c01838e892d3"

# === 데이터 불러오기 ===
df = pd.read_excel(FILE_PATH)

# 마지막 줄 제거
df = df.iloc[:-1]

region_col = '광역지역'
subregion_col = '세부지역'

df['시도'] = df[region_col].astype(str).str.extract(r'^(.*?[시도])')[0]
df['시도'] = df['시도'].fillna('알수없음').astype(str).str.replace(" ", "")
df['주소'] = (df[region_col].astype(str) + " " + df[subregion_col].astype(str)).str.strip()

# 시도별 통계
sido_counts = df['시도'].value_counts().reset_index()
sido_counts.columns = ['시도', '발전소수']
total_count = sido_counts['발전소수'].sum()
max_count = sido_counts['발전소수'].max()

# 좌표 캐시
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        coords_cache = json.load(f)
else:
    coords_cache = {}

def get_coords_kakao(address):
    if address in coords_cache:
        return address, coords_cache[address]
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    params = {"query": address}
    try:
        response = requests.get(url, headers=headers, params=params, timeout=5)
        result = response.json()
        if result['documents']:
            x = float(result['documents'][0]['x'])
            y = float(result['documents'][0]['y'])
            coords_cache[address] = [y, x]
        else:
            coords_cache[address] = [36.5, 127.8]
    except:
        coords_cache[address] = [36.5, 127.8]
    return address, coords_cache[address]

# 병렬 좌표 조회
unique_addresses = df['주소'].dropna().unique()
addresses_to_fetch = [a for a in unique_addresses if a not in coords_cache]
if addresses_to_fetch:
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(get_coords_kakao, addr) for addr in addresses_to_fetch]
        for f in tqdm(as_completed(futures), total=len(futures), desc="좌표 변환 중"):
            f.result()
            time.sleep(0.05)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(coords_cache, f, ensure_ascii=False, indent=2)

# 좌표 병합
df['coords'] = df['주소'].map(coords_cache)
df[['위도', '경도']] = pd.DataFrame(df['coords'].tolist(), index=df.index)
df = df.dropna(subset=['위도', '경도'])

# 지도 초기화
m = folium.Map(location=[36.5, 127.8], zoom_start=7, tiles="CartoDB positron")
colormap = linear.YlOrRd_09.scale(0, max_count)
colormap.caption = '시도별 발전소 수'
colormap.add_to(m)

# 시도별 대표 좌표
representative_addresses = (
    df.groupby('시도')[[region_col, subregion_col]]
      .first()
      .assign(주소=lambda x: (x[region_col].astype(str) + " " + x[subregion_col].astype(str)).str.strip())
)
representative_addresses['coords'] = representative_addresses['주소'].map(coords_cache)
sido_coords = representative_addresses['coords'].to_dict()

# 파이차트 캐시
pie_cache = {}
for _, row in sido_counts.iterrows():
    sido_name = row['시도']
    sizes = [row['발전소수'], total_count - row['발전소수']]
    fig, ax = plt.subplots(figsize=(1.8,1.8))
    ax.pie(sizes, labels=['해당 시도','기타'], autopct='%1.1f%%', startangle=90, colors=['#ff6666','#dddddd'])
    ax.axis('equal')
    buf = BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    pie_cache[sido_name] = base64.b64encode(buf.read()).decode('utf-8')

# 시도별 마커 레이어
sido_layer = folium.FeatureGroup(name="시도 요약", show=True)
for _, row in sido_counts.iterrows():
    sido = row['시도']
    coord = sido_coords.get(sido, [36.5,127.8])
    if not coord or len(coord)!=2:
        continue
    lat, lon = coord
    count = row['발전소수']
    color = colormap(count)
    radius = 8 + (count / max_count)*20
    img_base64 = pie_cache[sido]
    html = f"""
        <h4>{sido}</h4>
        발전소 수: {count}개<br>
        전체 대비: {count/total_count*100:.2f}%<br>
        <img src="data:image/png;base64,{img_base64}" width="120">
    """
    folium.CircleMarker(
        location=[lat, lon],
        radius=radius,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.9,
        popup=folium.Popup(html, max_width=300)
    ).add_to(sido_layer)
sido_layer.add_to(m)

# 세부 발전소 레이어 (클러스터)
plant_layer = folium.FeatureGroup(name="세부 발전소", show=False)
marker_cluster = MarkerCluster(disableClusteringAtZoom=10).add_to(plant_layer)
for _, row in df.iterrows():
    lat, lon = row['위도'], row['경도']
    popup_html = f"""
        <b>회사명:</b> {row.get('회사명','정보없음')}<br>
        <b>발전기명:</b> {row.get('발전기명','정보없음')}<br>
        <b>위치:</b> {row[region_col]} {row[subregion_col]}<br>
        <b>발전원:</b> {row.get('발전원','정보없음')}<br>
        <b>설비용량:</b> {row.get('설비용량','정보없음')} MW
    """
    folium.Marker(
        location=[lat, lon],
        popup=folium.Popup(popup_html, max_width=300),
        icon=folium.Icon(color='blue', icon='bolt', prefix='fa')
    ).add_to(marker_cluster)
plant_layer.add_to(m)

# 레이어 컨트롤
folium.LayerControl().add_to(m)

# JS: 줌 레벨 토글
zoom_toggle_js = """
function toggleLayers(e){
    var map = e.target;
    var zoom = map.getZoom();
    var sidoLayer = map._layers[Object.keys(map._layers).filter(function(k){return map._layers[k].options && map._layers[k].options.name==='시도 요약'})[0]];
    var plantLayer = map._layers[Object.keys(map._layers).filter(function(k){return map._layers[k].options && map._layers[k].options.name==='세부 발전소'})[0]];
    if(zoom>=10){
        if(map.hasLayer(sidoLayer)) map.removeLayer(sidoLayer);
        if(!map.hasLayer(plantLayer)) map.addLayer(plantLayer);
    }else{
        if(!map.hasLayer(sidoLayer)) map.addLayer(sidoLayer);
        if(map.hasLayer(plantLayer)) map.removeLayer(plantLayer);
    }
}
map.on('zoomend', toggleLayers);
"""
m.get_root().html.add_child(folium.Element(f"<script>{zoom_toggle_js}</script>"))

# 지도 iframe 스타일 조정 (가로 200px, 세로 400px, 중앙 정렬)
map_style = """
<style>
    .folium-map {
        width: 200px !important;
        height: 400px !important;
        margin: 0 auto;
    }
    .folium-map .leaflet-container {
        width: 100% !important;
        height: 100% !important;
    }
</style>
<script>
    var maps = document.getElementsByClassName('folium-map');
    for(var i=0;i<maps.length;i++){
        maps[i].style.width='200px';
        maps[i].style.height='400px';
        maps[i].style.margin='0 auto';
    }
</script>
"""
m.get_root().html.add_child(folium.Element(map_style))

# 정상적으로 파일 저장
m.save(OUTPUT_HTML)
print(f"✅ 지도 저장 완료: {OUTPUT_HTML}")