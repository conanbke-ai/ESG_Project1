import os, json, time, re, colorsys, logging, base64, webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import requests
from tqdm import tqdm
import folium
from folium import FeatureGroup, LayerControl
import matplotlib.pyplot as plt
from matplotlib import rc
from io import BytesIO

logging.basicConfig(level=logging.INFO, format='[%(asctime)s]✅ %(message)s')

FILE_PATH   = r"C:\ESG_Project1\file\generator_file\HOME_발전설비_발전기별.xlsx"
CACHE_FILE  = r"C:\ESG_Project1\map\coord_cache.json"
OUTPUT_HTML = r"C:\ESG_Project1\map\solar_dashboard.html"
GEOJSON_FILE = r"C:\ESG_Project1\map\geoJson.json"
KAKAO_API_KEY = "93c089f75a2730af2f15c01838e892d3"

try:
    rc('font', family='Malgun Gothic')
except:
    logging.warning("한글 폰트 설정 불가, 기본 폰트 사용")

# ==== 유틸 함수 ====
def clean_cols(cols: pd.Index) -> pd.Index:
    return (cols.str.replace('\ufeff', '', regex=False)
                .str.replace(r'\s+', ' ', regex=True)
                .str.strip())

PROVINCE_MAP = {
    "전북특별자치도": "전라북도", "전북": "전라북도",
    "전남": "전라남도", "경북": "경상북도", "경남": "경상남도",
    "충북": "충청북도", "충남": "충청남도",
    "서울시": "서울특별시", "부산시": "부산광역시", "대구시": "대구광역시",
    "인천시": "인천광역시", "광주시": "광주광역시", "대전시": "대전광역시",
    "울산시": "울산광역시", "세종시": "세종특별자치시",
    "제주도": "제주특별자치도", "강원특별자치도": "강원도",
}

GEOJSON_TO_KOREAN = {
    "Seoul": "서울", "Busan": "부산", "Daegu": "대구",
    "Incheon": "인천", "Gwangju": "광주", "Daejeon": "대전",
    "Ulsan": "울산", "Sejong": "세종",
    "Gyeonggi": "경기", "Gangwon": "강원",
    "North Chungcheong": "충청북도", "South Chungcheong": "충청남도",
    "North Jeolla": "전라북도", "South Jeolla": "전라남도",
    "North Gyeongsang": "경상북도", "South Gyeongsang": "경상남도",
    "Jeju": "제주"
}

def normalize_region(s: str) -> str:
    if pd.isna(s): return ""
    s = re.sub(r"\s+", "", str(s).strip())
    return PROVINCE_MAP.get(s, s)

def normalize_subregion(s: str) -> str:
    if pd.isna(s): return ""
    return re.sub(r"\s+", " ", str(s).strip())

BAD_LABELS = {"", "nan", "None", "알수없음"}
def valid_region(x) -> bool:
    if x is None: return False
    s = str(x).strip()
    return s not in BAD_LABELS

def _hsv_hex(h, s=0.85, v=0.9):
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return '#%02x%02x%02x' % (int(r*255), int(g*255), int(b*255))

# ==== 데이터 로드 ====
logging.info("엑셀 파일 불러오는 중...")
df = pd.read_excel(FILE_PATH)
df.columns = clean_cols(df.columns)
region_col, subregion_col = '광역지역', '세부지역'
df['설비용량'] = pd.to_numeric(df.get('설비용량', 0), errors='coerce').fillna(0)
df['광역지역_norm'] = df[region_col].apply(normalize_region)
df['세부지역_norm']  = df[subregion_col].apply(normalize_subregion)
df['주소'] = df['세부지역_norm']
df = df[df['광역지역_norm'].apply(valid_region)].copy()
logging.info(f"데이터 로드 완료: {len(df)}건")

# ==== 색상 팔레트 ====
unique_regions = sorted({r for r in df['광역지역_norm'] if valid_region(r)})
palette = [_hsv_hex(i / max(1, len(unique_regions))) for i in range(len(unique_regions))]
REGION_COLORS = dict(zip(unique_regions, palette))
def pick_region_color(region): return REGION_COLORS.get(region, "#7f7f7f")

# ==== 좌표 처리 ====
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        coords_cache = json.load(f)
else:
    coords_cache = {}

def get_coords_kakao(address: str):
    if address in coords_cache: return address, coords_cache[address]
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    params = {"query": address}
    try:
        r = requests.get(url, headers=headers, params=params, timeout=5)
        r.raise_for_status()
        data = r.json()
        if data.get('documents'):
            x=float(data['documents'][0]['x']); y=float(data['documents'][0]['y'])
            coords_cache[address] = [y, x]
        else:
            coords_cache[address] = [None,None]
    except:
        coords_cache[address] = [None,None]
    return address, coords_cache[address]

targets = [a for a in df['주소'].dropna().unique() if a not in coords_cache]
if targets:
    logging.info(f"카카오 API로 {len(targets)}개 주소 좌표 변환 중...")
    with ThreadPoolExecutor(max_workers=8) as ex:
        futures = [ex.submit(get_coords_kakao, addr) for addr in targets]
        for _ in tqdm(as_completed(futures), total=len(futures), desc="좌표 변환"):
            _; time.sleep(0.05)
    with open(CACHE_FILE,"w",encoding="utf-8") as f:
        json.dump(coords_cache,f,ensure_ascii=False,indent=2)

df['coords'] = df['주소'].map(coords_cache)
df[['위도','경도']] = pd.DataFrame(df['coords'].tolist(), index=df.index)
df = df.dropna(subset=['위도','경도'])
logging.info("좌표 변환 완료")

# ==== 요약 데이터 ====
grouped_sub = df.groupby(['위도','경도'], as_index=False).agg(
    발전소수=('발전기명','count'),
    총설비용량=('설비용량','sum'),
    대표광역=('광역지역_norm', lambda x: x.value_counts().idxmax()),
    세부지역=('세부지역_norm', lambda x: x.value_counts().idxmax())
)

grouped_region = df.groupby('광역지역_norm', as_index=False).agg(
    발전소수=('발전기명','count'),
    총설비용량=('설비용량','sum')
)
grouped_region = grouped_region.sort_values('발전소수', ascending=False)
grouped_region.insert(0, "순위", range(1, len(grouped_region)+1))
grouped_region.rename(columns={
    '광역지역_norm':'시도별',
    '발전소수':'발전소 수(개)',
    '총설비용량':'총 설비용량(MW)'
}, inplace=True)

# ==== 그래프 ====
colors = [REGION_COLORS.get(r,'#999') for r in grouped_region['시도별']]
fig, ax1 = plt.subplots(figsize=(7,3.5))
ax1.bar(grouped_region['시도별'], grouped_region['발전소 수(개)'], color=colors,label="발전소수")
ax2 = ax1.twinx()
ax2.plot(grouped_region['시도별'], grouped_region['총 설비용량(MW)'], color='black', linestyle='--', marker='o', label="총설비용량(MW)")
ax1.set_title("시도별 발전소 수(개)/ 설비 용량(MW)", fontsize=12)
ax1.tick_params(axis='x', rotation=45)
plt.tight_layout()
buf = BytesIO()
plt.savefig(buf, format="png", bbox_inches="tight")
buf.seek(0)
chart_base64 = base64.b64encode(buf.read()).decode('utf-8')
plt.close()

# ==== 지도 ====
m = folium.Map(location=[36.5,127.8], zoom_start=7, tiles=None)
folium.TileLayer(tiles='https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
                 attr='Map data © OpenStreetMap contributors', name='OSM').add_to(m)

# 광역지역 레이어
with open(GEOJSON_FILE,"r",encoding="utf-8") as f:
    geo_json = json.load(f)

region_layer = folium.FeatureGroup(name='광역지역', show=True)
max_count = grouped_region['발전소 수(개)'].max()
for feature in geo_json['features']:
    eng_name = feature['properties']['name']
    kor_name = GEOJSON_TO_KOREAN.get(eng_name, eng_name)
    row = grouped_region[grouped_region['시도별']==kor_name]
    count = row['발전소 수(개)'].values[0] if not row.empty else 0
    alpha = 0.3 + 0.7*(count/max_count) if max_count>0 else 0.3
    base_color = REGION_COLORS.get(kor_name,"#7f7f7f")
    tooltip_html = f"<b>{kor_name}</b><br>발전소 수: {count:,}개<br>총 설비용량: {row['총 설비용량(MW)'].values[0]:,.2f} MW" if not row.empty else ""
    
    folium.GeoJson(
        feature,
        style_function=lambda f, col=base_color, a=alpha: {
            'fillColor': col, 'fillOpacity': a, 'color':'black','weight':1
        },
        tooltip=folium.Tooltip(tooltip_html),
        highlight_function=lambda x:{'weight':3,'color':'orange','fillOpacity':0.5},
        name='광역지역'
    ).add_to(region_layer)
region_layer.add_to(m)

# 세부지역 마커
sub_layer = folium.FeatureGroup(name='세부지역', show=True)
for _, r in grouped_sub.iterrows():
    clr = pick_region_color(r['대표광역'])
    folium.CircleMarker(
        location=[r['위도'], r['경도']],
        radius=(r['발전소수'] ** 0.2),
        color=clr,
        fill=True, fill_color=clr,
        fill_opacity=0.85,
        popup=folium.Popup(
    f"<b>{r['세부지역']}</b><br>발전소 수: {r['발전소수']}<br>총 설비용량: {r['총설비용량']:.2f} MW",
    max_width=120, min_width=60
    )).add_to(sub_layer)
sub_layer.add_to(m)

# ==== 범례 HTML + JS ====
legend_html = f"""
<div id='legend_sub' style='position: fixed; bottom: 20px; left: 20px;
background:white; border:2px solid grey; padding:6px; z-index:9999;
display:flex; flex-direction:column; gap:4px; border-radius:6px;'>
<b style='text-align:center;'>시도별</b>
{"".join([
f"<div style='display:flex; align-items:center; gap:4px;'>"
f"<div style='width:18px; height:18px; background:{REGION_COLORS[r]}; border:1px solid #000;'></div>"
f"<span>{r}</span></div>"
for r in unique_regions
])}
</div>
<script>
function updateLegend(){{;
    var subLayer = Object.values(map._layers).filter(l=>l.options && l.options.name=='세부지역');
    document.getElementById('legend_sub').style.display = (subLayer.length>0 && map.hasLayer(subLayer[0])) ? 'flex':'none';
}}
map.on('overlayadd overlayremove', updateLegend);
updateLegend();
</script>
"""
m.get_root().html.add_child(folium.Element(legend_html))

LayerControl(collapsed=False, position='topright').add_to(m)

# ==== HTML + 대시보드 ====
table_html = grouped_region.to_html(index=False, justify='center', border=0,
                                    classes='data-table', float_format='{:,.2f}'.format)
final_html = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>태양광 발전소 지도 대시보드</title>
<style>
body {{ display:flex; flex-wrap:wrap; margin:0; background:#f9f9f9; font-family:'Malgun Gothic','Segoe UI',sans-serif; }}
#left-panel {{ flex:1 1 40%; min-width:400px; margin:10px; }}
#chart {{ text-align:center; margin-bottom:20px; }}
#chart img {{ width:95%; border-radius:10px; box-shadow:0 0 10px rgba(0,0,0,0.2); }}
#table {{ width:95%; margin:auto; text-align:center; }}
.data-table {{ border-collapse:collapse; width:100%; box-shadow:0 0 10px rgba(0,0,0,0.1); }}
.data-table th {{ background-color:#4CAF50; color:white; padding:8px; }}
.data-table td {{ border:1px solid #ddd; padding:8px; }}
.data-table tr:nth-child(even) {{ background-color:#f2f2f2; }}
.data-table tr:hover {{ background-color:#ddd; }}
#map {{ flex:1 1 55%; min-width:500px; height:calc(100vh - 20px); margin:10px; border-radius:12px; box-shadow:0 0 10px rgba(0,0,0,0.3); }}
</style>
</head>
<body>
<div id="left-panel">
  <div id="chart">
    <h2><시도별 발전소 수 및 설비용량></h2>
    <img src="data:image/png;base64,{chart_base64}" alt="chart">
  </div>
  <div id="table">
    <h2><시도별 요약표></h2>
    {table_html}
  </div>
</div>
<div id="map">{m._repr_html_()}</div>
</body>
</html>
"""

with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(final_html)

logging.info(f"✅ 대시보드 생성 완료: {OUTPUT_HTML}")
webbrowser.open(OUTPUT_HTML)
