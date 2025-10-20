import os, json, time, html, webbrowser, re, colorsys, logging
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import requests
from tqdm import tqdm
import folium
from folium import Element, FeatureGroup, LayerControl
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from matplotlib import font_manager, rc

# ===== 로깅 =====
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] ✅ %(message)s')

# ===== 한글 폰트 =====
font_path = "C:/Windows/Fonts/malgun.ttf"
if os.path.exists(font_path):
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font_name)
else:
    logging.warning("한글 폰트를 찾을 수 없습니다. 그래프가 깨질 수 있습니다.")

# ===== 경로 / API =====
FILE_PATH   = r"C:\ESG_Project1\file\generator_file\HOME_발전설비_발전기별.xlsx"
CACHE_FILE  = r"C:\ESG_Project1\map\coord_cache.json"
OUTPUT_HTML = r"C:\ESG_Project1\map\solar_dashboard.html"
KAKAO_API_KEY = "93c089f75a2730af2f15c01838e892d3"

# ===== 유틸 =====
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

# ===== 데이터 로드 =====
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

# ===== 색상 팔레트 =====
unique_regions = sorted({r for r in df['광역지역_norm'] if valid_region(r)})
palette = [_hsv_hex(i / max(1, len(unique_regions))) for i in range(len(unique_regions))]
REGION_COLORS = dict(zip(unique_regions, palette))
def pick_region_color(region):
    return REGION_COLORS.get(region, "#7f7f7f")

# ===== 좌표 캐시 =====
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        coords_cache = json.load(f)
else:
    coords_cache = {}

def get_coords_kakao(address: str):
    if address in coords_cache:
        return address, coords_cache[address]
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    params = {"query": address}
    try:
        r = requests.get(url, headers=headers, params=params, timeout=5)
        r.raise_for_status()
        data = r.json()
        if data.get('documents'):
            x = float(data['documents'][0]['x']); y = float(data['documents'][0]['y'])
            coords_cache[address] = [y, x]
        else:
            coords_cache[address] = [None, None]
    except Exception:
        coords_cache[address] = [None, None]
    return address, coords_cache[address]

targets = [a for a in df['주소'].dropna().unique() if a not in coords_cache]
if targets:
    logging.info(f"카카오 API로 {len(targets)}개 주소 좌표 변환 중...")
    with ThreadPoolExecutor(max_workers=8) as ex:
        futures = [ex.submit(get_coords_kakao, addr) for addr in targets]
        for _ in tqdm(as_completed(futures), total=len(futures), desc="좌표 변환"):
            _; time.sleep(0.05)
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(coords_cache, f, ensure_ascii=False, indent=2)

df['coords'] = df['주소'].map(coords_cache)
df[['위도','경도']] = pd.DataFrame(df['coords'].tolist(), index=df.index)
df = df.dropna(subset=['위도','경도'])
logging.info("좌표 변환 완료.")

# ===== 요약 (세부지역 단위) =====
grouped = (
    df.groupby(['위도','경도'], as_index=False)
      .agg(
          발전소수=('발전기명','count'),
          총설비용량=('설비용량','sum'),
          대표광역=('광역지역_norm', lambda x: x.value_counts().idxmax()),
          세부지역=('세부지역_norm', lambda x: x.value_counts().idxmax()),
      )
)

# ===== 지도 생성 =====
m = folium.Map(location=[36.5, 127.8], zoom_start=7, tiles="CartoDB positron")
region_layer = FeatureGroup(name="광역지역", show=True).add_to(m)
subregion_layer = FeatureGroup(name="세부지역", show=False).add_to(m)
bounds = []

for _, r in grouped.iterrows():
    lat, lon = float(r['위도']), float(r['경도'])
    color = pick_region_color(r['대표광역'])
    radius = (r['발전소수']**0.05)
    
    # 세부지역 레이어
    folium.CircleMarker(
        location=[lat, lon],
        radius=radius,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.85,
        popup=folium.Popup(f"<b>{r['세부지역']}</b><br>발전소 수: {r['발전소수']}<br>총 설비용량: {r['총설비용량']:.2f} MW", max_width=320),
        tooltip=r['세부지역']
    ).add_to(subregion_layer)

    # 광역지역 레이어
    folium.CircleMarker(
        location=[lat, lon],
        radius=radius,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.6,
        popup=folium.Popup(f"<b>{r['대표광역']}</b>", max_width=200),
        tooltip=r['대표광역']
    ).add_to(region_layer)
    
    bounds.append([lat, lon])

if bounds:
    m.fit_bounds(bounds)

# 범례
legend_items = ''.join(
    f'<div><span style="display:inline-block;width:14px;height:14px;background:{color};border-radius:50%;margin-right:6px;border:1px solid #333;"></span>{name}</div>'
    for name, color in sorted(REGION_COLORS.items())
)
legend_html = f'''
<div style="position: fixed; left: 20px; bottom: 20px; z-index: 9999;
 background: rgba(255,255,255,0.95); padding: 8px 10px; border: 1px solid #888;
 border-radius: 6px; font-size: 12px; max-height: 260px; overflow:auto;">
  <b>지역 색상</b>
  <div style="margin-top:6px;">{legend_items}</div>
</div>
'''
m.get_root().html.add_child(Element(legend_html))

# 토글 버튼
LayerControl(collapsed=False).add_to(m)

# ===== 그래프 생성 (Matplotlib) =====
region_stats = (
    df.groupby('광역지역_norm', as_index=False)
      .agg(발전소수=('발전기명','count'), 총설비용량=('설비용량','sum'))
      .sort_values('발전소수', ascending=False)
)

fig, ax1 = plt.subplots(figsize=(8,4))
bar_colors = [pick_region_color(r) for r in region_stats['광역지역_norm']]
ax1.bar(region_stats['광역지역_norm'], region_stats['발전소수'], color=bar_colors, label="발전소 수")
ax2 = ax1.twinx()
ax2.plot(region_stats['광역지역_norm'], region_stats['총설비용량'],
         color='black', linestyle='--', marker='o', label="총 설비용량(MW)")
ax1.set_xlabel("광역지역")
ax1.set_ylabel("발전소 수")
ax2.set_ylabel("총 설비용량(MW)")
ax1.set_title("광역지역별 발전소 수 및 설비용량")
ax1.tick_params(axis='x', rotation=45)
ax1.legend(loc="upper left")
ax2.legend(loc="upper right")
plt.tight_layout()

buf = BytesIO()
plt.savefig(buf, format="png", bbox_inches="tight")
buf.seek(0)
chart_base64 = base64.b64encode(buf.read()).decode('utf-8')
plt.close()

# ===== HTML 결합 =====
final_html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>태양광 발전소 대시보드</title>
<style>
body {{
  display: flex;
  flex-wrap: wrap;
  margin: 0;
  font-family: 'Segoe UI', sans-serif;
  background: #f9f9f9;
}}
#chart {{
  flex: 1 1 35%;
  text-align: center;
  margin: 10px;
  min-width: 380px;
}}
#chart img {{
  width: 95%;
  border-radius: 10px;
  box-shadow: 0 0 10px rgba(0,0,0,0.2);
}}
#map {{
  flex: 1 1 60%;
  height: 95vh;
  margin: 10px;
  min-width: 500px;
  border-radius: 12px;
  box-shadow: 0 0 10px rgba(0,0,0,0.3);
}}
@media (max-width: 900px) {{
  body {{ flex-direction: column; align-items: center; }}
  #map, #chart {{ width: 90%; height: auto; }}
}}
</style>
</head>
<body>
<div id="chart">
  <h2>광역지역별 발전소 수 및 설비용량</h2>
  <img src="data:image/png;base64,{chart_base64}" alt="chart">
</div>
<div id="map">{m._repr_html_()}</div>
</body>
</html>
"""

with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(final_html)

logging.info(f"대시보드 생성 완료: {OUTPUT_HTML}")
webbrowser.open(OUTPUT_HTML)
