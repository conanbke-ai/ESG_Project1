import os, json, time, re, colorsys, logging, base64, webbrowser, html
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import requests
from tqdm import tqdm
import folium
from folium import Element, FeatureGroup, LayerControl
import matplotlib.pyplot as plt
from matplotlib import rc
from io import BytesIO

# ===== 로깅 =====
logging.basicConfig(level=logging.INFO, format='[%(asctime)s]✅ %(message)s')

# ===== 경로 / API =====
FILE_PATH   = r"C:\ESG_Project1\file\generator_file\HOME_발전설비_발전기별.xlsx"
CACHE_FILE  = r"C:\ESG_Project1\map\json\coord_cache.json"
GEOJSON_FILE = r"C:\ESG_Project1\map\json\geoJson.json"
OUTPUT_HTML = r"C:\ESG_Project1\map\html\solar_dashboard.html"
KAKAO_API_KEY = "93c089f75a2730af2f15c01838e892d3"

# ===== 폰트 =====
try:
    rc('font', family='Malgun Gothic')
except:
    logging.warning("한글 폰트 설정 불가, 기본 폰트 사용")

# ==== 유틸 함수 ====
def clean_cols(cols: pd.Index) -> pd.Index:
    return (cols.str.replace('\ufeff', '', regex=False)
                .str.replace(r'\s+', ' ', regex=True)
                .str.strip())

def _hsv_hex(h, s=0.85, v=0.9):
    r, g, b = colorsys.hsv_to_rgb(h, s, v)
    return '#%02x%02x%02x' % (int(r*255), int(g*255), int(b*255))

# ---- 표준 라벨 & 패턴 ----

CANON = [
    "서울특별시","부산광역시","대구광역시","인천광역시","광주광역시","대전광역시","울산광역시","세종특별자치시",
    "경기도","강원특별자치도","충청북도","충청남도","전라북도","전라남도","경상북도","경상남도","제주특별자치도"
]
REGION_PATTERNS = {
    "서울특별시":        r"서울(특별)?\s*시?",
    "부산광역시":        r"부산(광역)?\s*시?",
    "대구광역시":        r"대구(광역)?\s*시?",
    "인천광역시":        r"인천(광역)?\s*시?",
    "광주광역시":        r"광주(광역)?\s*시?",
    "대전광역시":        r"대전(광역)?\s*시?",
    "울산광역시":        r"울산(광역)?\s*시?",
    "세종특별자치시":    r"세종(특별자치)?\s*시?",
    "경기도":            r"경기\s*(도)?",
    "강원특별자치도":    r"강원(특별자치)?\s*도?",
    "충청북도":          r"(충북|충청북도)",
    "충청남도":          r"(충남|충청남도)",
    "전라북도":          r"(전북특별자치도|전북|전라북도)",
    "전라남도":          r"(전남|전라남도)",
    "경상북도":          r"(경북|경상북도)",
    "경상남도":          r"(경남|경상남도)",
    "제주특별자치도":    r"(제주(특별자치)?\s*도?)",
}

DISPLAY_NAME = {
    "전라북도": "전북특별자치도",
    "세종특별자치시": "세종특별자치시",
    "강원특별자치도": "강원특별자치도",
    "제주특별자치도": "제주특별자치도",
}

GEOJSON_TO_KOREAN = {
    "Seoul": "서울특별시", "Busan": "부산광역시", "Daegu": "대구광역시",
    "Incheon": "인천광역시", "Gwangju": "광주광역시", "Daejeon": "대전광역시",
    "Ulsan": "울산광역시", "Sejong": "세종특별자치시",
    "Gyeonggi": "경기도", "Gangwon": "강원특별자치도",
    "North Chungcheong": "충청북도", "South Chungcheong": "충청남도",
    "North Jeolla": "전라북도", "South Jeolla": "전라남도",
    "North Gyeongsang": "경상북도", "South Gyeongsang": "경상남도",
    "Jeju": "제주특별자치도"
}

compiled = {k: re.compile(v) for k, v in REGION_PATTERNS.items()}

def to_canonical(s: str) -> str:
    if pd.isna(s): return ""
    t = re.sub(r"\s+", "", str(s))
    for canon, pat in compiled.items():
        if pat.search(t):
            return canon
    return str(s).strip()

def display_region_name(canon: str) -> str:
    return DISPLAY_NAME.get(canon, canon)

def normalize_subregion(s: str) -> str:
    if pd.isna(s): return ""
    return re.sub(r"\s+", " ", str(s).strip())

# === 세부지역 접두어 제거 ===
_PREFIX_CANDIDATES = set(CANON) | {"전북특별자치도"}
_REGION_PREFIX_ANY = re.compile(
    r"^\s*(?:%s)\s*" % "|".join(map(re.escape, sorted(_PREFIX_CANDIDATES, key=len, reverse=True)))
)
def strip_region_prefix_any(subregion: str) -> str:
    if not isinstance(subregion, str):
        return ""
    s = subregion.strip()
    while True:
        new = _REGION_PREFIX_ANY.sub("", s, count=1)
        if new == s: break
        s = new.strip()
    return s

BAD_LABELS = {"", "nan", "None", "알수없음"}
def valid_region(x) -> bool:
    if x is None: return False
    s = str(x).strip()
    return s not in BAD_LABELS

# ==== 데이터 로드 ====
logging.info("엑셀 파일 불러오는 중...")
df = pd.read_excel(FILE_PATH)
df.columns = clean_cols(df.columns)

region_col, subregion_col = '광역지역', '세부지역'
df['설비용량'] = pd.to_numeric(df.get('설비용량', 0), errors='coerce').fillna(0)
df['광역지역_std'] = df[region_col].map(to_canonical)
df['세부지역_std']  = df[subregion_col].astype(str).str.strip()
df['주소'] = df['세부지역_std']
df = df[df['광역지역_std'] != ""].copy()
logging.info(f"데이터 로드 완료: {len(df)}건")

# ==== 좌표 캐시 ====
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
    대표광역=('광역지역_std', lambda x: x.value_counts().idxmax()),
    세부지역=('세부지역_std', lambda x: x.value_counts().idxmax())
)

grouped_region = df.groupby('광역지역_std', as_index=False).agg(
    발전소수=('발전기명','count'),
    총설비용량=('설비용량','sum')
)

# ==== 색상 팔레트 ====
unique_regions = sorted(grouped_region['광역지역_std'].unique().tolist(),
                        key=lambda x: CANON.index(x) if x in CANON else 999)
palette = [_hsv_hex(i / max(1, len(unique_regions))) for i in range(len(unique_regions))]
REGION_COLORS = dict(zip(unique_regions, palette))
def pick_region_color(region): return REGION_COLORS.get(region, "#7f7f7f")

# ===== 표 =====
table_df = grouped_region.rename(
    columns={'광역지역_std': '시도별', '발전소수': '발전소 수(개)', '총 설비용량': '설비용량(MW)'}
).copy()
table_df['시도별'] = table_df['시도별'].map(display_region_name)
table_df = table_df.sort_values('발전소 수(개)', ascending=False).reset_index(drop=True)
table_df.insert(0, '순위', range(1, len(table_df) + 1))
table_html = table_df.to_html(index=False, justify='center', border=0,
                              classes='data-table', float_format='{:,.2f}'.format)
logging.info("표 데이터 처리 완료")

# ==== 그래프 ====
region_stats = grouped_region.copy()
region_stats['표시광역'] = region_stats['광역지역_std'].map(display_region_name)
region_stats = region_stats.sort_values('발전소수', ascending=False).reset_index(drop=True)

x = range(len(region_stats))
colors = [pick_region_color(r) for r in region_stats['광역지역_std']]

fig, ax1 = plt.subplots(figsize=(8.8, 4.4))
plt.subplots_adjust(bottom=0.30, left=0.10, right=0.92, top=0.88)

bars = ax1.bar(x, region_stats['발전소수'], color=colors, width=0.68, label='발전소 수(개)')
ax2 = ax1.twinx()
line, = ax2.plot(x, region_stats['총설비용량'], color='black', linestyle='--', marker='o', label='설비용량(MW)')

ax1.set_xticks(x)
ax1.set_xticklabels(region_stats['표시광역'], rotation=45, ha='right')
ax1.margins(x=0.04); ax2.margins(x=0.04)

ax1.legend([bars, line], ['발전소 수(개)', '설비용량(MW)'], loc='upper right', fontsize=9, frameon=True)

plt.tight_layout()
buf = BytesIO(); plt.savefig(buf, format="png", bbox_inches="tight"); buf.seek(0)
chart_base64 = base64.b64encode(buf.read()).decode('utf-8')
plt.close()
logging.info("그래프 데이터 처리 완료")

# ==== 지도 ====
m = folium.Map(location=[36.5,127.8], zoom_start=7, tiles=None)
folium.TileLayer(tiles='https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
                 attr='Map data © OpenStreetMap contributors', name='OSM').add_to(m)
logging.info("지도 생성 완료")

# 광역지역 레이어
with open(GEOJSON_FILE,"r",encoding="utf-8") as f:
    geo_json = json.load(f)

region_layer = folium.FeatureGroup(name='광역지역', show=True)
max_count = grouped_region['발전소수'].max()
for feature in geo_json['features']:
    eng_name = feature['properties']['name']
    kor_name = GEOJSON_TO_KOREAN.get(eng_name, eng_name)
    row = grouped_region[grouped_region['광역지역_std']==kor_name]
    count = row['발전소수'].values[0] if not row.empty else 0
    alpha = 0.3 + 0.7*(count/max_count) if max_count>0 else 0.3
    base_color = REGION_COLORS.get(kor_name,"#7f7f7f")
    tooltip_html = f"<b>{kor_name}</b><br>발전소 수: {count:,}개<br>총 설비용량: {row['총설비용량'].values[0]:,.2f} MW" if not row.empty else ""
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
logging.info("광역지역 레이어 완료")

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
    f"<b>{r['세부지역']}</b><br>발전소 수: {r['발전소수']}개<br>총 설비용량: {r['총설비용량']:.2f} MW",
    max_width=120, min_width=60
    )).add_to(sub_layer)
sub_layer.add_to(m)
logging.info("세부지역 마커 완료")

# ==== 범례 HTML + JS ====
legend_items = ''.join(
    f'''
    <div style="display:flex;align-items:center;margin:4px 0;">
      <span style="display:inline-block;width:14px;height:14px;border-radius:50%;
                   background:{REGION_COLORS[name]};border:1px solid #333;margin-right:6px;"></span>
      <span>{html.escape(display_region_name(name))}</span>
    </div>
    '''
    for name in sorted(REGION_COLORS.keys(), key=lambda x: CANON.index(x) if x in CANON else 999)
)
legend_html = f'''
<div style="
  position: fixed; left: 16px; bottom: 16px; z-index: 9999;
  background: rgba(255,255,255,0.95);
  border: 1px solid #e5e7eb; border-radius: 8px;
  box-shadow: 0 6px 18px rgba(0,0,0,0.08);
  padding: 8px 10px; font-size: 13px; line-height: 1.2;
  max-height: 260px; width: 160px; overflow: auto;
">
  <div style="font-weight:700; margin-bottom:6px;">지역 색상</div>
  {legend_items}
</div>
'''
m.get_root().html.add_child(Element(legend_html))
LayerControl(collapsed=False).add_to(m)

# ==== HTML + 대시보드 ====
final_html = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>태양광 발전소 지도 대시보드</title>
<style>
html, body {{ height:100%; }}
  body {{
    min-height:100vh;
    display:flex;
    flex-wrap:nowrap;      /* 한 줄 유지 */
    align-items:stretch;   /* 양 칼럼 같은 높이 */
    margin:0;
    background:#f9f9f9;
    font-family:'Malgun Gothic','Segoe UI',sans-serif;
  }}
#left-panel {{ flex:1 1 40%; min-width:400px; margin:10px; }}
#chart {{ text-align:center; margin-bottom:20px; }}
#chart img {{ width:95%; border-radius:10px; box-shadow:0 0 10px rgba(0,0,0,0.2); }}
#table {{ width:95%; margin:auto; text-align:center; }}
.data-table {{ border-collapse:collapse; width:100%; box-shadow:0 0 10px rgba(0,0,0,0.1); }}
.data-table th {{ background-color:#4CAF50; color:white; padding:8px; }}
.data-table td {{ border:1px solid #ddd; padding:8px; }}
.data-table tr:nth-child(even) {{ background-color:#f2f2f2; }}
.data-table tr:hover {{ background-color:#ddd; }}
/* 지도 컨테이너: 높이는 JS가 동기화 */
  #map {{
    flex:1 1 64%;
    min-width:520px;
    height:auto;
    margin:0;
    border-radius:12px; box-shadow:0 0 10px rgba(0,0,0,0.3);
    position:relative;
  }}
  /* folium 내부 컨테이너를 꽉 채우기 위한 기본값 */
  #map > div              {{ height:100% !important; }}
  #map .folium-map        {{ height:100% !important; padding-bottom:0 !important; }}
  #map .folium-map > div  {{ height:100% !important; }}
  #map .leaflet-container {{ height:100% !important; }}
  #map iframe             {{ height:100% !important; }}
</style>
</head>
<body>
<div id="left-panel">
  <div id="chart">
    <h2>🅰️ 시도별 발전소 수 및 설비용량</h2>
    <img src="data:image/png;base64,{chart_base64}" alt="chart">
  </div>
  <div id="table">
    <h2>🅱️ 시도별 요약표</h2>
    {table_html}
  </div>
</div>
<div id="map">{m._repr_html_()}</div>
<script>
(function() {{
  function setMapHeight() {{
    var left = document.getElementById('left-panel');
    var map  = document.getElementById('map');
    if (!left || !map) return;

    // 왼쪽 패널과 뷰포트 중 더 큰 값을 사용
    var want = Math.max(window.innerHeight, left.scrollHeight, left.getBoundingClientRect().height);

    // 컨테이너 및 내부 요소 모두 동일 높이로 강제
    map.style.height = want + 'px';
    var targets = map.querySelectorAll('.leaflet-container, .folium-map, #map > div, #map iframe');
    targets.forEach(function(el) {{
      el.style.height = want + 'px';
      el.style.minHeight = want + 'px';
      el.style.maxHeight = want + 'px';
    }});

    // Leaflet 사이즈 재계산
    setTimeout(function() {{
      window.dispatchEvent(new Event('resize'));
      if (window.L && typeof L !== 'undefined') {{
        // folium이 생성한 맵 div 찾아 invalidateSize 시도
        var mapDiv = map.querySelector('.leaflet-container');
        if (mapDiv && mapDiv._leaflet_id && window._leaflet_map) {{
          try {{ window._leaflet_map.invalidateSize(); }} catch(e) {{}}
        }}
      }}
    }}, 50);
  }}

  // 초기/리사이즈
  window.addEventListener('load', setMapHeight);
  window.addEventListener('resize', setMapHeight);

  // 내용 변동 감지(표 정렬 등으로 높이 변할 때)
  var mo = new MutationObserver(setMapHeight);
  mo.observe(document.getElementById('left-panel'), {{ subtree:true, childList:true, attributes:true }});
}})();
</script>
</body>
</html>
"""

# HTML 저장 및 열기
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(final_html)

logging.info(f" 대시보드 생성 완료: {OUTPUT_HTML}")
webbrowser.open(OUTPUT_HTML)
