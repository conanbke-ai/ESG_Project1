import os
import json
import time
import requests
import pandas as pd
import folium
from folium.features import GeoJsonTooltip
from branca.colormap import linear
import webbrowser
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# ================== 경로 ==================
EXCEL_FILE = r"C:\ESG_Project1\file\generator_file\HOME_발전설비_발전기별.xlsx"
CACHE_FILE = r"C:\ESG_Project1\map\coord_cache.json"
GEOJSON_FILE = r"C:\ESG_Project1\map\geoJson.json"
OUTPUT_HTML = r"C:\ESG_Project1\map\generator_map.html"

# ================== Kakao API ==================
KAKAO_API_KEY = "YOUR_KAKAO_API_KEY"
KAKAO_URL = "https://dapi.kakao.com/v2/local/search/address.json"

# ================== 데이터 로드 ==================
df = pd.read_excel(EXCEL_FILE)
required_cols = ['광역지역','세부지역','설비용량']
for col in required_cols:
    if col not in df.columns:
        raise ValueError(f"엑셀에 '{col}' 컬럼 필요")

# ================== 좌표 캐시 ==================
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE,"r",encoding="utf-8") as f:
        coord_cache = json.load(f)
else:
    coord_cache = {}

# ================== Kakao 좌표 함수 ==================
def get_coord(address):
    if address in coord_cache:
        return coord_cache[address]
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    try:
        r = requests.get(KAKAO_URL, headers=headers, params={"query":address}, timeout=5)
        r.raise_for_status()
        docs = r.json().get("documents", [])
        if docs:
            coord = [float(docs[0]['y']), float(docs[0]['x'])]
            coord_cache[address] = coord
            return coord
    except Exception as e:
        logging.warning(f"좌표 조회 실패: {address} → {e}")
    coord_cache[address] = [36.5, 127.5]
    return [36.5,127.5]

# ================== 세부지역 좌표 생성 ==================
for _, row in df.iterrows():
    loc = f"{row['광역지역']} {row['세부지역']}"
    if loc not in coord_cache:
        get_coord(loc)
        time.sleep(0.2)

with open(CACHE_FILE,"w",encoding="utf-8") as f:
    json.dump(coord_cache,f,ensure_ascii=False, indent=2)

# ================== 광역지역 요약 ==================
region_summary = df.groupby('광역지역').agg(
    총발전소수=('세부지역','count'),
    총설비용량=('설비용량','sum')
).reset_index()

# ================== Folium 지도 생성 ==================
m = folium.Map(location=[36.5,127.8], zoom_start=7)

# ================== GeoJSON 로드 ==================
with open(GEOJSON_FILE, "r", encoding="utf-8") as f:
    geojson_data = json.load(f)

# ================== 광역지역 색상 ==================
min_val, max_val = region_summary['총발전소수'].min(), region_summary['총발전소수'].max()
colormap = linear.YlOrRd_09.scale(min_val, max_val)
colormap.caption = '광역지역별 총 발전소 수'
colormap.add_to(m)

def style_function(feature):
    region = feature['properties']['name']
    row = region_summary[region_summary['광역지역']==region]
    if not row.empty:
        val = row['총발전소수'].values[0]
        return {'fillColor': colormap(val), 'color':'black','weight':1,'fillOpacity':0.7}
    return {'fillColor':'#dddddd','color':'black','weight':0.5,'fillOpacity':0.4}

geojson_layer = folium.GeoJson(
    geojson_data,
    style_function=style_function,
    tooltip=GeoJsonTooltip(fields=['name'], aliases=['광역지역']),
    highlight_function=lambda x: {'weight':3,'color':'orange','fillOpacity':0.2}
).add_to(m)

# ================== Plotly 버블 JS ==================
grouped_json = df.to_dict(orient='records')

bubble_js = f"""
<div id="right-graph" style="position:fixed; top:10px; right:10px; width:480px; height:500px;
background:white; z-index:9999; padding:10px; border:2px solid black; overflow:auto;"></div>
<script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
<script>
var df = {json.dumps(grouped_json, ensure_ascii=False)};

function drawGraph(region){{
    var filtered = df.filter(r=>r['광역지역']===region);
    if(filtered.length===0){{
        document.getElementById('right-graph').innerHTML="<p>데이터 없음</p>";
        return;
    }}
    var x = filtered.map(r=>r['세부지역']);
    var y = filtered.map(r=>r['설비용량']);
    var trace={{x:x,y:y,text:x.map((loc,i)=>loc+"<br>설비용량:"+y[i]+" MW"),
                mode:"markers", marker:{{size:y.map(v=>Math.sqrt(v)*3),color:"steelblue",sizemode:"area",sizemin:5}}}};
    var layout={{title:region+" 세부지역 발전소 현황",
                xaxis:{{title:"세부지역",tickangle:-45}},
                yaxis:{{title:"설비용량 (MW)"}}, margin:{{l:40,r:10,t:40,b:100}}, hovermode:"closest"}};
    Plotly.newPlot("right-graph",[trace],layout,{{responsive:true}});
}}

// 클릭 이벤트 부착
setTimeout(function(){{
    var layers = document.getElementsByClassName('leaflet-interactive');
    for(var i=0;i<layers.length;i++){{
        layers[i].addEventListener('click', function(){{
            drawGraph(this.__data__.properties.name);
        }});
    }}
}}, 1000);
</script>
"""

m.get_root().html.add_child(folium.Element(bubble_js))
folium.LayerControl().add_to(m)

# ================== 저장 ==================
m.save(OUTPUT_HTML)
webbrowser.open(OUTPUT_HTML)
print("✅ 지도 + 광역지역 색상 + Plotly 버블 표시 완료")
