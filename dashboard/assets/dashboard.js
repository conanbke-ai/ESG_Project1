(async function () {
  const app = document.getElementById("app");
  const view = document.body.dataset.view || "coverage";
  const format = new Intl.NumberFormat("ko-KR");
  const pct = (value) => value == null ? "-" : `${(Number(value) * 100).toFixed(1)}%`;
  const number = (value, digits = 0) => value == null ? "-" : Number(value).toLocaleString("ko-KR", { maximumFractionDigits: digits });
  const badge = (text, kind) => `<span class="badge ${kind}">${esc(text)}</span>`;

  try {
    const response = await fetch("data/dashboard_data.json", { cache: "no-store" });
    if (!response.ok) throw new Error(`dashboard_data.json HTTP ${response.status}`);
    const data = await response.json();
    let boundaries = null;
    if (view === "coverage") {
      const boundaryResponse = await fetch("data/korea_provinces.geojson", { cache: "no-store" });
      if (boundaryResponse.ok) boundaries = await boundaryResponse.json();
    }
    app.innerHTML = view === "quality" ? qualityView(data) : coverageView(data);
    if (view === "quality") drawTrainingMap(data.plants);
    else drawNationalMap(data.national_inventory, boundaries);
  } catch (error) {
    app.innerHTML = `<p class="error">대시보드 데이터를 읽지 못했습니다. 프로젝트 루트에서 <code>python app.py build-dashboard</code>를 실행한 뒤 <code>python app.py serve-dashboard</code>로 열어주세요. (${esc(error.message)})</p>`;
  }

  function esc(value) {
    return String(value ?? "-").replace(/[&<>'"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[ch]));
  }

  function hero(data, title, description, scope) {
    return `<section class="hero"><div><p class="eyebrow">National solar intelligence</p><h1>${esc(title)}</h1><p class="lede">${esc(description)}</p></div><p class="meta">대시보드 생성 ${esc(data.meta.generated_at)}<br>${esc(scope)}</p></section>`;
  }

  function kpis(items) {
    return `<section class="kpis">${items.map(([label, value, unit = ""]) => `<article class="card kpi"><span>${esc(label)}</span><strong>${esc(value)}</strong>${unit ? `<small>${esc(unit)}</small>` : ""}</article>`).join("")}</section>`;
  }

  function coverageView(data) {
    const inventory = data.national_inventory;
    const s = inventory.summary;
    const source = inventory.source;
    const capacityLeader = [...inventory.regions].sort((left, right) => right.capacity_mw - left.capacity_mw)[0];
    const recordLeader = [...inventory.regions].sort((left, right) => right.generator_records - left.generator_records)[0];
    return hero(
      data,
      "전국 태양광 발전설비 현황",
      "학습 데이터 보유 여부와 관계없이 전국의 계통연계 태양광 설비 등록 현황을 지역별로 집계합니다.",
      `${source.provider} ${source.source_system} · 기준일 ${source.reference_date}`,
    ) +
      kpis([
        ["태양광 설비 등록 레코드", format.format(s.generator_records), "건"],
        ["총 설비용량", number(s.total_capacity_mw, 2), "MW"],
        ["최대 설비용량 지역", capacityLeader.region, `${number(capacityLeader.capacity_mw, 2)} MW`],
        ["최다 등록 지역", recordLeader.region, `${format.format(recordLeader.generator_records)}건`],
      ]) +
      `<p class="data-caption">※ 건수는 물리적 발전소 개소가 아닌 EPSIS 설비 등록 레코드 기준입니다.</p>` +
      `<section class="grid national-grid"><article class="card panel map-panel"><div class="panel-head"><div><h2>시도별 태양광 설비 분포</h2><p class="panel-note">색이 진할수록 선택한 지표가 큽니다. 지역을 가리키거나 클릭하면 상세 수치를 볼 수 있습니다.</p></div><div class="metric-switch" role="group" aria-label="지도 및 순위 기준"><button type="button" class="active" data-map-metric="capacity" aria-pressed="true">설비용량</button><button type="button" data-map-metric="records" aria-pressed="false">등록건수</button></div></div><div id="map" role="img" aria-label="전국 시도별 태양광 설비 단계구분도"></div><div class="choropleth-legend" aria-hidden="true"><strong id="map-scale-title">설비용량</strong><span>낮음</span><span class="color-scale"><i></i><i></i><i></i><i></i><i></i></span><span>높음</span></div></article>` +
      `<article class="card panel"><h2 id="region-table-title">시도별 설비용량 순위</h2><p class="panel-note">지도 기준을 바꾸면 순위와 비교 막대도 함께 바뀝니다.</p><div id="region-table">${nationalRegionTable(inventory.regions, "capacity")}</div></article></section>`;
  }

  function qualityView(data) {
    const s = data.summary;
    const inventory = data.data_inventory;
    return hero(
      data,
      "발전소·지역 매핑 및 학습 데이터 품질",
      "발전량 학습 표본의 행정구역, 발전소 좌표, ASOS 관측소와 센서 품질을 검증합니다.",
      data.meta.training_scope,
    ) +
      kpis([
        ["태양광 학습 가능", format.format(s.eligible_solar_assets), "개 자산"],
        ["격리", format.format(s.quarantined_solar_assets), "개 자산"],
        ["원본 CSV", format.format(inventory.role_counts.provider_original || 0), "개 파일"],
        ["표준 명명 파일", format.format(inventory.filename_counts.canonical || 0), "개 파일"],
      ]) +
      `<section class="grid"><article class="card panel"><h2>매핑 검증</h2><p class="panel-note">현재 registry 구조와 한국 좌표 범위를 자동 검사한 결과입니다.</p><div class="checks">${data.mapping.validation.map(check => `<div class="check"><span>${esc(check.check)}</span>${check.passed ? badge("통과", "ok") : badge(`${check.violations}건`, "fail")}</div>`).join("")}</div><h2 class="section-gap">학습 표본 지도</h2><p class="panel-note">발전소 실좌표와 학습에 연결된 ASOS 대리좌표를 구분합니다.</p><div id="map"></div><div class="legend"><span><i class="dot exact"></i>발전소 실좌표</span><span><i class="dot proxy"></i>ASOS 대리좌표</span><span><i class="dot quarantine"></i>매핑 검토 필요</span></div></article>` +
      `<article class="card panel"><h2>파일 포맷 판단</h2><p class="notice">${esc(inventory.decision)}</p><div class="table-wrap compact"><table><thead><tr><th>인코딩</th><th class="num">파일</th></tr></thead><tbody>${Object.entries(inventory.encoding_counts).map(([key, value]) => `<tr><td>${esc(key)}</td><td class="num">${format.format(value)}</td></tr>`).join("")}</tbody></table></div><p class="footer-note">한글이 깨져 보이는 주된 원인은 CP949 원본을 UTF-8로 열거나, BOM 없는 UTF-8을 Excel에서 잘못 감지하는 경우입니다. Bronze 원본은 바이트 그대로 보존하고 표준화 계층만 UTF-8로 변환합니다.</p></article></section>` +
      `<section class="card panel panel-spaced"><h2>발전소별 학습 준비도</h2><p class="panel-note">태양광 학습 표본만 표시합니다. 격리 자산은 매핑 또는 품질 근거 보강 전까지 학습에서 제외됩니다.</p>${plantTable(data.plants)}</section>`;
  }

  function nationalRegionTable(regions, metric = "capacity") {
    const metricKey = metric === "records" ? "generator_records" : "capacity_mw";
    const ordered = [...regions].sort((left, right) => Number(right[metricKey]) - Number(left[metricKey]));
    const maximum = Math.max(...ordered.map((row) => Number(row[metricKey]) || 0), 1);
    return `<div class="table-wrap"><table><thead><tr><th>순위</th><th>시도</th><th class="num">등록건수</th><th class="num">용량(MW)</th></tr></thead><tbody>${ordered.map((row, index) => `<tr><td class="rank">${index + 1}</td><td><strong>${esc(row.region)}</strong><div class="mini-track"><span style="width:${Math.max(2, (Number(row[metricKey]) / maximum) * 100)}%"></span></div></td><td class="num">${format.format(row.generator_records)}</td><td class="num">${number(row.capacity_mw, 2)}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function plantTable(plants) {
    return `<div class="table-wrap"><table><thead><tr><th>발전사</th><th>발전소</th><th>행정구역</th><th>ASOS</th><th>위치표시</th><th>상태</th><th>센서위험</th><th class="num">시간커버리지</th><th class="num">기상결측</th></tr></thead><tbody>${plants.map(row => `<tr><td>${esc(row.company_name)}</td><td>${esc(row.plant)}</td><td>${esc(row.admin_province)} ${esc(row.admin_city || "")}</td><td>${esc(row.weather_station_name)}</td><td>${esc(row.location_basis)}</td><td>${row.status === "eligible" ? badge("학습 가능", "ok") : badge("격리", "warn")}</td><td>${row.sensor_risk === "low" ? badge("낮음", "ok") : badge(row.sensor_risk, "warn")}</td><td class="num">${pct(row.hourly_coverage)}</td><td class="num">${pct(row.missing_weather_rate)}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function drawNationalMap(inventory, boundaries) {
    const node = document.getElementById("map");
    if (!node || typeof L === "undefined") return;
    if (!boundaries) {
      node.innerHTML = '<p class="map-empty">시도 경계 데이터를 불러오지 못했습니다. 우측 표에서 지역별 수치를 확인해 주세요.</p>';
      return;
    }

    const map = L.map(node, {
      zoomControl: false,
      preferCanvas: true,
      scrollWheelZoom: false,
      attributionControl: false,
      minZoom: 5,
      maxZoom: 9,
      zoomSnap: 0.25,
      zoomDelta: 0.5,
    });
    L.control.zoom({ position: "bottomright" }).addTo(map);
    L.control.attribution({ position: "bottomright", prefix: false })
      .addAttribution("경계 © SimpleMaps")
      .addTo(map);
    const regionNames = {
      KR11: "서울특별시", KR26: "부산광역시", KR27: "대구광역시", KR28: "인천광역시", KR29: "광주광역시", KR30: "대전광역시", KR31: "울산광역시", KR41: "경기도", KR42: "강원특별자치도", KR43: "충청북도", KR44: "충청남도", KR45: "전북특별자치도", KR46: "전라남도", KR47: "경상북도", KR48: "경상남도", KR49: "제주특별자치도", KR50: "세종특별자치시",
    };
    const shortRegionNames = {
      서울특별시: "서울", 부산광역시: "부산", 대구광역시: "대구", 인천광역시: "인천", 광주광역시: "광주", 대전광역시: "대전", 울산광역시: "울산", 세종특별자치시: "세종", 경기도: "경기", 강원특별자치도: "강원", 충청북도: "충북", 충청남도: "충남", 전북특별자치도: "전북", 전라남도: "전남", 경상북도: "경북", 경상남도: "경남", 제주특별자치도: "제주",
    };
    const regionByName = new Map(inventory.regions.map((row) => [row.region, row]));
    const locationsByRegion = new Map();
    (inventory.locations || []).forEach((location) => {
      const rows = locationsByRegion.get(location.region) || [];
      rows.push(location);
      locationsByRegion.set(location.region, rows);
    });
    const colors = ["#e8f3ef", "#c1dfd4", "#84c3ad", "#3f9a78", "#0d6049"];
    let metric = "capacity";

    function metricKey() {
      return metric === "records" ? "generator_records" : "capacity_mw";
    }

    function metricMaximum() {
      const key = metricKey();
      return Math.max(...inventory.regions.map((row) => Number(row[key]) || 0), 1);
    }

    function fillColor(value) {
      const ratio = Number(value || 0) / metricMaximum();
      if (ratio >= 0.72) return colors[4];
      if (ratio >= 0.42) return colors[3];
      if (ratio >= 0.20) return colors[2];
      if (ratio >= 0.08) return colors[1];
      return colors[0];
    }

    function styleRegion(feature) {
      const row = regionByName.get(regionNames[feature.properties.id]);
      return {
        color: "#496c61",
        weight: 1.8,
        fillColor: fillColor(row ? row[metricKey()] : 0),
        fillOpacity: 0.94,
      };
    }

    function labelContent(name, row) {
      if (!row) return esc(shortRegionNames[name] || name);
      const value = metric === "records"
        ? `${format.format(row.generator_records)}건`
        : `${number(row.capacity_mw, 0)} MW`;
      return `<strong>${esc(shortRegionNames[name] || name)}</strong><span>${esc(value)}</span>`;
    }

    function popupContent(name, row) {
      if (!row) return `<strong>${esc(name)}</strong>`;
      const key = metricKey();
      const locations = [...(locationsByRegion.get(name) || [])]
        .sort((left, right) => Number(right[key]) - Number(left[key]))
        .slice(0, 3);
      const detailLabel = metric === "records" ? "등록건수 상위 세부지역" : "설비용량 상위 세부지역";
      const details = locations.map((location) => {
        const value = metric === "records"
          ? `${format.format(location.generator_records)}건`
          : `${number(location.capacity_mw, 1)} MW`;
        return `<li><span>${esc(location.subregion.replace(`${name} `, ""))}</span><strong>${esc(value)}</strong></li>`;
      }).join("");
      return `<div class="map-popup"><h3>${esc(name)}</h3><dl><div><dt>설비 등록</dt><dd>${format.format(row.generator_records)}건</dd></div><div><dt>설비용량</dt><dd>${number(row.capacity_mw, 2)} MW</dd></div></dl>${details ? `<p>${detailLabel}</p><ol>${details}</ol>` : ""}</div>`;
    }

    const provinceLayer = L.geoJSON(boundaries, {
      style: styleRegion,
      onEachFeature: (feature, layer) => {
        const name = regionNames[feature.properties.id] || feature.properties.name;
        const row = regionByName.get(name);
        layer._dashboardRegionName = name;
        layer._dashboardRegionRow = row;
        layer.bindTooltip(labelContent(name, row), {
          permanent: false,
          sticky: true,
          direction: "top",
          className: "province-hover-tooltip",
        });
        layer.on({
          mouseover: (event) => {
            event.target.setStyle({ color: "#183f34", weight: 3, fillOpacity: 1 });
            event.target.bringToFront();
          },
          mouseout: (event) => provinceLayer.resetStyle(event.target),
          click: () => layer.bindPopup(popupContent(name, row), { maxWidth: 300 }).openPopup(),
        });
      },
    }).addTo(map);

    const displayBounds = L.latLngBounds([[32.9, 124.5], [38.8, 130.1]]);
    map.fitBounds(displayBounds, { padding: [14, 14] });
    map.setZoom(Math.min(map.getZoom() + 0.5, map.getMaxZoom()));
    map.setMaxBounds(displayBounds.pad(0.22));

    function updateMetric(nextMetric) {
      metric = nextMetric;
      provinceLayer.setStyle(styleRegion);
      provinceLayer.eachLayer((layer) => {
        layer.setTooltipContent(labelContent(layer._dashboardRegionName, layer._dashboardRegionRow));
      });
      document.querySelectorAll("[data-map-metric]").forEach((button) => {
        const active = button.dataset.mapMetric === metric;
        button.classList.toggle("active", active);
        button.setAttribute("aria-pressed", String(active));
      });
      const table = document.getElementById("region-table");
      const title = document.getElementById("region-table-title");
      const scaleTitle = document.getElementById("map-scale-title");
      if (table) table.innerHTML = nationalRegionTable(inventory.regions, metric);
      if (title) title.textContent = metric === "records" ? "시도별 등록건수 순위" : "시도별 설비용량 순위";
      if (scaleTitle) scaleTitle.textContent = metric === "records" ? "등록건수" : "설비용량";
    }

    document.querySelectorAll("[data-map-metric]").forEach((button) => {
      button.addEventListener("click", () => updateMetric(button.dataset.mapMetric));
    });
  }

  function drawTrainingMap(plants) {
    const node = document.getElementById("map");
    if (!node || typeof L === "undefined") return;
    const map = L.map(node, { zoomControl: true, preferCanvas: true }).setView([36.2, 127.8], 7);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 18, attribution: "&copy; OpenStreetMap contributors" }).addTo(map);
    const bounds = [];
    plants.filter(row => row.latitude != null && row.longitude != null).forEach(row => {
      const color = row.status !== "eligible" ? "#a45a00" : row.location_basis === "plant_coordinate" ? "#087f5b" : "#2878b5";
      const marker = L.circleMarker([row.latitude, row.longitude], { radius: row.location_basis === "plant_coordinate" ? 7 : 5, color, fillColor: color, fillOpacity: 0.75, weight: 2 });
      marker.bindPopup(`<strong>${esc(row.plant)}</strong><br>${esc(row.company_name)}<br>행정구역: ${esc(row.admin_province)} ${esc(row.admin_city || "")}<br>기상관측소: ${esc(row.weather_station_name)}<br>좌표 유형: ${esc(row.location_basis)}`);
      marker.addTo(map);
      bounds.push([row.latitude, row.longitude]);
    });
    if (bounds.length) map.fitBounds(bounds, { padding: [24, 24], maxZoom: 9 });
  }
})();
