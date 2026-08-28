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

  function safeUrl(value) {
    try {
      const url = new URL(String(value));
      return ["https:", "http:"].includes(url.protocol) ? esc(url.href) : "#";
    } catch (_) {
      return "#";
    }
  }

  function hero(data, title, description, scope) {
    return `<section class="hero"><div><p class="eyebrow">National solar intelligence</p><h1>${esc(title)}</h1><p class="lede">${esc(description)}</p></div><p class="meta">대시보드 생성 ${esc(data.meta.generated_at)}<br>${esc(scope)}</p></section>`;
  }

  function kpis(items) {
    return `<section class="kpis">${items.map(([label, value, unit = ""]) => `<article class="card kpi"><span>${esc(label)}</span><strong>${value}</strong>${unit ? `<small>${esc(unit)}</small>` : ""}</article>`).join("")}</section>`;
  }

  function coverageView(data) {
    const inventory = data.national_inventory;
    const s = inventory.summary;
    const source = inventory.source;
    return hero(
      data,
      "전국 태양광 발전설비 현황",
      "학습 데이터 보유 여부와 관계없이 전국의 계통연계 태양광 설비 등록 현황을 지역별로 집계합니다.",
      `${source.provider} ${source.source_system} · 기준일 ${source.reference_date}`,
    ) +
      kpis([
        ["태양광 설비 등록 레코드", format.format(s.generator_records), "건"],
        ["총 설비용량", number(s.total_capacity_mw, 2), "MW"],
        ["광역자치단체", format.format(s.regions_with_records), "개 시도"],
        ["표준 세부지역", format.format(s.subregions), `${format.format(s.source_subregion_labels)}개 원본명 통합`],
      ]) +
      `<p class="notice"><strong>집계 단위 안내</strong><br>한 행은 물리적 발전소 개소가 아니라 EPSIS 발전기·등록 레코드입니다. 학습 가능 여부는 이 수치에 영향을 주지 않으며, 학습 표본은 상단의 ‘학습 매핑·품질’ 화면에서 별도로 관리합니다.</p>` +
      `<section class="grid national-grid"><article class="card panel"><h2>전국 지역 분포</h2><p class="panel-note">시도 경계 색은 등록 레코드 수, 원은 세부지역의 설비용량을 나타냅니다. 좌표가 없는 새 행정구역은 시도 중심 대리좌표로 표시되며 지역 집계에서는 제외하지 않습니다.</p><div id="map"></div><div class="legend"><span><i class="dot region-fill"></i>시도별 등록 규모</span><span><i class="dot location-dot"></i>세부지역 집계</span><span><i class="dot proxy"></i>시도 중심 대리좌표</span></div></article>` +
      `<article class="card panel"><h2>시도별 설비 현황</h2><p class="panel-note">공식 원본의 중복 의심 레코드를 임의 삭제하지 않은 수치입니다.</p>${nationalRegionTable(inventory.regions)}</article></section>` +
      `<section class="grid"><article class="card panel"><h2>수집 범위와 출처</h2>${sourcePanel(source)}</article><article class="card panel"><h2>원천 품질 점검</h2>${nationalQuality(inventory.quality, source)}</article></section>`;
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

  function nationalRegionTable(regions) {
    const ordered = [...regions].sort((left, right) => right.generator_records - left.generator_records);
    const maxRecords = Math.max(...ordered.map((row) => Number(row.generator_records) || 0), 1);
    const totalCapacity = ordered.reduce((sum, row) => sum + Number(row.capacity_mw || 0), 0);
    return `<div class="table-wrap"><table><thead><tr><th>순위</th><th>시도</th><th class="num">설비 레코드</th><th class="num">총용량(MW)</th><th class="num">용량 비중</th></tr></thead><tbody>${ordered.map((row, index) => `<tr><td class="rank">${index + 1}</td><td><strong>${esc(row.region)}</strong><div class="mini-track"><span style="width:${Math.max(2, (row.generator_records / maxRecords) * 100)}%"></span></div></td><td class="num">${format.format(row.generator_records)}</td><td class="num">${number(row.capacity_mw, 2)}</td><td class="num">${pct(totalCapacity ? row.capacity_mw / totalCapacity : 0)}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function sourcePanel(source) {
    const limitations = (source.limitations || []).map((item) => `<li>${esc(item)}</li>`).join("");
    return `<dl class="source-grid"><div><dt>제공기관</dt><dd>${esc(source.provider)} ${esc(source.source_system)}</dd></div><div><dt>기준일</dt><dd>${esc(source.reference_date)}</dd></div><div><dt>다운로드</dt><dd>${esc(source.downloaded_at)}</dd></div><div><dt>원본 형식</dt><dd>${esc(source.encoding)}${source.bytes ? ` · ${number(source.bytes / 1024 / 1024, 1)} MB` : ""}</dd></div></dl><p><a class="source-link" href="${safeUrl(source.source_url)}" target="_blank" rel="noreferrer">공식 EPSIS 원천 화면 열기</a></p><ul class="source-list">${limitations}</ul><p class="footer-note">원본 SHA-256 <code>${esc(source.sha256)}</code></p>`;
  }

  function nationalQuality(quality, source) {
    const checks = [
      ["원본 SHA-256", source.sha256_verified === true, 1],
      ["필수 11개 컬럼", quality.schema_valid, quality.schema_valid ? 0 : 1],
      ["합계행 레코드 수", quality.footer.record_count_matches, 1],
      ["합계행 설비용량", quality.footer.capacity_matches, 1],
      ["음수 설비용량", quality.negative_capacity_records === 0, quality.negative_capacity_records],
      ["미분류 광역지역", quality.unknown_region_records === 0, quality.unknown_region_records],
    ];
    return `<div class="checks">${checks.map(([label, passed, violations]) => `<div class="check"><span>${esc(label)}</span>${passed ? badge("통과", "ok") : badge(`${format.format(violations)}건`, "warn")}</div>`).join("")}</div><dl class="quality-stats"><div><dt>완전 동일 중복</dt><dd>${format.format(quality.exact_duplicate_records || 0)}건</dd></div><div><dt>0 MW 레코드</dt><dd>${format.format(quality.zero_capacity_records || 0)}건</dd></div><div><dt>좌표 대리표시</dt><dd>${format.format((quality.coordinate_basis_counts || {}).province_centroid || 0)}개 지역</dd></div><div><dt>한글 대체문자</dt><dd>${format.format(quality.replacement_character_cells || 0)}셀</dd></div></dl><p class="footer-note">동일 레코드는 공식 고유키가 없어 제거하지 않습니다. 합계행은 검증에만 사용하고 지역·용량 집계에서는 제외합니다.</p>`;
  }

  function plantTable(plants) {
    return `<div class="table-wrap"><table><thead><tr><th>발전사</th><th>발전소</th><th>행정구역</th><th>ASOS</th><th>위치표시</th><th>상태</th><th>센서위험</th><th class="num">시간커버리지</th><th class="num">기상결측</th></tr></thead><tbody>${plants.map(row => `<tr><td>${esc(row.company_name)}</td><td>${esc(row.plant)}</td><td>${esc(row.admin_province)} ${esc(row.admin_city || "")}</td><td>${esc(row.weather_station_name)}</td><td>${esc(row.location_basis)}</td><td>${row.status === "eligible" ? badge("학습 가능", "ok") : badge("격리", "warn")}</td><td>${row.sensor_risk === "low" ? badge("낮음", "ok") : badge(row.sensor_risk, "warn")}</td><td class="num">${pct(row.hourly_coverage)}</td><td class="num">${pct(row.missing_weather_rate)}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function drawNationalMap(inventory, boundaries) {
    const node = document.getElementById("map");
    if (!node || typeof L === "undefined") return;
    const map = L.map(node, { zoomControl: true, preferCanvas: true }).setView([36.2, 127.8], 7);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 18, attribution: "&copy; OpenStreetMap contributors" }).addTo(map);
    const regionNames = {
      KR11: "서울특별시", KR26: "부산광역시", KR27: "대구광역시", KR28: "인천광역시", KR29: "광주광역시", KR30: "대전광역시", KR31: "울산광역시", KR41: "경기도", KR42: "강원특별자치도", KR43: "충청북도", KR44: "충청남도", KR45: "전북특별자치도", KR46: "전라남도", KR47: "경상북도", KR48: "경상남도", KR49: "제주특별자치도", KR50: "세종특별자치시",
    };
    const regionByName = new Map(inventory.regions.map((row) => [row.region, row]));
    const maxRecords = Math.max(...inventory.regions.map((row) => Number(row.generator_records) || 0), 1);
    if (boundaries) {
      L.geoJSON(boundaries, {
        style: (feature) => {
          const row = regionByName.get(regionNames[feature.properties.id]);
          const ratio = row ? Math.sqrt(row.generator_records / maxRecords) : 0;
          return { color: "#ffffff", weight: 1.4, fillColor: "#087f5b", fillOpacity: 0.12 + ratio * 0.68 };
        },
        onEachFeature: (feature, layer) => {
          const name = regionNames[feature.properties.id] || feature.properties.name;
          const row = regionByName.get(name);
          layer.bindPopup(row ? `<strong>${esc(name)}</strong><br>설비 등록 레코드: ${format.format(row.generator_records)}건<br>총 설비용량: ${number(row.capacity_mw, 2)} MW` : `<strong>${esc(name)}</strong>`);
        },
      }).addTo(map);
    }
    const bounds = [];
    (inventory.locations || []).filter((row) => row.latitude != null && row.longitude != null).forEach((row) => {
      const proxy = row.coordinate_basis === "province_centroid";
      const radius = Math.min(14, 3.5 + Math.log10(Math.max(1, row.capacity_mw)) * 2.4);
      const color = proxy ? "#2878b5" : "#f08c00";
      const marker = L.circleMarker([row.latitude, row.longitude], { radius, color: "#ffffff", fillColor: color, fillOpacity: 0.82, weight: 1.2 });
      marker.bindPopup(`<strong>${esc(row.subregion)}</strong><br>${esc(row.region)}<br>설비 등록 레코드: ${format.format(row.generator_records)}건<br>총 설비용량: ${number(row.capacity_mw, 2)} MW<br>좌표 근거: ${esc(row.coordinate_basis)}`);
      marker.addTo(map);
      bounds.push([row.latitude, row.longitude]);
    });
    if (bounds.length && !boundaries) map.fitBounds(bounds, { padding: [24, 24], maxZoom: 8 });
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
