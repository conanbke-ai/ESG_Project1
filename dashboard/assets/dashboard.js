(async function () {
  const app = document.getElementById("app");
  const view = document.body.dataset.view || "coverage";
  const format = new Intl.NumberFormat("ko-KR");
  const pct = (value) => value == null ? "-" : `${(value * 100).toFixed(1)}%`;
  const number = (value, digits = 0) => value == null ? "-" : Number(value).toLocaleString("ko-KR", { maximumFractionDigits: digits });
  const badge = (text, kind) => `<span class="badge ${kind}">${text}</span>`;
  const esc = (value) => String(value ?? "-").replace(/[&<>'"]/g, (ch) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[ch]));

  try {
    const response = await fetch("data/dashboard_data.json", { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const data = await response.json();
    app.innerHTML = view === "quality" ? qualityView(data) : coverageView(data);
    drawMap(data.plants);
  } catch (error) {
    app.innerHTML = `<p class="error">대시보드 데이터를 읽지 못했습니다. 프로젝트 루트에서 <code>python app.py build-dashboard</code>를 실행한 뒤 로컬 HTTP 서버로 <code>dashboard/</code>를 열어주세요. (${esc(error.message)})</p>`;
  }

  function hero(data, title, description) {
    return `<section class="hero"><div><p class="eyebrow">National solar data portfolio</p><h1>${title}</h1><p class="lede">${description}</p></div><p class="meta">생성 시각 ${esc(data.meta.generated_at)}<br>${esc(data.meta.scope)}</p></section>`;
  }

  function kpis(items) {
    return `<section class="kpis">${items.map(([label, value]) => `<article class="card kpi"><span>${label}</span><strong>${value}</strong></article>`).join("")}</section>`;
  }

  function coverageView(data) {
    const s = data.summary;
    return hero(data, "전국 태양광 발전 데이터 현황", "전국 전체 인허가 설비 수가 아니라, 공식 시간별 발전실적과 기상자료를 함께 확보한 학습 포트폴리오입니다.") +
      kpis([
        ["태양광 등록 자산", format.format(s.solar_assets)],
        ["학습 가능", format.format(s.eligible_solar_assets)],
        ["매핑 검토 필요", format.format(s.quarantined_solar_assets)],
        ["Gold 전체 관측행", format.format(s.model_rows_all_energy)],
      ]) +
      `<p class="notice">${esc(data.meta.location_policy)}</p>` +
      `<section class="grid"><article class="card panel"><h2>발전소 및 기상 대리좌표</h2><p class="panel-note">초록은 발전소 실좌표, 파랑은 ASOS 관측소 대리좌표입니다. 대리좌표를 실제 발전소 위치로 해석하면 안 됩니다.</p><div id="map"></div><div class="legend"><span><i class="dot exact"></i>발전소 실좌표</span><span><i class="dot proxy"></i>ASOS 대리좌표</span><span><i class="dot quarantine"></i>매핑 검토 필요</span></div></article>` +
      `<article class="card panel"><h2>지역별 학습 범위</h2><p class="panel-note">행정구역이 확인된 태양광 자산 기준입니다.</p>${regionTable(data.regions)}</article></section>` +
      `<section class="grid"><article class="card panel"><h2>발전사별 구성</h2>${bars(data.companies)}</article><article class="card panel"><h2>데이터 기간</h2><p class="lede">${esc(s.generation_start)}<br>— ${esc(s.generation_end)}</p><p class="footer-note">Gold 표에는 모든 발전원을 보존하지만 XGBoost와 CNN-BiLSTM 학습 시 <code>energy_source=solar</code>와 품질 마스크를 적용합니다.</p></article></section>`;
  }

  function qualityView(data) {
    const s = data.summary;
    const inventory = data.data_inventory;
    return hero(data, "발전소·지역 매핑 및 품질", "행정구역, 실제 발전소 좌표, 학습용 ASOS 관측소를 분리해 검증합니다. 근거가 없는 매핑은 자동 추정하지 않고 격리합니다.") +
      kpis([
        ["태양광 학습 가능", format.format(s.eligible_solar_assets)],
        ["격리", format.format(s.quarantined_solar_assets)],
        ["원본 CSV", format.format(inventory.role_counts.provider_original || 0)],
        ["표준 명명 파일", format.format(inventory.filename_counts.canonical || 0)],
      ]) +
      `<section class="grid"><article class="card panel"><h2>매핑 검증</h2><p class="panel-note">현재 registry 구조와 한국 좌표 범위를 자동 검사한 결과입니다.</p><div class="checks">${data.mapping.validation.map(check => `<div class="check"><span>${esc(check.check)}</span>${check.passed ? badge("통과", "ok") : badge(`${check.violations}건`, "fail")}</div>`).join("")}</div><h2 style="margin-top:22px">지도 검수</h2><div id="map"></div></article>` +
      `<article class="card panel"><h2>파일 포맷 판단</h2><p class="notice">${esc(inventory.decision)}</p><div class="table-wrap"><table><thead><tr><th>인코딩</th><th class="num">파일</th></tr></thead><tbody>${Object.entries(inventory.encoding_counts).map(([key, value]) => `<tr><td>${esc(key)}</td><td class="num">${format.format(value)}</td></tr>`).join("")}</tbody></table></div><p class="footer-note">한글이 깨져 보이는 주된 원인은 CP949 원본을 UTF-8로 열거나, BOM 없는 UTF-8을 Excel에서 잘못 감지하는 경우입니다. 원본 자체를 일괄 재인코딩하지 않습니다.</p></article></section>` +
      `<section class="card panel" style="margin-top:18px"><h2>발전소별 학습 준비도</h2><p class="panel-note">태양광만 표시합니다. 격리 자산은 매핑 근거 보강 전까지 학습에서 제외됩니다.</p>${plantTable(data.plants)}</section>`;
  }

  function regionTable(regions) {
    return `<div class="table-wrap"><table><thead><tr><th>지역</th><th class="num">전체</th><th class="num">학습</th><th class="num">용량(MW)</th></tr></thead><tbody>${regions.map(row => `<tr><td>${esc(row.name)}</td><td class="num">${format.format(row.assets)}</td><td class="num">${format.format(row.eligible)}</td><td class="num">${number(row.known_capacity_mw, 2)}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function plantTable(plants) {
    return `<div class="table-wrap"><table><thead><tr><th>발전사</th><th>발전소</th><th>행정구역</th><th>ASOS</th><th>위치표시</th><th>상태</th><th>센서위험</th><th class="num">시간커버리지</th><th class="num">기상결측</th></tr></thead><tbody>${plants.map(row => `<tr><td>${esc(row.company_name)}</td><td>${esc(row.plant)}</td><td>${esc(row.admin_province)} ${esc(row.admin_city || "")}</td><td>${esc(row.weather_station_name)}</td><td>${esc(row.location_basis)}</td><td>${row.status === "eligible" ? badge("학습 가능", "ok") : badge("격리", "warn")}</td><td>${row.sensor_risk === "low" ? badge("낮음", "ok") : badge(esc(row.sensor_risk), "warn")}</td><td class="num">${pct(row.hourly_coverage)}</td><td class="num">${pct(row.missing_weather_rate)}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function bars(companies) {
    const max = Math.max(...companies.map(row => row.eligible), 1);
    return `<div class="bars">${companies.map(row => `<div class="bar-row"><span>${esc(row.name)}</span><div class="track"><div class="fill" style="width:${(row.eligible / max) * 100}%"></div></div><strong>${format.format(row.eligible)}</strong></div>`).join("")}</div>`;
  }

  function drawMap(plants) {
    const node = document.getElementById("map");
    if (!node || typeof L === "undefined") return;
    const map = L.map(node, { zoomControl: true }).setView([36.2, 127.8], 7);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 18, attribution: "&copy; OpenStreetMap contributors" }).addTo(map);
    const bounds = [];
    plants.filter(row => row.latitude != null && row.longitude != null).forEach(row => {
      const color = row.status !== "eligible" ? "#a45a00" : row.location_basis === "plant_coordinate" ? "#087f5b" : "#2878b5";
      const marker = L.circleMarker([row.latitude, row.longitude], { radius: row.location_basis === "plant_coordinate" ? 7 : 5, color, fillColor: color, fillOpacity: .75, weight: 2 });
      marker.bindPopup(`<strong>${esc(row.plant)}</strong><br>${esc(row.company_name)}<br>행정구역: ${esc(row.admin_province)} ${esc(row.admin_city || "")}<br>기상관측소: ${esc(row.weather_station_name)}<br>좌표 유형: ${esc(row.location_basis)}`);
      marker.addTo(map);
      bounds.push([row.latitude, row.longitude]);
    });
    if (bounds.length) map.fitBounds(bounds, { padding: [24, 24], maxZoom: 9 });
  }
})();
