(async function () {
  const app = document.getElementById("app");
  const view = document.body.dataset.view || "coverage";
  const int = new Intl.NumberFormat("ko-KR");
  const korean = new Intl.Collator("ko", { numeric: true, sensitivity: "base" });
  const state = {
    nationalMetric: "capacity",
    selectedRegion: null,
    subregionQuery: "",
    detailSortKey: "capacity",
    detailSortDirection: "desc",
    tab: "comparison",
    region: "all",
    plantId: "all",
    forecastModelId: "all",
  };
  const metrics = {
    nmae_capacity: { label: "용량 정규화 MAE", short: "NMAE", unit: "%", better: "lower", digits: 2 },
    mae: { label: "평균 절대 오차", short: "MAE", unit: "MWh", better: "lower", digits: 3 },
    rmse: { label: "평균 제곱근 오차", short: "RMSE", unit: "MWh", better: "lower", digits: 3 },
    r2: { label: "설명력", short: "R²", unit: "", better: "higher", digits: 3 },
  };

  try {
    const response = await fetch("data/dashboard_data.json", { cache: "no-store" });
    if (!response.ok) throw new Error("dashboard data unavailable");
    const data = await response.json();
    if (view === "analysis") {
      const analysis = data.model_analysis || emptyAnalysis();
      app.innerHTML = analysisPage(analysis);
      setupAnalysis(analysis);
    } else if (view === "forecast") {
      const analysis = data.model_analysis || emptyAnalysis();
      app.innerHTML = forecastPage(analysis);
      setupForecast(analysis);
    } else {
      let boundaries = null;
      try {
        const boundaryResponse = await fetch("data/korea_provinces.geojson", { cache: "no-store" });
        if (boundaryResponse.ok) boundaries = await boundaryResponse.json();
      } catch (boundaryError) {
        console.warn("province boundaries unavailable", boundaryError);
      }
      app.innerHTML = coveragePage(data.national_inventory);
      setupCoverage(data.national_inventory, boundaries);
    }
  } catch (error) {
    console.error(error);
    app.innerHTML = `<section class="empty-state" role="alert"><h1>화면을 불러오지 못했습니다</h1><p>잠시 후 새로고침해 주세요.</p></section>`;
  }

  function esc(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (character) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;",
    })[character]);
  }

  function finite(value) {
    return value !== null && value !== "" && Number.isFinite(Number(value));
  }

  function number(value, digits = 0) {
    if (!finite(value)) return "-";
    return Number(value).toLocaleString("ko-KR", { maximumFractionDigits: digits });
  }

  function ratio(value) {
    return finite(value) ? `${number(Number(value) * 100, 1)}%` : "-";
  }

  function intro(title, description, context = "") {
    return `<section class="page-intro"><div><h1>${esc(title)}</h1><p>${esc(description)}</p></div>${context ? `<p class="page-context">${esc(context)}</p>` : ""}</section>`;
  }

  function stats(items) {
    return `<section class="summary-grid">${items.map(([label, value, unit = ""]) => `<article class="summary-item"><span>${esc(label)}</span><div class="summary-value"><strong>${esc(value)}</strong>${unit ? `<small>${esc(unit)}</small>` : ""}</div></article>`).join("")}</section>`;
  }

  function emptyState(title, message) {
    return `<div class="empty-state"><h2>${esc(title)}</h2><p>${esc(message)}</p></div>`;
  }

  function shortRegion(region) {
    return ({
      서울특별시: "서울", 부산광역시: "부산", 대구광역시: "대구", 인천광역시: "인천",
      광주광역시: "광주", 전남광주통합특별시: "전남·광주", 대전광역시: "대전", 울산광역시: "울산", 세종특별자치시: "세종",
      경기도: "경기", 강원특별자치도: "강원", 충청북도: "충북", 충청남도: "충남",
      전북특별자치도: "전북", 전라남도: "전남", 경상북도: "경북", 경상남도: "경남",
      제주특별자치도: "제주", unknown: "지역 미확인",
    })[region] || region;
  }

  function displayRegion(region) {
    return !region || region === "unknown" ? "지역 미확인" : region;
  }

  function coveragePage(inventory = {}) {
    const summary = inventory.summary || {};
    const regions = inventory.regions || [];
    const leader = [...regions].sort((a, b) => Number(b.capacity_mw) - Number(a.capacity_mw))[0];
    state.selectedRegion = leader?.region || regions[0]?.region || null;
    return intro("전국 태양광 발전설비 현황", "전국 설비 규모를 한눈에 비교하고 원하는 세부지역을 바로 찾아볼 수 있습니다.") +
      stats([
        ["전국 설비용량", number(summary.total_capacity_mw, 2), "MW"],
        ["설비 등록", int.format(Number(summary.generator_records) || 0), "건"],
        ["설비용량 1위", leader?.region || "-", leader ? `${number(leader.capacity_mw, 2)} MW` : ""],
      ]) +
      `<section class="national-layout">
        <article class="surface map-surface">
          <div class="section-head">
            <div><h2>시도별 설비 분포</h2><p>지도나 목록에서 지역을 선택해 설비 규모를 비교해 보세요.</p></div>
            <div class="national-controls">
              <label class="field-label" for="province-select"><span>시도</span><select id="province-select">${regionOptions(regions)}</select></label>
              <div class="segmented-control" role="group" aria-label="전국 현황 기준">
                <button type="button" class="is-active" data-map-metric="capacity" aria-pressed="true">설비용량</button>
                <button type="button" data-map-metric="records" aria-pressed="false">등록건수</button>
              </div>
            </div>
          </div>
          <div id="map" role="region" aria-label="전국 시도별 태양광 설비 분포 지도"></div>
          <div class="map-scale" aria-hidden="true"><strong id="map-scale-title">설비용량</strong><span>낮음</span><span class="scale-colors"><i></i><i></i><i></i><i></i><i></i></span><span>높음</span></div>
        </article>
        <article class="surface region-surface">
          <div class="section-head"><div><h2 id="region-list-title">시도별 설비용량</h2><p>전국 ${int.format(regions.length)}개 시도의 규모와 순위를 비교합니다. 전남·광주는 2026년 7월 1일 통합 행정구역 기준입니다.</p></div></div>
          <div id="region-list" class="region-list">${regionList(regions)}</div>
        </article>
      </section>
      <section class="surface detail-surface" aria-labelledby="subregion-title">
        <div class="section-head detail-head">
          <div><h2 id="subregion-title"></h2><p id="subregion-caption"></p></div>
          <div class="detail-controls">
            <label class="field-label search-field" for="subregion-search"><span>전국 세부지역 검색</span><input id="subregion-search" type="search" autocomplete="off" placeholder="예: 전남 여수시, 강원"></label>
            <label class="field-label sort-field" for="subregion-sort"><span>정렬 기준</span><select id="subregion-sort">
              <option value="capacity" selected>설비용량</option>
              <option value="records">등록건수</option>
              <option value="name">세부지역</option>
            </select></label>
            <label class="field-label direction-field" for="subregion-order"><span>정렬 방향</span><select id="subregion-order">${detailDirectionOptions()}</select></label>
          </div>
        </div>
        <div id="subregion-summary"></div>
        <div id="subregion-table"></div>
      </section>`;
  }

  function regionOptions(regions) {
    return [...regions].sort((a, b) => a.region.localeCompare(b.region, "ko"))
      .map((row) => `<option value="${esc(row.region)}"${row.region === state.selectedRegion ? " selected" : ""}>${esc(row.region)}</option>`).join("");
  }

  function nationalValue(row, metric = state.nationalMetric) {
    return metric === "records" ? Number(row.generator_records) || 0 : Number(row.capacity_mw) || 0;
  }

  function nationalText(row, metric = state.nationalMetric) {
    return metric === "records" ? `${int.format(Number(row.generator_records) || 0)}건` : `${number(row.capacity_mw, 2)} MW`;
  }

  function regionList(regions) {
    const ordered = [...regions].sort((a, b) => nationalValue(b) - nationalValue(a));
    const maximum = Math.max(...ordered.map((row) => nationalValue(row)), 1);
    return ordered.map((row, index) => {
      const secondary = state.nationalMetric === "records" ? `${number(row.capacity_mw, 1)} MW` : `${int.format(Number(row.generator_records) || 0)}건`;
      const width = Math.max(2, nationalValue(row) / maximum * 100);
      return `<button type="button" class="region-item${row.region === state.selectedRegion ? " is-selected" : ""}" data-region="${esc(row.region)}" aria-pressed="${row.region === state.selectedRegion}">
        <span class="region-rank">${index + 1}</span>
        <span class="region-copy"><strong>${esc(shortRegion(row.region))}</strong><small>${esc(secondary)}</small><i class="value-track"><b style="width:${width.toFixed(2)}%"></b></i></span>
        <span class="region-value">${esc(nationalText(row))}</span>
      </button>`;
    }).join("");
  }

  function cleanSubregion(row) {
    const name = String(row.subregion || "");
    const prefix = `${row.region} `;
    return name.startsWith(prefix) ? name.slice(prefix.length) : name;
  }

  function regionSearchTerms(region) {
    const aliases = {
      강원특별자치도: ["강원도"],
      전북특별자치도: ["전라북도"],
      전남광주통합특별시: ["전남", "전라남도", "광주", "광주시", "광주광역시", "광주전남"],
      제주특별자치도: ["제주도"],
      세종특별자치시: ["세종시"],
    };
    return [region, shortRegion(region), ...(aliases[region] || [])];
  }

  function normalizedSearch(value) {
    return String(value || "").trim().replace(/\s+/g, " ").toLocaleLowerCase("ko");
  }

  function administrativeAliasMatch(inventory, query = state.subregionQuery) {
    const normalized = normalizedSearch(query);
    if (!normalized) return null;
    const match = Object.entries(inventory.location_search_aliases || {})
      .find(([alias]) => normalizedSearch(alias) === normalized);
    return match ? { alias: match[0], location: match[1] } : null;
  }

  function locationSearchTerms(inventory, row) {
    const aliases = Object.entries(inventory.location_search_aliases || {})
      .filter(([, location]) => location === row.subregion)
      .map(([alias]) => alias);
    return [...regionSearchTerms(row.region), cleanSubregion(row), row.subregion, ...aliases];
  }

  function detailDirectionOptions() {
    const labels = state.detailSortKey === "name"
      ? { desc: "가나다 역순", asc: "가나다순" }
      : { desc: "높은 순", asc: "낮은 순" };
    return ["desc", "asc"].map((direction) =>
      `<option value="${direction}"${direction === state.detailSortDirection ? " selected" : ""}>${labels[direction]}</option>`
    ).join("");
  }

  function detailNumericValue(row, key = state.detailSortKey) {
    return key === "records" ? Number(row.generator_records) || 0 : Number(row.capacity_mw) || 0;
  }

  function compareDetailRows(left, right) {
    const direction = state.detailSortDirection === "asc" ? 1 : -1;
    const byName = korean.compare(cleanSubregion(left), cleanSubregion(right))
      || korean.compare(displayRegion(left.region), displayRegion(right.region));
    if (state.detailSortKey === "name") return direction * byName;
    const byValue = detailNumericValue(left) - detailNumericValue(right);
    return byValue ? direction * byValue : byName;
  }

  function detailSortDescription() {
    if (state.detailSortKey === "name") {
      return state.detailSortDirection === "asc" ? "세부지역 가나다순" : "세부지역 역순";
    }
    const metric = state.detailSortKey === "records" ? "등록건수" : "설비용량";
    return `${metric} ${state.detailSortDirection === "asc" ? "낮은 순" : "높은 순"}`;
  }

  function detailAriaSort(key) {
    if (state.detailSortKey !== key) return "";
    return ` aria-sort="${state.detailSortDirection === "asc" ? "ascending" : "descending"}"`;
  }

  function detailSortHeader(key, label, numeric = false) {
    const active = state.detailSortKey === key;
    const indicator = active
      ? state.detailSortDirection === "asc" ? "↑" : "↓"
      : "↕";
    const current = active
      ? state.detailSortDirection === "asc" ? "오름차순" : "내림차순"
      : "정렬 기준 선택";
    return `<th${numeric ? ' class="numeric"' : ""}${detailAriaSort(key)}><button type="button" class="table-sort-button${active ? " is-active" : ""}" data-detail-sort="${key}" aria-label="${esc(`${label}, ${current}`)}">${esc(label)}<span aria-hidden="true">${indicator}</span></button></th>`;
  }

  function detailRows(inventory, query = state.subregionQuery) {
    const normalized = normalizedSearch(query);
    return (inventory.locations || [])
      .filter((row) => {
        if (!normalized) return row.region === state.selectedRegion;
        const searchText = locationSearchTerms(inventory, row)
          .join(" ")
          .toLocaleLowerCase("ko");
        return searchText.includes(normalized);
      })
      .sort(compareDetailRows);
  }

  function detailTable(rows, searching, aliasMatch = null) {
    if (!rows.length) {
      const message = aliasMatch
        ? `${aliasMatch.alias}는 ${aliasMatch.location}에 속하지만 현재 원천에는 해당 세부지역으로 등록된 설비 행이 없습니다.`
        : searching ? "시도명·약칭 또는 세부지역명을 바꿔 검색해 보세요." : "표시할 세부지역 정보가 없습니다.";
      return `<div class="detail-table-shell is-empty">${emptyState("세부지역 결과 없음", message)}</div>`;
    }
    const showTrack = state.detailSortKey !== "name";
    const maximum = showTrack
      ? Math.max(...rows.map((row) => detailNumericValue(row)), 1)
      : 1;
    const rankLabel = state.detailSortKey === "name" ? "순서" : "순위";
    return `<div class="table-responsive detail-table-shell"><table class="data-table subregion-table">
      <thead><tr><th>${rankLabel}</th>${searching ? "<th>시도</th>" : ""}${detailSortHeader("name", "세부지역")}${detailSortHeader("records", "등록건수", true)}${detailSortHeader("capacity", "설비용량(MW)", true)}</tr></thead>
      <tbody>${rows.map((row, index) => `<tr>
        <td class="rank-cell">${index + 1}</td>
        ${searching ? `<td><button type="button" class="text-button" data-detail-region="${esc(row.region)}" aria-label="${esc(row.region)} 상세 보기">${esc(shortRegion(row.region))}</button></td>` : ""}
        <td><div class="name-line"><strong>${esc(cleanSubregion(row))}</strong>${row.source_region_conflict ? '<span class="status-tag caution">지역 확인 필요</span>' : ""}</div>${showTrack ? `<div class="inline-track"><span style="width:${Math.max(1, detailNumericValue(row) / maximum * 100).toFixed(2)}%"></span></div>` : ""}</td>
        <td class="numeric">${int.format(Number(row.generator_records) || 0)}</td><td class="numeric">${number(row.capacity_mw, 2)}</td>
      </tr>`).join("")}</tbody>
    </table></div>`;
  }

  function setupCoverage(inventory, boundaries) {
    const regions = inventory.regions || [];
    const byName = new Map(regions.map((row) => [row.region, row]));
    let mapController;

    function bindRows() {
      document.querySelectorAll("[data-region]").forEach((button) => button.addEventListener("click", () => selectRegion(button.dataset.region)));
    }

    function renderList() {
      document.getElementById("region-list").innerHTML = regionList(regions);
      document.getElementById("region-list-title").textContent = state.nationalMetric === "records" ? "시도별 등록건수" : "시도별 설비용량";
      bindRows();
    }

    function renderDetail() {
      const row = byName.get(state.selectedRegion);
      const searching = Boolean(String(state.subregionQuery).trim());
      const rows = detailRows(inventory);
      const aliasMatch = administrativeAliasMatch(inventory);
      const totals = rows.reduce((summary, item) => ({
        generatorRecords: summary.generatorRecords + (Number(item.generator_records) || 0),
        capacityMw: summary.capacityMw + (Number(item.capacity_mw) || 0),
      }), { generatorRecords: 0, capacityMw: 0 });
      const summary = searching ? totals : {
        generatorRecords: Number(row?.generator_records) || totals.generatorRecords,
        capacityMw: Number(row?.capacity_mw) || totals.capacityMw,
      };
      document.getElementById("subregion-title").textContent = searching ? "전국 세부지역 검색 결과" : `${state.selectedRegion || "선택 지역"} 세부지역`;
      document.getElementById("subregion-caption").textContent = searching
        ? aliasMatch && rows.length
          ? `${aliasMatch.alias}는 ${aliasMatch.location}에 속합니다. 표는 ${cleanSubregion(rows[0])} 전체 등록 집계이며 ${detailSortDescription()}입니다.`
          : `${rows.length}개 결과를 ${detailSortDescription()}으로 표시합니다. 시도를 선택하면 해당 지역의 전체 현황으로 이동합니다.`
        : `${rows.length}개 세부지역을 ${detailSortDescription()}으로 비교합니다.`;
      document.getElementById("subregion-summary").innerHTML = rows.length ? `<div class="selected-summary"><span><small>${searching ? "검색 결과 등록" : "설비 등록"}</small><strong>${int.format(summary.generatorRecords)}건</strong></span><span><small>${searching ? "검색 결과 용량" : "설비용량"}</small><strong>${number(summary.capacityMw, 2)} MW</strong></span></div>` : "";
      document.getElementById("subregion-table").innerHTML = detailTable(rows, searching, aliasMatch);
      document.querySelectorAll("[data-detail-region]").forEach((button) => button.addEventListener("click", () => selectRegion(button.dataset.detailRegion)));
      document.querySelectorAll("[data-detail-sort]").forEach((button) => button.addEventListener("click", () => {
        const key = button.dataset.detailSort;
        const horizontalScroll = button.closest(".detail-table-shell")?.scrollLeft || 0;
        if (state.detailSortKey === key) {
          state.detailSortDirection = state.detailSortDirection === "asc" ? "desc" : "asc";
        } else {
          state.detailSortKey = key;
          state.detailSortDirection = key === "name" ? "asc" : "desc";
        }
        sortSelect.value = state.detailSortKey;
        directionSelect.innerHTML = detailDirectionOptions();
        renderDetail();
        const refreshedShell = document.querySelector("#subregion-table .detail-table-shell");
        if (refreshedShell) refreshedShell.scrollLeft = horizontalScroll;
        document.querySelector(`[data-detail-sort="${key}"]`)?.focus({ preventScroll: true });
      }));
    }

    function selectRegion(region) {
      if (!byName.has(region)) return;
      state.selectedRegion = region;
      state.subregionQuery = "";
      document.getElementById("province-select").value = region;
      document.getElementById("subregion-search").value = "";
      renderList();
      renderDetail();
      mapController?.select(region);
    }

    document.getElementById("province-select").addEventListener("change", (event) => selectRegion(event.target.value));
    const searchInput = document.getElementById("subregion-search");
    let composingSearch = false;
    let searchRenderFrame = null;
    function scheduleSearchRender(value) {
      state.subregionQuery = value;
      if (composingSearch) return;
      if (searchRenderFrame !== null) cancelAnimationFrame(searchRenderFrame);
      searchRenderFrame = requestAnimationFrame(() => {
        searchRenderFrame = null;
        renderDetail();
      });
    }
    searchInput.addEventListener("compositionstart", () => {
      composingSearch = true;
    });
    searchInput.addEventListener("compositionend", (event) => {
      composingSearch = false;
      scheduleSearchRender(event.target.value);
    });
    searchInput.addEventListener("input", (event) => scheduleSearchRender(event.target.value));
    const sortSelect = document.getElementById("subregion-sort");
    const directionSelect = document.getElementById("subregion-order");
    sortSelect.addEventListener("change", (event) => {
      state.detailSortKey = event.target.value;
      state.detailSortDirection = state.detailSortKey === "name" ? "asc" : "desc";
      directionSelect.innerHTML = detailDirectionOptions();
      renderDetail();
    });
    directionSelect.addEventListener("change", (event) => {
      state.detailSortDirection = event.target.value;
      renderDetail();
    });
    document.querySelectorAll("[data-map-metric]").forEach((button) => button.addEventListener("click", () => {
      state.nationalMetric = button.dataset.mapMetric;
      document.querySelectorAll("[data-map-metric]").forEach((candidate) => {
        const active = candidate === button;
        candidate.classList.toggle("is-active", active);
        candidate.setAttribute("aria-pressed", String(active));
      });
      document.getElementById("map-scale-title").textContent = state.nationalMetric === "records" ? "등록건수" : "설비용량";
      renderList();
      renderDetail();
      mapController?.metric(state.nationalMetric);
    }));
    bindRows();
    renderDetail();
    mapController = drawMap(inventory, boundaries, selectRegion);
  }

  function drawMap(inventory, boundaries, onSelect) {
    const node = document.getElementById("map");
    if (!node) return null;
    if (typeof L === "undefined") {
      node.innerHTML = '<p class="map-empty">지도 모듈을 불러오지 못했습니다. 시도 목록에서 지역을 선택해 주세요.</p>';
      return null;
    }
    if (!boundaries) {
      node.innerHTML = '<p class="map-empty">지도를 불러오지 못했습니다. 시도 목록에서 지역을 선택해 주세요.</p>';
      return null;
    }
    const names = {
      KR11: "서울특별시", KR26: "부산광역시", KR27: "대구광역시", KR28: "인천광역시",
      KR29: "전남광주통합특별시", KR30: "대전광역시", KR31: "울산광역시", KR41: "경기도",
      KR42: "강원특별자치도", KR43: "충청북도", KR44: "충청남도", KR45: "전북특별자치도",
      KR46: "전남광주통합특별시", KR47: "경상북도", KR48: "경상남도", KR49: "제주특별자치도", KR50: "세종특별자치시",
    };
    const byName = new Map((inventory.regions || []).map((row) => [row.region, row]));
    const colors = ["#eff6f4", "#d8eae5", "#afd2c8", "#72ae9e", "#327c69"];
    let metric = state.nationalMetric;
    let selected = state.selectedRegion;
    const map = L.map(node, {
      zoomControl: false,
      zoomSnap: .1,
      zoomDelta: .25,
      preferCanvas: true,
      scrollWheelZoom: false,
      doubleClickZoom: false,
      attributionControl: false,
      minZoom: 5,
      maxZoom: 8,
    });
    L.control.zoom({ position: "bottomright" }).addTo(map);
    const attribution = L.control.attribution({ position: "bottomright", prefix: false }).addTo(map);
    attribution.addAttribution('&copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener">OpenStreetMap contributors</a>');
    attribution.addAttribution('경계 &copy; <a href="https://www.data.go.kr/data/15129688/fileData.do" target="_blank" rel="noopener">국가데이터처 SGIS</a>');
    L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
      opacity: .62,
      maxZoom: 19,
      attribution: "",
    }).addTo(map);
    const key = () => metric === "records" ? "generator_records" : "capacity_mw";
    const maximum = () => Math.max(...(inventory.regions || []).map((row) => Number(row[key()]) || 0), 1);
    const color = (value) => {
      const share = Number(value || 0) / maximum();
      return share >= .72 ? colors[4] : share >= .42 ? colors[3] : share >= .20 ? colors[2] : share >= .08 ? colors[1] : colors[0];
    };
    const featureName = (feature) => names[feature.properties.id] || feature.properties.name;
    const style = (feature) => {
      const name = featureName(feature);
      const active = name === selected;
      return { color: active ? "#0a5b48" : "#698c81", weight: active ? 2.4 : 1.1, fillColor: color(byName.get(name)?.[key()]), fillOpacity: active ? .68 : .5 };
    };
    const layer = L.geoJSON(boundaries, {
      style,
      onEachFeature: (feature, shape) => {
        const name = featureName(feature);
        const row = byName.get(name) || {};
        shape.bindTooltip(() => `<strong>${esc(name)}</strong><div class="map-tooltip-metrics"><span><small>설비 등록</small><b>${esc(`${int.format(Number(row.generator_records) || 0)}건`)}</b></span><span><small>설비용량</small><b>${esc(`${number(row.capacity_mw, 2)} MW`)}</b></span></div>`, { sticky: true, direction: "top", className: "province-hover-tooltip", opacity: 1 });
        shape.on({
          mouseover: (event) => event.target.setStyle({ color: "#0a5b48", weight: 2, fillOpacity: .64 }),
          mouseout: () => layer.setStyle(style),
          click: () => onSelect(name),
        });
      },
    }).addTo(map);
    map.fitBounds(layer.getBounds(), { padding: [4, 4], animate: false });
    map.setMaxBounds(layer.getBounds().pad(.08));
    return {
      metric(next) { metric = next; layer.setStyle(style); },
      select(region) { selected = region; layer.setStyle(style); },
    };
  }

  function emptyAnalysis() {
    return { status: "empty", message: "비교 가능한 평가 결과가 아직 없습니다.", evaluation: {}, models: [], regions: [], plants: [], series: [], anomalies: { prediction_summary: { total: 0, returned_top_events: 0, by_model: [], by_region: [], by_plant: [] }, prediction_signals: [], data_quality_signals: [] } };
  }

  function date(value) {
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? String(value || "").slice(0, 10) : new Intl.DateTimeFormat("ko-KR", { year: "numeric", month: "2-digit", day: "2-digit" }).format(parsed);
  }

  function evaluationContext(evaluation = {}) {
    const parts = [];
    if (evaluation.from && evaluation.to) parts.push(`${date(evaluation.from)} ~ ${date(evaluation.to)}`);
    if (finite(evaluation.horizon_hours) && Number(evaluation.horizon_hours) > 0) parts.push(`${number(evaluation.horizon_hours)}시간 예측`);
    if (finite(evaluation.common_samples) && Number(evaluation.common_samples) > 0) parts.push(`공통 평가 ${int.format(Number(evaluation.common_samples))}건`);
    return parts.join(" · ");
  }

  function forecastModels(analysis) {
    const points = normalizeSeries(analysis);
    return (analysis.models || []).filter((model) =>
      points.some((point) => finite(point.predictions?.[model.id]))
    );
  }

  function forecastModelOptions(analysis) {
    return forecastModels(analysis).map((model) =>
      `<option value="${esc(model.id)}"${model.id === state.forecastModelId ? " selected" : ""}>${esc(model.label)}</option>`
    ).join("");
  }

  function regionsForForecast(analysis) {
    return [...new Set(normalizeSeries(analysis).map((row) => row.region).filter(Boolean))]
      .sort((a, b) => displayRegion(a).localeCompare(displayRegion(b), "ko"));
  }

  function forecastRegionOptions(analysis) {
    return '<option value="all">전국</option>' + regionsForForecast(analysis).map((region) =>
      `<option value="${esc(region)}"${region === state.region ? " selected" : ""}>${esc(displayRegion(region))}</option>`
    ).join("");
  }

  function plantsForForecast(analysis) {
    const result = new Map();
    normalizeSeries(analysis).forEach((row) => {
      if (row.plant_id && (state.region === "all" || row.region === state.region)) {
        result.set(String(row.plant_id), row.plant || String(row.plant_id));
      }
    });
    return [...result.entries()]
      .map(([plant_id, plant]) => ({ plant_id, plant }))
      .sort((a, b) => a.plant.localeCompare(b.plant, "ko"));
  }

  function forecastPlantOptions(analysis) {
    return '<option value="all">발전소 선택</option>' + plantsForForecast(analysis).map((row) =>
      `<option value="${esc(row.plant_id)}"${row.plant_id === state.plantId ? " selected" : ""}>${esc(row.plant)}</option>`
    ).join("");
  }

  function forecastPage(analysis) {
    const models = forecastModels(analysis);
    if (!models.some((model) => model.id === state.forecastModelId)) {
      state.forecastModelId = models[0]?.id || "all";
    }
    const header = intro(
      "테스트 구간 발전량 예측 결과",
      "정식 평가에 사용된 과거 Test 구간의 실제 발전량과 모델 예측값을 발전소별로 확인합니다.",
      evaluationContext(analysis.evaluation),
    ) +
      `<p class="forecast-note">현재·미래 운영 예측이 아니라 저장된 Test 평가 결과입니다. 예보 발행시각이 보존된 기상예보 입력이 연결되기 전에는 미래 발전량을 표시하지 않습니다.</p>` +
      (analysis.status !== "ready" && analysis.message ? `<p class="status-message">${esc(analysis.message)}</p>` : "");
    if (!models.length) {
      return header + `<section id="forecast-content" class="forecast-content">${emptyState("예측 결과가 없습니다", analysis.message || "정식 모델의 Test 예측이 완료되면 발전소별 시계열을 표시합니다.")}</section>`;
    }
    return header + `<section class="analysis-toolbar forecast-toolbar" aria-label="발전량 예측 조건">
        <label class="field-label" for="forecast-region"><span>지역</span><select id="forecast-region">${forecastRegionOptions(analysis)}</select></label>
        <label class="field-label" for="forecast-plant"><span>발전소</span><select id="forecast-plant">${forecastPlantOptions(analysis)}</select></label>
        <label class="field-label" for="forecast-model"><span>예측 모델</span><select id="forecast-model">${forecastModelOptions(analysis)}</select></label>
      </section>
      <section id="forecast-content" class="forecast-content"></section>`;
  }

  function setupForecast(analysis) {
    const content = document.getElementById("forecast-content");
    const regionSelect = document.getElementById("forecast-region");
    if (!regionSelect) return;
    const render = () => { content.innerHTML = forecastView(analysis); };
    regionSelect.addEventListener("change", (event) => {
      state.region = event.target.value;
      state.plantId = "all";
      document.getElementById("forecast-plant").innerHTML = forecastPlantOptions(analysis);
      render();
    });
    document.getElementById("forecast-plant").addEventListener("change", (event) => {
      state.plantId = event.target.value;
      render();
    });
    document.getElementById("forecast-model").addEventListener("change", (event) => {
      state.forecastModelId = event.target.value;
      render();
    });
    render();
  }

  function forecastView(analysis) {
    const models = forecastModels(analysis);
    const model = models.find((candidate) => candidate.id === state.forecastModelId);
    if (!model) {
      return emptyState("예측 결과가 없습니다", analysis.message || "정식 모델의 Test 예측이 완료되면 발전소별 시계열을 표시합니다.");
    }
    if (state.plantId === "all") {
      return `<section class="surface forecast-series">${emptyState("발전소를 선택해 주세요", "상단에서 지역과 발전소를 선택하면 실제 발전량과 선택 모델의 예측값을 표시합니다.")}</section>`;
    }
    const selected = plantsForForecast(analysis).find((row) => row.plant_id === state.plantId);
    const label = selected?.plant || state.plantId;
    const points = normalizeSeries(analysis).filter((row) =>
      (state.region === "all" || row.region === state.region) &&
      row.plant_id === state.plantId &&
      finite(row.predictions?.[model.id])
    );
    if (!points.length) {
      return emptyState("선택 조건의 예측 결과가 없습니다", "다른 발전소나 모델을 선택해 확인해 주세요.");
    }
    const horizon = finite(analysis.evaluation?.horizon_hours) ? number(analysis.evaluation.horizon_hours) : "-";
    return stats([
      ["선택 모델", model.label || model.id],
      ["표시 시점", int.format(points.length), "건"],
      ["예측 범위", horizon, finite(analysis.evaluation?.horizon_hours) ? "시간" : ""],
    ]) +
      `<section class="surface forecast-series"><div class="section-head"><div><h2>${esc(label)} 발전량 예측</h2><p>Test 구간에서 가장 최근의 연속 최대 168시간을 실제 발전량과 ${esc(model.label || model.id)} 예측값으로 표시합니다.</p></div></div>${lineChart(points, [model], label)}</section>`;
  }

  function analysisPage(analysis) {
    return intro("태양광 모델 성능 비교·분석", "동일한 평가 표본에서 모델 정확도를 비교하고 지역·발전소별 오차와 이상 신호를 살펴봅니다.", evaluationContext(analysis.evaluation)) +
      (analysis.status !== "ready" && analysis.message ? `<p class="status-message">${esc(analysis.message)}</p>` : "") +
      `<section class="analysis-toolbar" aria-label="분석 조건">
        <label class="field-label" for="analysis-region"><span>지역</span><select id="analysis-region">${analysisRegionOptions(analysis)}</select></label>
        <label class="field-label" for="analysis-plant"><span>발전소</span><select id="analysis-plant">${analysisPlantOptions(analysis)}</select></label>
      </section>
      <div class="analysis-tabs" role="tablist" aria-label="모델 성능 분석 보기">
        <button type="button" role="tab" id="tab-comparison" aria-controls="analysis-content" aria-selected="true" data-tab="comparison">모델 비교</button>
        <button type="button" role="tab" id="tab-performance" aria-controls="analysis-content" aria-selected="false" data-tab="performance">지역·발전소 성능</button>
        <button type="button" role="tab" id="tab-anomalies" aria-controls="analysis-content" aria-selected="false" data-tab="anomalies">이상·오차 분석</button>
      </div>
      <section id="analysis-content" class="analysis-content" role="tabpanel" aria-labelledby="tab-comparison"></section>`;
  }

  function predictionEvents(analysis) {
    const source = analysis.anomalies || {};
    return source.prediction_signals || [];
  }

  function qualitySignals(analysis) {
    return analysis.anomalies?.data_quality_signals || [];
  }

  function normalizedSummaryRows(value, keyName) {
    const rows = Array.isArray(value) ? value : Object.entries(value || {}).map(([key, row]) => (
      row && typeof row === "object"
        ? { [keyName]: key, ...row }
        : { [keyName]: key, count: Number(row) || 0 }
    ));
    return rows.map((row) => ({
      ...row,
      count: Number(row.count ?? row.signals) || 0,
    }));
  }

  function predictionSummary(analysis) {
    const raw = analysis.anomalies?.prediction_summary || {};
    const events = predictionEvents(analysis);
    return {
      total: finite(raw.total) ? Number(raw.total) : events.length,
      returned_top_events: finite(raw.returned_top_events) ? Number(raw.returned_top_events) : events.length,
      by_model: normalizedSummaryRows(raw.by_model, "model"),
      by_region: normalizedSummaryRows(raw.by_region, "region"),
      by_plant: normalizedSummaryRows(raw.by_plant, "plant_id"),
    };
  }

  function regionsForAnalysis(analysis) {
    const result = new Set();
    (analysis.regions || []).forEach((row) => row.region && result.add(row.region));
    (analysis.plants || []).forEach((row) => row.region && result.add(row.region));
    predictionEvents(analysis).forEach((row) => row.region && result.add(row.region));
    qualitySignals(analysis).forEach((row) => row.region && result.add(row.region));
    predictionSummary(analysis).by_region.forEach((row) => row.region && result.add(row.region));
    predictionSummary(analysis).by_plant.forEach((row) => row.region && result.add(row.region));
    return [...result].sort((a, b) => displayRegion(a).localeCompare(displayRegion(b), "ko"));
  }

  function plantsForAnalysis(analysis) {
    const result = new Map();
    const sources = [
      ...(analysis.plants || []),
      ...(analysis.series || []),
      ...predictionEvents(analysis),
      ...qualitySignals(analysis),
      ...predictionSummary(analysis).by_plant,
    ];
    sources.forEach((row) => {
      if (row.plant_id && (state.region === "all" || row.region === state.region)) {
        result.set(String(row.plant_id), row.plant || String(row.plant_id));
      }
    });
    return [...result.entries()]
      .map(([plant_id, plant]) => ({ plant_id, plant }))
      .sort((a, b) => a.plant.localeCompare(b.plant, "ko"));
  }

  function analysisRegionOptions(analysis) {
    return '<option value="all">전국</option>' + regionsForAnalysis(analysis).map((region) => `<option value="${esc(region)}"${region === state.region ? " selected" : ""}>${esc(displayRegion(region))}</option>`).join("");
  }

  function analysisPlantOptions(analysis) {
    return '<option value="all">전체 발전소</option>' + plantsForAnalysis(analysis).map((row) => `<option value="${esc(row.plant_id)}"${row.plant_id === state.plantId ? " selected" : ""}>${esc(row.plant)}</option>`).join("");
  }

  function setupAnalysis(analysis) {
    const content = document.getElementById("analysis-content");
    const toolbar = document.querySelector(".analysis-toolbar");
    const tabs = [...document.querySelectorAll("[data-tab]")];
    const render = () => {
      toolbar.hidden = state.tab === "comparison";
      content.innerHTML = state.tab === "performance" ? performanceView(analysis) : state.tab === "anomalies" ? anomalyView(analysis) : comparisonView(analysis);
      content.setAttribute("aria-labelledby", `tab-${state.tab}`);
      bindAnalysisLinks(analysis, render);
    };
    tabs.forEach((button, index) => {
      button.addEventListener("click", () => {
        state.tab = button.dataset.tab;
        tabs.forEach((candidate) => candidate.setAttribute("aria-selected", String(candidate === button)));
        render();
      });
      button.addEventListener("keydown", (event) => {
        if (!["ArrowLeft", "ArrowRight"].includes(event.key)) return;
        event.preventDefault();
        const next = tabs[(index + (event.key === "ArrowRight" ? 1 : -1) + tabs.length) % tabs.length];
        next.focus();
        next.click();
      });
    });
    document.getElementById("analysis-region").addEventListener("change", (event) => {
      state.region = event.target.value;
      state.plantId = "all";
      document.getElementById("analysis-plant").innerHTML = analysisPlantOptions(analysis);
      render();
    });
    document.getElementById("analysis-plant").addEventListener("change", (event) => { state.plantId = event.target.value; render(); });
    render();
  }

  function bindAnalysisLinks(analysis, render) {
    document.querySelectorAll("[data-analysis-region]").forEach((button) => button.addEventListener("click", () => {
      state.region = button.dataset.analysisRegion;
      state.plantId = "all";
      document.getElementById("analysis-region").value = state.region;
      document.getElementById("analysis-plant").innerHTML = analysisPlantOptions(analysis);
      render();
    }));
    document.querySelectorAll("[data-analysis-plant-id]").forEach((button) => button.addEventListener("click", () => {
      state.plantId = button.dataset.analysisPlantId;
      const select = document.getElementById("analysis-plant");
      if ([...select.options].some((option) => option.value === state.plantId)) select.value = state.plantId;
      render();
    }));
  }

  function metricValue(row, key) {
    const value = row?.metrics?.[key];
    return finite(value) ? Number(value) : null;
  }

  function metricText(key, value) {
    const definition = metrics[key] || metrics.mae;
    return finite(value) ? `${number(value, definition.digits)}${definition.unit ? ` ${definition.unit}` : ""}` : "-";
  }

  function metricHint(key) {
    const direction = metrics[key].better === "higher" ? "높을수록 좋음" : "낮을수록 좋음";
    return key === "nmae_capacity" ? `${direction} · 설비용량 확인 표본 기준` : direction;
  }

  function comparisonView(analysis) {
    const allModels = analysis.models || [];
    const models = allModels.filter((model) => Object.keys(metrics).some((key) => metricValue(model, key) !== null));
    const comparable = models.filter((model) => model.comparable !== false);
    if (!models.length) return emptyState("모델 평가 결과가 없습니다", analysis.message || "모델 평가가 완료되면 네 가지 성능 지표가 함께 표시됩니다.");
    if (comparable.length < 2) {
      return emptyState("모델 간 비교 결과가 없습니다", analysis.message || "동일한 테스트 표본의 모델이 두 개 이상 확보되면 네 가지 지표 비교가 활성화됩니다.") +
        `<section class="surface analysis-section"><div class="section-head"><div><h2>확인 가능한 개별 결과</h2><p>평가 구간이 같지 않아 모델 간 순위나 막대 비교에는 사용하지 않습니다.</p></div></div>${modelTable(models)}</section>`;
    }
    const commonSamples = finite(analysis.evaluation?.common_samples) ? Number(analysis.evaluation.common_samples) : 0;
    const horizon = finite(analysis.evaluation?.horizon_hours) ? number(analysis.evaluation.horizon_hours) : "-";
    let result = stats([
      ["비교 모델", int.format(comparable.length), "개"],
      ["공통 평가 표본", int.format(commonSamples), "건"],
      ["예측 범위", horizon, finite(analysis.evaluation?.horizon_hours) ? "시간" : ""],
    ]);
    result += `<section class="metric-overview-grid" aria-label="모델별 네 가지 성능 지표">${Object.keys(metrics).map((key) => metricPanel(comparable, key)).join("")}</section>`;
    result += `<section class="surface analysis-section"><div class="section-head"><div><h2>모델별 전체 평가 지표</h2><p>NMAE, MAE, RMSE, R²를 같은 평가 구간에서 함께 확인합니다.</p></div></div>${modelTable(models)}</section>`;
    return result;
  }

  function metricPanel(models, key) {
    const available = models.filter((model) => metricValue(model, key) !== null);
    return `<article class="surface metric-panel"><div class="metric-panel-head"><div><strong>${esc(metrics[key].short)}</strong><span>${esc(metrics[key].label)}</span></div><small>${esc(metricHint(key))}</small></div>${available.length ? modelBars(available, key) : emptyState(`${metrics[key].short} 결과 없음`, "이 지표를 계산할 수 있는 평가 표본이 없습니다.")}</article>`;
  }

  function modelBars(models, key) {
    const values = models.map((model) => metricValue(model, key));
    const maximum = Math.max(...values, 0);
    const minimum = Math.min(...values, 0);
    return `<div class="metric-bars">${models.map((model, index) => {
      const value = metricValue(model, key);
      const width = metrics[key].better === "higher" ? (maximum === minimum ? 100 : (value - minimum) / (maximum - minimum) * 100) : (maximum ? value / maximum * 100 : 0);
      return `<div class="metric-bar-row"><div><strong>${esc(model.label)}</strong><span>${esc(metricText(key, value))}</span></div><div class="metric-track"><i class="series-${index % 4}" style="width:${Math.max(3, width).toFixed(2)}%"></i></div></div>`;
    }).join("")}</div>`;
  }

  function modelTable(models) {
    return `<div class="table-responsive"><table class="data-table"><thead><tr><th>모델</th><th class="numeric">NMAE</th><th class="numeric">용량 적용률</th><th class="numeric">MAE</th><th class="numeric">RMSE</th><th class="numeric">R²</th><th class="numeric">평가 표본</th></tr></thead><tbody>${models.map((model) => `<tr><td><strong>${esc(model.label)}</strong>${model.comparable === false ? '<span class="status-tag neutral">비교 제외</span>' : ""}</td><td class="numeric">${esc(metricText("nmae_capacity", metricValue(model, "nmae_capacity")))}</td><td class="numeric">${finite(model.metrics?.capacity_coverage) ? `${number(Number(model.metrics.capacity_coverage) * 100, 1)}%` : "-"}</td><td class="numeric">${esc(metricText("mae", metricValue(model, "mae")))}</td><td class="numeric">${esc(metricText("rmse", metricValue(model, "rmse")))}</td><td class="numeric">${esc(metricText("r2", metricValue(model, "r2")))}</td><td class="numeric">${finite(model.metrics?.n_samples) ? int.format(Number(model.metrics.n_samples)) : "-"}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function normalizeSeries(analysis) {
    const grouped = new Map();
    (analysis.series || []).forEach((row) => {
      if (!row.timestamp || !row.plant_id) return;
      const key = [row.timestamp, row.plant_id].join("|");
      const point = grouped.get(key) || { timestamp: row.timestamp, plant_id: String(row.plant_id), region: row.region, plant: row.plant, y_true: finite(row.y_true) ? Number(row.y_true) : null, predictions: {} };
      if (finite(row.y_true)) point.y_true = Number(row.y_true);
      if (row.predictions && typeof row.predictions === "object") Object.entries(row.predictions).forEach(([model, value]) => { if (finite(value)) point.predictions[model] = Number(value); });
      (analysis.models || []).forEach((model) => {
        const id = String(model.id || "");
        const lower = id.toLowerCase();
        const candidates = [row[`${id}_pred`], lower.includes("xg") ? row.xgb_pred : null, lower.includes("cnn") ? row.cnn_pred : null, lower.includes("hybrid") ? row.hybrid_pred : null, row.model === id ? row.y_pred : null];
        const value = candidates.find(finite);
        if (value !== undefined) point.predictions[id] = Number(value);
      });
      grouped.set(key, point);
    });
    return [...grouped.values()].sort((a, b) => String(a.timestamp).localeCompare(String(b.timestamp)));
  }

  function lineChart(source, models, plantLabel) {
    const points = source.length <= 168 ? source : Array.from({ length: 168 }, (_, index) => source[Math.round(index * (source.length - 1) / 167)]);
    const visible = models.filter((model) => points.some((point) => finite(point.predictions[model.id])));
    const values = points.flatMap((point) => [point.y_true, ...visible.map((model) => point.predictions[model.id])]).filter(finite).map(Number);
    if (!values.length) return emptyState("표시할 값이 없습니다", "실제값과 예측값을 확인할 수 없습니다.");
    const width = 920, height = 300, pad = { left: 58, right: 20, top: 20, bottom: 42 };
    let minimum = Math.min(...values), maximum = Math.max(...values);
    if (minimum === maximum) { minimum -= 1; maximum += 1; }
    const x = (index) => pad.left + index / Math.max(1, points.length - 1) * (width - pad.left - pad.right);
    const y = (value) => pad.top + (maximum - value) / (maximum - minimum) * (height - pad.top - pad.bottom);
    const path = (getter) => {
      let open = false;
      return points.map((point, index) => {
        const value = getter(point);
        if (!finite(value)) { open = false; return ""; }
        const command = open ? "L" : "M";
        open = true;
        return `${command}${x(index).toFixed(2)},${y(Number(value)).toFixed(2)}`;
      }).join(" ");
    };
    const ticks = [minimum, minimum + (maximum - minimum) / 2, maximum];
    const timeIndexes = [...new Set([0, Math.floor((points.length - 1) / 2), points.length - 1])];
    return `<div class="chart-legend"><span><i class="actual"></i>실제 발전량</span>${visible.map((model, index) => `<span><i class="series-${index % 4}"></i>${esc(model.label)}</span>`).join("")}</div>
      <svg class="line-chart" viewBox="0 0 ${width} ${height}" role="img" aria-labelledby="series-chart-title series-chart-desc">
        <title id="series-chart-title">${esc(plantLabel)} 실제 발전량과 예측값</title><desc id="series-chart-desc">실제 발전량과 모델별 예측값을 비교한 선 그래프입니다.</desc>
        ${ticks.map((tick) => `<line class="chart-gridline" x1="${pad.left}" x2="${width - pad.right}" y1="${y(tick)}" y2="${y(tick)}"></line><text class="chart-axis-label" x="${pad.left - 10}" y="${y(tick) + 4}" text-anchor="end">${esc(number(tick, 2))}</text>`).join("")}
        ${timeIndexes.map((index) => `<text class="chart-axis-label" x="${x(index)}" y="${height - 14}" text-anchor="${index === 0 ? "start" : index === points.length - 1 ? "end" : "middle"}">${esc(shortTime(points[index].timestamp))}</text>`).join("")}
        <text class="chart-axis-title" x="14" y="${height / 2}" transform="rotate(-90 14 ${height / 2})">발전량(MWh)</text>
        <path class="chart-line actual" d="${path((point) => point.y_true)}"></path>${visible.map((model, index) => `<path class="chart-line series-${index % 4}" d="${path((point) => point.predictions[model.id])}"></path>`).join("")}
      </svg>`;
  }

  function shortTime(value) {
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? String(value).slice(0, 16) : new Intl.DateTimeFormat("ko-KR", { month: "2-digit", day: "2-digit", hour: "2-digit" }).format(parsed);
  }

  function performanceView(analysis) {
    const regionRows = analysis.regions || [];
    const plantRows = analysis.plants || [];
    if (!regionRows.length && !plantRows.length) {
      return emptyState("지역·발전소 성능 결과가 없습니다", analysis.message || "세부 평가가 완료되면 지역과 발전소별 성능이 표시됩니다.");
    }
    return `<section class="surface analysis-section">
        <div class="section-head"><div><h2>지역별 네 가지 성능 지표</h2><p>모델별 NMAE, MAE, RMSE, R²를 함께 보고 지역을 선택해 발전소 결과로 이어갑니다.</p></div></div>
        ${regionMatrix(analysis)}
      </section>
      <section class="surface analysis-section">
        <div class="section-head"><div><h2>${state.region === "all" ? "발전소별 성능" : `${esc(displayRegion(state.region))} 발전소별 성능`}</h2><p>발전소마다 네 가지 지표와 평가 표본을 함께 확인합니다.</p></div></div>
        ${plantTable(analysis)}
      </section>`;
  }

  function regionMatrix(analysis) {
    const models = analysis.models || [];
    const grouped = new Map();
    (analysis.regions || []).forEach((row) => {
      if (!row.region) return;
      const values = grouped.get(row.region) || new Map();
      values.set(row.model, row.metrics || {});
      grouped.set(row.region, values);
    });
    if (!grouped.size) return emptyState("지역별 결과 없음", "지역 단위 평가 결과가 아직 없습니다.");
    const ordered = [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0], "ko"));
    return `<div class="table-responsive"><table class="data-table region-matrix">
      <thead><tr><th>지역</th>${models.map((model) => `<th>${esc(model.label)}</th>`).join("")}</tr></thead>
      <tbody>${ordered.map(([region, values]) => `<tr class="${region === state.region ? "is-selected" : ""}">
        <td><button type="button" class="text-button" data-analysis-region="${esc(region)}">${esc(displayRegion(region))}</button></td>
        ${models.map((model) => `<td>${metricQuartet(values.get(model.id))}</td>`).join("")}
      </tr>`).join("")}</tbody>
    </table></div>`;
  }

  function metricQuartet(values = {}) {
    return `<div class="metric-quartet">${Object.keys(metrics).map((key) => `<span><small>${esc(metrics[key].short)}</small><strong>${esc(metricText(key, values?.[key]))}</strong></span>`).join("")}</div>`;
  }

  function plantTable(analysis) {
    if (state.region === "all") return emptyState("지역을 선택해 주세요", "위 지역별 비교에서 확인할 지역을 선택하면 발전소별 결과가 표시됩니다.");
    let rows = (analysis.plants || []).filter((row) => row.region === state.region);
    if (state.plantId !== "all") rows = rows.filter((row) => String(row.plant_id) === state.plantId);
    rows.sort((a, b) => String(a.plant || "").localeCompare(String(b.plant || ""), "ko") || String(a.model || "").localeCompare(String(b.model || "")));
    if (!rows.length) return emptyState("발전소별 결과 없음", "선택한 지역의 발전소별 평가 결과가 아직 없습니다.");
    return `<div class="table-responsive"><table class="data-table plant-performance-table">
      <thead><tr><th>발전소</th><th>모델</th><th class="numeric">NMAE</th><th class="numeric">MAE</th><th class="numeric">RMSE</th><th class="numeric">R²</th><th class="numeric">평가 표본</th></tr></thead>
      <tbody>${rows.map((row) => {
        const model = (analysis.models || []).find((candidate) => candidate.id === row.model);
        return `<tr><td><button type="button" class="text-button" data-analysis-plant-id="${esc(row.plant_id)}">${esc(row.plant)}</button></td><td>${esc(model?.label || row.model)}</td><td class="numeric">${esc(metricText("nmae_capacity", metricValue(row, "nmae_capacity")))}</td><td class="numeric">${esc(metricText("mae", metricValue(row, "mae")))}</td><td class="numeric">${esc(metricText("rmse", metricValue(row, "rmse")))}</td><td class="numeric">${esc(metricText("r2", metricValue(row, "r2")))}</td><td class="numeric">${finite(row.metrics?.n_samples) ? int.format(Number(row.metrics.n_samples)) : "-"}</td></tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function anomalyView(analysis) {
    const summary = predictionSummary(analysis);
    const events = filteredPredictionEvents(analysis);
    const quality = filteredQualitySignals(analysis);
    const predictionCount = selectedPredictionCount(analysis);
    const total = predictionCount + quality.length;
    if (!total) {
      return emptyState(
        "선택 조건의 이상 신호가 없습니다",
        summary.total + qualitySignals(analysis).length ? "다른 지역이나 발전소를 선택해 확인해 주세요." : "예측 오차와 데이터 패턴 분석이 완료되면 검토할 신호가 표시됩니다.",
      );
    }
    return `<p class="interpretation-note">이상 신호는 고장 확정이 아니라 예측 오차 또는 데이터 패턴을 추가로 확인할 대상입니다.</p>
      ${stats([["전체 이상 신호", int.format(total), "건"], ["예측 오차 신호", int.format(predictionCount), "건"], ["데이터 패턴 신호", int.format(quality.length), "건"]])}
      <section class="analysis-grid anomaly-grid">
        <article class="surface analysis-section"><div class="section-head"><div><h2>신호 유형</h2><p>대표 이벤트가 아닌 전체 집계 기준입니다.</p></div></div>${typeBars(combinedTypeCounts(predictionCount, quality))}</article>
        <article class="surface analysis-section"><div class="section-head"><div><h2>지역별 이상 신호</h2><p>예측 오차와 데이터 패턴 신호를 합산합니다.</p></div></div>${signalRegionBars(combinedRegionCounts(analysis, quality))}</article>
      </section>
      <section class="surface analysis-section"><div class="section-head"><div><h2>대표 예측 오차 이벤트</h2><p>${esc(representativeEventCaption(summary, predictionCount, events.length))}</p></div></div>${predictionEventTable(events, predictionCount)}</section>
      <section class="surface analysis-section"><div class="section-head"><div><h2>데이터 패턴 확인 대상</h2><p>선택 조건에 해당하는 품질 신호를 모두 표시합니다.</p></div></div>${qualitySignalTable(quality)}</section>`;
  }

  function signalTypes(signal) {
    const source = Array.isArray(signal.signal_types) ? signal.signal_types : signal.signal_types ? [signal.signal_types] : ["pattern_review"];
    const labels = {
      missing_weather: "기상 데이터 공백", missing_weather_rate: "기상 데이터 공백",
      daylight_zero: "주간 0발전", daylight_zero_rate: "주간 0발전",
      flatline: "동일값 지속", positive_flatline: "양수 동일값 지속",
      temporal_pattern: "시간대 패턴 차이", temporal_profile: "시간대 패턴 차이",
      peer_pattern: "지역 내 패턴 차이", large_residual: "예측 오차 임계 초과",
      capacity_exceeded: "설비용량 초과", pattern_review: "데이터 패턴 검토",
    };
    return source.map((value) => labels[String(value)] || String(value).replaceAll("_", " "));
  }

  function filteredPredictionEvents(analysis) {
    return predictionEvents(analysis)
      .filter((row) => state.region === "all" || row.region === state.region)
      .filter((row) => state.plantId === "all" || String(row.plant_id) === state.plantId);
  }

  function filteredQualitySignals(analysis) {
    return qualitySignals(analysis)
      .filter((row) => state.region === "all" || row.region === state.region)
      .filter((row) => state.plantId === "all" || String(row.plant_id) === state.plantId);
  }

  function selectedPredictionCount(analysis) {
    const summary = predictionSummary(analysis);
    if (state.plantId !== "all") {
      const row = summary.by_plant.find((candidate) => String(candidate.plant_id) === state.plantId);
      return row ? Number(row.count) || 0 : filteredPredictionEvents(analysis).length;
    }
    if (state.region !== "all") {
      const row = summary.by_region.find((candidate) => candidate.region === state.region);
      return row ? Number(row.count) || 0 : filteredPredictionEvents(analysis).length;
    }
    return summary.total;
  }

  function combinedTypeCounts(predictionCount, quality) {
    const counts = new Map();
    if (predictionCount) counts.set("예측 오차 임계 초과", predictionCount);
    quality.forEach((signal) => signalTypes(signal).forEach((type) => counts.set(type, (counts.get(type) || 0) + 1)));
    return counts;
  }

  function typeBars(counts) {
    const ordered = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    const maximum = Math.max(...ordered.map((row) => row[1]), 1);
    return `<div class="simple-bars">${ordered.map(([type, count]) => `<div><span>${esc(type)}</span><i><b style="width:${(count / maximum * 100).toFixed(2)}%"></b></i><strong>${int.format(count)}건</strong></div>`).join("")}</div>`;
  }

  function combinedRegionCounts(analysis, quality) {
    const summary = predictionSummary(analysis);
    const counts = new Map();
    if (state.plantId !== "all") {
      const row = summary.by_plant.find((candidate) => String(candidate.plant_id) === state.plantId);
      if (row?.region && Number(row.count)) counts.set(row.region, Number(row.count));
    } else {
      const rows = summary.by_region.length
        ? summary.by_region
        : [...predictionEvents(analysis).reduce((grouped, event) => {
            grouped.set(event.region, (grouped.get(event.region) || 0) + 1);
            return grouped;
          }, new Map()).entries()].map(([region, count]) => ({ region, count }));
      rows.filter((row) => state.region === "all" || row.region === state.region)
        .forEach((row) => counts.set(row.region, Number(row.count) || 0));
    }
    quality.forEach((signal) => counts.set(signal.region, (counts.get(signal.region) || 0) + 1));
    return counts;
  }

  function signalRegionBars(counts) {
    const ordered = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    const maximum = Math.max(...ordered.map((row) => row[1]), 1);
    return `<div class="simple-bars">${ordered.map(([region, count]) => `<div><button type="button" class="text-button" data-analysis-region="${esc(region)}">${esc(displayRegion(region))}</button><i><b style="width:${(count / maximum * 100).toFixed(2)}%"></b></i><strong>${int.format(count)}건</strong></div>`).join("")}</div>`;
  }

  function severity(severityValue) {
    return ({ critical: ["긴급", "danger"], high: ["높음", "danger"], medium: ["주의", "caution"], review: ["검토", "caution"], low: ["낮음", "neutral"] })[String(severityValue || "review").toLowerCase()] || ["검토", "caution"];
  }

  function signalMetricText(signal) {
    const values = [];
    const source = signal.metrics || {};
    const absoluteError = source.absolute_residual ?? source.absolute_error ?? signal.absolute_error;
    const threshold = source.anomaly_threshold ?? source.threshold ?? signal.threshold;
    if (finite(absoluteError)) values.push(`절대오차 ${number(absoluteError, 3)} MWh`);
    if (finite(threshold)) values.push(`판단기준 ${number(threshold, 3)} MWh`);
    if (finite(signal.y_true)) values.push(`실제 ${number(signal.y_true, 3)} MWh`);
    if (finite(signal.y_pred)) values.push(`예측 ${number(signal.y_pred, 3)} MWh`);
    if (finite(source.missing_weather_rate)) values.push(`기상 공백 ${ratio(source.missing_weather_rate)}`);
    if (finite(source.daylight_zero_rate)) values.push(`주간 0발전 ${ratio(source.daylight_zero_rate)}`);
    if (finite(source.hourly_coverage)) values.push(`시간 자료 ${ratio(source.hourly_coverage)}`);
    if (finite(source.peer_pattern_correlation)) values.push(`지역 패턴 유사도 ${number(source.peer_pattern_correlation, 2)}`);
    return values.join(" · ") || "-";
  }

  function representativeEventCaption(summary, selectedTotal, visibleEvents) {
    if (state.region === "all" && state.plantId === "all") {
      return `전체 ${int.format(selectedTotal)}건 중 오차가 큰 대표 ${int.format(summary.returned_top_events)}건을 보관하며, 현재 ${int.format(visibleEvents)}건을 표시합니다.`;
    }
    return `선택 조건의 전체 ${int.format(selectedTotal)}건 중 보관된 대표 이벤트 ${int.format(visibleEvents)}건을 표시합니다.`;
  }

  function thresholdSourceLabel(source) {
    return ({
      plant_capacity_normalized: "발전소별 용량 정규화",
      global_capacity_normalized: "전체 용량 정규화 보완",
      plant_absolute_capacity_missing: "발전소별 절대오차 · 용량 미확인",
      plant_absolute_capacity_fallback: "발전소별 절대오차 보완",
    })[String(source || "")] || "Calibration 잔차 기준";
  }

  function predictionEventTable(events, selectedTotal) {
    if (!events.length) {
      return emptyState(
        "표시할 대표 이벤트가 없습니다",
        selectedTotal ? "전체 집계에는 신호가 있지만 선택 조건에 해당하는 대표 이벤트는 보관되지 않았습니다." : "선택 조건의 예측 오차 신호가 없습니다.",
      );
    }
    return `<div class="table-responsive"><table class="data-table prediction-event-table">
      <thead><tr><th>모델</th><th>시각</th><th>지역</th><th>발전소</th><th class="numeric">실제(MWh)</th><th class="numeric">예측(MWh)</th><th class="numeric">절대오차</th><th class="numeric">판단기준(MWh)</th><th class="numeric">기준 대비</th><th>판단 방식</th></tr></thead>
      <tbody>${events.map((event) => `<tr><td>${esc(event.model_label || event.model)}</td><td>${esc(shortTime(event.timestamp))}</td><td>${esc(displayRegion(event.region))}</td><td>${event.plant_id ? `<button type="button" class="text-button" data-analysis-plant-id="${esc(event.plant_id)}">${esc(event.plant || event.plant_id)}</button>` : esc(event.plant || "-")}</td><td class="numeric">${number(event.y_true, 3)}</td><td class="numeric">${number(event.y_pred, 3)}</td><td class="numeric">${number(event.absolute_error, 3)}</td><td class="numeric">${number(event.threshold, 3)}</td><td class="numeric">${finite(event.exceedance_ratio) ? `${number(event.exceedance_ratio, 2)}배` : "-"}</td><td>${esc(thresholdSourceLabel(event.threshold_source))}</td></tr>`).join("")}</tbody>
    </table></div>`;
  }

  function qualitySignalTable(signals) {
    if (!signals.length) return emptyState("데이터 패턴 신호가 없습니다", "선택 조건에서 확인할 데이터 패턴 신호가 없습니다.");
    const rank = { critical: 4, high: 3, medium: 2, review: 2, low: 1 };
    const ordered = [...signals].sort((a, b) => (rank[String(b.severity).toLowerCase()] || 2) - (rank[String(a.severity).toLowerCase()] || 2));
    return `<div class="table-responsive"><table class="data-table quality-signal-table">
      <thead><tr><th>우선도</th><th>지역</th><th>발전소</th><th>신호</th><th>관련 지표</th></tr></thead>
      <tbody>${ordered.map((signal) => {
        const [label, kind] = severity(signal.severity);
        const plant = signal.plant_id ? `<button type="button" class="text-button" data-analysis-plant-id="${esc(signal.plant_id)}">${esc(signal.plant || signal.plant_id)}</button>` : esc(signal.plant || "-");
        return `<tr><td><span class="status-tag ${kind}">${esc(label)}</span></td><td>${esc(displayRegion(signal.region))}</td><td>${plant}</td><td>${esc(signal.summary || signalTypes(signal).join(", "))}</td><td>${esc(signalMetricText(signal))}</td></tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }
})();
