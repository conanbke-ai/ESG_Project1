// Coverage view helpers; composed into one private browser scope.
  function renderCoveragePage(inventory = {}) {
    const summary = inventory.summary || {};
    const regions = inventory.regions || [];
    const leader = [...regions].sort((a, b) => Number(b.capacity_mw) - Number(a.capacity_mw))[0];
    state.selectedRegion = leader?.region || regions[0]?.region || null;
    return renderPageIntro("전국 태양광 발전설비 현황", "전국 설비 규모를 한눈에 비교하고 원하는 세부지역을 바로 찾아볼 수 있습니다.") +
      renderSummaryStats([
        ["전국 설비용량", formatNumber(summary.total_capacity_mw, 2), "MW"],
        ["설비 등록", int.format(Number(summary.generator_records) || 0), "건"],
        ["설비용량 1위", leader ? displayRegion(leader.region) : "-", leader ? `${formatNumber(leader.capacity_mw, 2)} MW` : ""],
      ]) +
      `<section class="national-layout">
        <article class="surface map-surface">
          <div class="section-head">
            <div><h2>시도별 설비 분포</h2><p>지도나 목록에서 지역을 선택해 설비 규모를 비교해 보세요.</p></div>
            <div class="national-controls">
              <label class="field-label" for="province-select"><span>시도</span><select id="province-select">${renderRegionOptions(regions)}</select></label>
              <div class="segmented-control" role="group" aria-label="전국 현황 기준">
                <button type="button" class="is-active" data-map-metric="capacity" aria-pressed="true">설비용량</button>
                <button type="button" data-map-metric="records" aria-pressed="false">등록건수</button>
              </div>
              <div class="segmented-control" role="group" aria-label="시도 목록 정렬 방향">
                <button type="button" class="is-active" data-map-order="desc" aria-pressed="true">높은 순</button>
                <button type="button" data-map-order="asc" aria-pressed="false">낮은 순</button>
              </div>
            </div>
          </div>
          <div id="map" role="region" aria-label="전국 시도별 태양광 설비 분포 지도"></div>
          <div class="map-scale" aria-hidden="true"><strong id="map-scale-title">설비용량</strong><span>낮음</span><span class="scale-colors"><i></i><i></i><i></i><i></i><i></i></span><span>높음</span></div>
        </article>
        <article class="surface region-surface">
          <div class="section-head"><div><h2 id="region-list-title">시도별 설비용량</h2><p>전국 ${int.format(regions.length)}개 시도의 규모와 순위를 비교합니다. 전남·광주는 2026년 7월 1일 통합 행정구역을 기준으로 함께 집계합니다.</p></div></div>
          <div class="region-list-shell">
            <div id="region-list" class="region-list">${regionList(regions)}</div>
          </div>
        </article>
      </section>
      <section class="surface detail-surface" aria-labelledby="subregion-title">
        <div class="section-head detail-head">
          <div class="detail-copy"><h2 id="subregion-title"></h2><p id="subregion-caption"></p></div>
          <div class="detail-toolbar">
            <div class="detail-controls">
              <label class="field-label search-field" for="subregion-search"><span>전국 세부지역 검색</span><input id="subregion-search" type="search" autocomplete="off" placeholder="예: 전남 여수시, 강원"></label>
              <label class="field-label sort-field" for="subregion-sort"><span>정렬 기준</span><select id="subregion-sort">
                <option value="capacity" selected>설비용량</option>
                <option value="records">등록건수</option>
                <option value="name">세부지역</option>
              </select></label>
              <label class="field-label direction-field" for="subregion-order"><span>정렬 방향</span><select id="subregion-order">${renderDetailDirectionOptions()}</select></label>
            </div>
            <div id="subregion-summary" class="detail-summary" aria-live="polite"></div>
          </div>
        </div>
        <div id="subregion-table"></div>
      </section>`;
  }

  function renderRegionOptions(regions) {
    return [...regions].sort((a, b) => displayRegion(a.region).localeCompare(displayRegion(b.region), "ko"))
      .map((row) => `<option value="${escapeHtml(row.region)}"${row.region === state.selectedRegion ? " selected" : ""}>${escapeHtml(displayRegion(row.region))}</option>`).join("");
  }

  function nationalValue(row, metric = state.nationalMetric) {
    return metric === "records" ? Number(row.generator_records) || 0 : Number(row.capacity_mw) || 0;
  }

  function nationalText(row, metric = state.nationalMetric) {
    return metric === "records" ? `${int.format(Number(row.generator_records) || 0)}건` : `${formatNumber(row.capacity_mw, 2)} MW`;
  }

  function regionList(regions) {
    const ordered = [...regions].sort((a, b) => {
      const direction = state.nationalSortDirection === "asc" ? 1 : -1;
      const value = direction * (nationalValue(a) - nationalValue(b));
      return value !== 0
        ? value
        : displayRegion(a.region).localeCompare(displayRegion(b.region), "ko");
    });
    const maximum = Math.max(...ordered.map((row) => nationalValue(row)), 1);
    return ordered.map((row, index) => {
      const secondary = state.nationalMetric === "records" ? `${formatNumber(row.capacity_mw, 1)} MW` : `${int.format(Number(row.generator_records) || 0)}건`;
      const width = Math.max(2, nationalValue(row) / maximum * 100);
      return `<button type="button" class="region-item${row.region === state.selectedRegion ? " is-selected" : ""}" data-region="${escapeHtml(row.region)}" aria-pressed="${row.region === state.selectedRegion}">
        <span class="region-rank">${index + 1}</span>
        <span class="region-copy"><strong>${escapeHtml(shortRegion(row.region))}</strong><small>${escapeHtml(secondary)}</small><i class="value-track"><b style="width:${width.toFixed(2)}%"></b></i></span>
        <span class="region-value">${escapeHtml(nationalText(row))}</span>
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

  function renderDetailDirectionOptions() {
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
    return `<th${numeric ? ' class="numeric"' : ""}${detailAriaSort(key)}><button type="button" class="table-sort-button${active ? " is-active" : ""}" data-detail-sort="${key}" aria-label="${escapeHtml(`${label}, ${current}`)}">${escapeHtml(label)}<span aria-hidden="true">${indicator}</span></button></th>`;
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

  function renderDetailTable(rows, searching, aliasMatch = null) {
    if (!rows.length) {
      const message = aliasMatch
        ? `${aliasMatch.alias}는 ${aliasMatch.location}에 속하지만 현재 원천에는 해당 세부지역으로 등록된 설비 행이 없습니다.`
        : searching ? "시도명·약칭 또는 세부지역명을 바꿔 검색해 보세요." : "표시할 세부지역 정보가 없습니다.";
      return `<div class="detail-table-shell is-empty">${renderEmptyState("세부지역 결과 없음", message)}</div>`;
    }
    const showTrack = state.detailSortKey !== "name";
    const maximum = showTrack
      ? Math.max(...rows.map((row) => detailNumericValue(row)), 1)
      : 1;
    return `<div class="table-responsive detail-table-shell"><table class="data-table subregion-table">
      <thead><tr><th>순위</th>${searching ? "<th>시도</th>" : ""}${detailSortHeader("name", "세부지역")}${detailSortHeader("records", "등록건수", true)}${detailSortHeader("capacity", "설비용량(MW)", true)}</tr></thead>
      <tbody>${rows.map((row, index) => `<tr>
        <td class="rank-cell">${index + 1}</td>
        ${searching ? `<td><button type="button" class="text-button" data-detail-region="${escapeHtml(row.region)}" aria-label="${escapeHtml(displayRegion(row.region))} 상세 보기">${escapeHtml(shortRegion(row.region))}</button></td>` : ""}
        <td><div class="name-line"><strong>${escapeHtml(cleanSubregion(row))}</strong>${row.source_region_conflict ? '<span class="status-tag caution">공식 주소 미확인</span>' : ""}</div>${showTrack ? `<div class="inline-track"><span style="width:${Math.max(1, detailNumericValue(row) / maximum * 100).toFixed(2)}%"></span></div>` : ""}</td>
        <td class="numeric">${int.format(Number(row.generator_records) || 0)}</td><td class="numeric">${formatNumber(row.capacity_mw, 2)}</td>
      </tr>`).join("")}</tbody>
    </table></div>`;
  }

  function bindCoverageEvents(inventory, boundaries) {
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
      const selectedRegionLabel = state.selectedRegion ? displayRegion(state.selectedRegion) : "선택 지역";
      document.getElementById("subregion-title").textContent = searching ? "전국 세부지역 검색 결과" : `${selectedRegionLabel} 세부지역`;
      document.getElementById("subregion-caption").textContent = searching
        ? aliasMatch && rows.length
          ? `${aliasMatch.alias}는 ${aliasMatch.location}에 속합니다. 표는 ${cleanSubregion(rows[0])} 전체 등록 집계이며 ${detailSortDescription()}입니다.`
          : `${rows.length}개 결과를 ${detailSortDescription()}으로 표시합니다. 시도를 선택하면 해당 지역의 전체 현황으로 이동합니다.`
        : `${rows.length}개 세부지역을 ${detailSortDescription()}으로 비교합니다.`;
      document.getElementById("subregion-summary").innerHTML = rows.length ? `<div class="selected-summary"><span><small>${searching ? "검색 결과 등록" : "설비 등록"}</small><strong>${int.format(summary.generatorRecords)}건</strong></span><span><small>${searching ? "검색 결과 용량" : "설비용량"}</small><strong>${formatNumber(summary.capacityMw, 2)} MW</strong></span></div>` : "";
      document.getElementById("subregion-table").innerHTML = renderDetailTable(rows, searching, aliasMatch);
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
        directionSelect.innerHTML = renderDetailDirectionOptions();
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
      directionSelect.innerHTML = renderDetailDirectionOptions();
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
    document.querySelectorAll("[data-map-order]").forEach((button) => button.addEventListener("click", () => {
      state.nationalSortDirection = button.dataset.mapOrder;
      document.querySelectorAll("[data-map-order]").forEach((candidate) => {
        const active = candidate === button;
        candidate.classList.toggle("is-active", active);
        candidate.setAttribute("aria-pressed", String(active));
      });
      renderList();
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
    const colors = ["#eaf4f1", "#cde2db", "#9cc4b4", "#64a58d", "#2d7c64"];
    const activeColors = ["#d8ebe5", "#b0d5ca", "#8ec0a3", "#6ea98a", "#3a8b68"];
    const selectedStroke = "#0a4c3a";
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
    const resolveColor = (value, active = false) => {
      const share = Number(value || 0) / maximum();
      const palette = active ? activeColors : colors;
      return share >= .72 ? palette[4] : share >= .42 ? palette[3] : share >= .20 ? palette[2] : share >= .08 ? palette[1] : palette[0];
    };
    const featureName = (feature) => names[feature.properties.id] || feature.properties.name;
    const style = (feature) => {
      const name = featureName(feature);
      const value = Number(byName.get(name)?.[key()] || 0);
      if (name === selected) {
        return {
          color: selectedStroke,
          weight: 3.6,
          fillColor: resolveColor(value, true),
          fillOpacity: .92,
          opacity: 1,
        };
      }
      return {
        color: "#698c81",
        weight: 1.1,
        fillColor: resolveColor(value),
        fillOpacity: .56,
        opacity: 1,
      };
    };
    const layer = L.geoJSON(boundaries, {
      style,
      onEachFeature: (feature, shape) => {
        const name = featureName(feature);
        const row = byName.get(name) || {};
        const renderHover = (active = false) => ({
          color: active || name === selected ? selectedStroke : "#0a5b48",
          weight: active || name === selected ? 3.6 : 1.1,
          fillColor: resolveColor(Number(byName.get(name)?.[key()] || 0), true),
          fillOpacity: active || name === selected ? .92 : .56,
          opacity: 1,
        });
        shape.bindTooltip(() => `<strong>${escapeHtml(displayRegion(name))}</strong><div class="map-tooltip-metrics"><span><small>설비 등록</small><b>${escapeHtml(`${int.format(Number(row.generator_records) || 0)}건`)}</b></span><span><small>설비용량</small><b>${escapeHtml(`${formatNumber(row.capacity_mw, 2)} MW`)}</b></span></div>`, { sticky: true, direction: "top", className: "province-hover-tooltip", opacity: 1 });
        shape.on({
          mouseover: (event) => {
            event.target.setStyle(renderHover(true));
            event.target.bringToFront();
          },
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
