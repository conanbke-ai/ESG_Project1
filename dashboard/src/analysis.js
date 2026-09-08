// Analysis view helpers; composed into one private browser scope.
  function renderAnalysisPage(analysis) {
    return renderPageIntro("태양광 모델 성능 비교·분석", "동일한 평가 표본에서 모델 정확도를 비교하고 지역·발전소별 오차와 이상 신호를 살펴봅니다.", evaluationContext(analysis.evaluation)) +
      (analysis.status !== "ready" && analysis.message ? `<p class="status-message">${escapeHtml(analysis.message)}</p>` : "") +
      `<section class="analysis-toolbar" aria-label="분석 조건">
        <label class="field-label" for="analysis-region"><span>지역</span><select id="analysis-region">${renderAnalysisRegionOptions(analysis)}</select></label>
        <label class="field-label" for="analysis-plant"><span>발전소</span><select id="analysis-plant">${renderAnalysisPlantOptions(analysis)}</select></label>
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
      total: isFiniteValue(raw.total) ? Number(raw.total) : events.length,
      returned_top_events: isFiniteValue(raw.returned_top_events) ? Number(raw.returned_top_events) : events.length,
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

  function renderAnalysisRegionOptions(analysis) {
    return '<option value="all">전국</option>' + regionsForAnalysis(analysis).map((region) => `<option value="${escapeHtml(region)}"${region === state.region ? " selected" : ""}>${escapeHtml(displayRegion(region))}</option>`).join("");
  }

  function renderAnalysisPlantOptions(analysis) {
    return '<option value="all">전체 발전소</option>' + plantsForAnalysis(analysis).map((row) => `<option value="${escapeHtml(row.plant_id)}"${row.plant_id === state.plantId ? " selected" : ""}>${escapeHtml(row.plant)}</option>`).join("");
  }

  function bindAnalysisEvents(analysis) {
    const content = document.getElementById("analysis-content");
    const toolbar = document.querySelector(".analysis-toolbar");
    const tabs = [...document.querySelectorAll("[data-tab]")];
    const render = () => {
      toolbar.hidden = state.tab === "comparison";
      content.innerHTML = state.tab === "performance" ? renderPerformanceView(analysis) : state.tab === "anomalies" ? renderAnomalyView(analysis) : renderComparisonView(analysis);
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
      document.getElementById("analysis-plant").innerHTML = renderAnalysisPlantOptions(analysis);
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
      document.getElementById("analysis-plant").innerHTML = renderAnalysisPlantOptions(analysis);
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
    return isFiniteValue(value) ? Number(value) : null;
  }

  function metricText(key, value) {
    const definition = metrics[key] || metrics.mae;
    return isFiniteValue(value) ? `${formatNumber(value, definition.digits)}${definition.unit ? ` ${definition.unit}` : ""}` : "-";
  }

  function metricHint(key) {
    const direction = metrics[key].better === "higher" ? "높을수록 좋음" : "낮을수록 좋음";
    return key === "nmae_capacity" ? `${direction} · 설비용량 확인 표본 기준` : direction;
  }

  function renderComparisonView(analysis) {
    const allModels = analysis.models || [];
    const models = allModels.filter((model) => Object.keys(metrics).some((key) => metricValue(model, key) !== null));
    const comparable = models.filter((model) => model.comparable !== false);
    if (!models.length) return renderEmptyState("모델 평가 결과가 없습니다", analysis.message || "모델 평가가 완료되면 네 가지 성능 지표가 함께 표시됩니다.");
    if (comparable.length < 2) {
      return renderEmptyState("모델 간 비교 결과가 없습니다", analysis.message || "동일한 테스트 표본의 모델이 두 개 이상 확보되면 네 가지 지표 비교가 활성화됩니다.") +
        `<section class="surface analysis-section"><div class="section-head"><div><h2>확인 가능한 개별 결과</h2><p>평가 구간이 같지 않아 모델 간 순위나 막대 비교에는 사용하지 않습니다.</p></div></div>${renderModelTable(models)}</section>`;
    }
    const commonSamples = isFiniteValue(analysis.evaluation?.common_samples) ? Number(analysis.evaluation.common_samples) : 0;
    const horizon = isFiniteValue(analysis.evaluation?.horizon_hours) ? formatNumber(analysis.evaluation.horizon_hours) : "-";
    let result = renderSummaryStats([
      ["비교 모델", int.format(comparable.length), "개"],
      ["공통 평가 표본", int.format(commonSamples), "건"],
      ["예측 범위", horizon, isFiniteValue(analysis.evaluation?.horizon_hours) ? "시간" : ""],
    ]);
    result += `<section class="metric-overview-grid" aria-label="모델별 네 가지 성능 지표">${Object.keys(metrics).map((key) => renderMetricPanel(comparable, key)).join("")}</section>`;
    result += `<section class="surface analysis-section"><div class="section-head"><div><h2>모델별 전체 평가 지표</h2><p>NMAE, MAE, RMSE, R²를 같은 평가 구간에서 함께 확인합니다.</p></div></div>${renderModelTable(models)}</section>`;
    return result;
  }

  function renderMetricPanel(models, key) {
    const available = models.filter((model) => metricValue(model, key) !== null);
    return `<article class="surface metric-panel"><div class="metric-panel-head"><div><strong>${escapeHtml(metrics[key].short)}</strong><span>${escapeHtml(metrics[key].label)}</span></div><small>${escapeHtml(metricHint(key))}</small></div>${available.length ? renderModelBars(available, key) : renderEmptyState(`${metrics[key].short} 결과 없음`, "이 지표를 계산할 수 있는 평가 표본이 없습니다.")}</article>`;
  }

  function renderModelBars(models, key) {
    const values = models.map((model) => metricValue(model, key));
    const maximum = Math.max(...values, 0);
    const minimum = Math.min(...values, 0);
    return `<div class="metric-bars">${models.map((model, index) => {
      const value = metricValue(model, key);
      const width = metrics[key].better === "higher" ? (maximum === minimum ? 100 : (value - minimum) / (maximum - minimum) * 100) : (maximum ? value / maximum * 100 : 0);
      return `<div class="metric-bar-row"><div><strong>${escapeHtml(model.label)}</strong><span>${escapeHtml(metricText(key, value))}</span></div><div class="metric-track"><i class="series-${index % 4}" style="width:${Math.max(3, width).toFixed(2)}%"></i></div></div>`;
    }).join("")}</div>`;
  }

  function renderModelTable(models) {
    return `<div class="table-responsive"><table class="data-table"><thead><tr><th>모델</th><th class="numeric">NMAE</th><th class="numeric">용량 적용률</th><th class="numeric">MAE</th><th class="numeric">RMSE</th><th class="numeric">R²</th><th class="numeric">평가 표본</th></tr></thead><tbody>${models.map((model) => `<tr><td><strong>${escapeHtml(model.label)}</strong>${model.comparable === false ? '<span class="status-tag neutral">비교 제외</span>' : ""}</td><td class="numeric">${escapeHtml(metricText("nmae_capacity", metricValue(model, "nmae_capacity")))}</td><td class="numeric">${isFiniteValue(model.metrics?.capacity_coverage) ? `${formatNumber(Number(model.metrics.capacity_coverage) * 100, 1)}%` : "-"}</td><td class="numeric">${escapeHtml(metricText("mae", metricValue(model, "mae")))}</td><td class="numeric">${escapeHtml(metricText("rmse", metricValue(model, "rmse")))}</td><td class="numeric">${escapeHtml(metricText("r2", metricValue(model, "r2")))}</td><td class="numeric">${isFiniteValue(model.metrics?.n_samples) ? int.format(Number(model.metrics.n_samples)) : "-"}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function renderPerformanceView(analysis) {
    const regionRows = analysis.regions || [];
    const plantRows = analysis.plants || [];
    if (!regionRows.length && !plantRows.length) {
      return renderEmptyState("지역·발전소 성능 결과가 없습니다", analysis.message || "세부 평가가 완료되면 지역과 발전소별 성능이 표시됩니다.");
    }
    return `<section class="surface analysis-section">
        <div class="section-head"><div><h2>지역별 네 가지 성능 지표</h2><p>모델별 NMAE, MAE, RMSE, R²를 함께 보고 지역을 선택해 발전소 결과로 이어갑니다.</p></div></div>
        ${renderRegionMatrix(analysis)}
      </section>
      <section class="surface analysis-section">
        <div class="section-head"><div><h2>${state.region === "all" ? "발전소별 성능" : `${escapeHtml(displayRegion(state.region))} 발전소별 성능`}</h2><p>발전소마다 네 가지 지표와 평가 표본을 함께 확인합니다.</p></div></div>
        ${renderPlantTable(analysis)}
      </section>`;
  }

  function renderRegionMatrix(analysis) {
    const models = analysis.models || [];
    const grouped = new Map();
    (analysis.regions || []).forEach((row) => {
      if (!row.region) return;
      const values = grouped.get(row.region) || new Map();
      values.set(row.model, row.metrics || {});
      grouped.set(row.region, values);
    });
    if (!grouped.size) return renderEmptyState("지역별 결과 없음", "지역 단위 평가 결과가 아직 없습니다.");
    const ordered = [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0], "ko"));
    return `<div class="table-responsive"><table class="data-table region-matrix">
      <thead><tr><th>지역</th>${models.map((model) => `<th>${escapeHtml(model.label)}</th>`).join("")}</tr></thead>
      <tbody>${ordered.map(([region, values]) => `<tr class="${region === state.region ? "is-selected" : ""}">
        <td><button type="button" class="text-button" data-analysis-region="${escapeHtml(region)}">${escapeHtml(displayRegion(region))}</button></td>
        ${models.map((model) => `<td>${renderMetricQuartet(values.get(model.id))}</td>`).join("")}
      </tr>`).join("")}</tbody>
    </table></div>`;
  }

  function renderMetricQuartet(values = {}) {
    return `<div class="metric-quartet">${Object.keys(metrics).map((key) => `<span><small>${escapeHtml(metrics[key].short)}</small><strong>${escapeHtml(metricText(key, values?.[key]))}</strong></span>`).join("")}</div>`;
  }

  function renderPlantTable(analysis) {
    if (state.region === "all") return renderEmptyState("지역을 선택해 주세요", "위 지역별 비교에서 확인할 지역을 선택하면 발전소별 결과가 표시됩니다.");
    let rows = (analysis.plants || []).filter((row) => row.region === state.region);
    if (state.plantId !== "all") rows = rows.filter((row) => String(row.plant_id) === state.plantId);
    rows.sort((a, b) => String(a.plant || "").localeCompare(String(b.plant || ""), "ko") || String(a.model || "").localeCompare(String(b.model || "")));
    if (!rows.length) return renderEmptyState("발전소별 결과 없음", "선택한 지역의 발전소별 평가 결과가 아직 없습니다.");
    return `<div class="table-responsive"><table class="data-table plant-performance-table">
      <thead><tr><th>발전소</th><th>모델</th><th class="numeric">NMAE</th><th class="numeric">MAE</th><th class="numeric">RMSE</th><th class="numeric">R²</th><th class="numeric">평가 표본</th></tr></thead>
      <tbody>${rows.map((row) => {
        const model = (analysis.models || []).find((candidate) => candidate.id === row.model);
        return `<tr><td><button type="button" class="text-button" data-analysis-plant-id="${escapeHtml(row.plant_id)}">${escapeHtml(row.plant)}</button></td><td>${escapeHtml(model?.label || row.model)}</td><td class="numeric">${escapeHtml(metricText("nmae_capacity", metricValue(row, "nmae_capacity")))}</td><td class="numeric">${escapeHtml(metricText("mae", metricValue(row, "mae")))}</td><td class="numeric">${escapeHtml(metricText("rmse", metricValue(row, "rmse")))}</td><td class="numeric">${escapeHtml(metricText("r2", metricValue(row, "r2")))}</td><td class="numeric">${isFiniteValue(row.metrics?.n_samples) ? int.format(Number(row.metrics.n_samples)) : "-"}</td></tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderAnomalyView(analysis) {
    const summary = predictionSummary(analysis);
    const events = filteredPredictionEvents(analysis);
    const quality = filteredQualitySignals(analysis);
    const predictionCount = selectedPredictionCount(analysis);
    const total = predictionCount + quality.length;
    if (!total) {
      return renderEmptyState(
        "선택 조건의 이상 신호가 없습니다",
        summary.total + qualitySignals(analysis).length ? "다른 지역이나 발전소를 선택해 확인해 주세요." : "예측 오차와 데이터 패턴 분석이 완료되면 검토할 신호가 표시됩니다.",
      );
    }
    return `<p class="interpretation-note">이상 신호는 고장 확정이 아니라 예측 오차 또는 데이터 패턴을 추가로 확인할 대상입니다.</p>
      ${renderSummaryStats([["전체 이상 신호", int.format(total), "건"], ["예측 오차 신호", int.format(predictionCount), "건"], ["데이터 패턴 신호", int.format(quality.length), "건"]])}
      <section class="analysis-grid anomaly-grid">
        <article class="surface analysis-section"><div class="section-head"><div><h2>신호 유형</h2><p>대표 이벤트가 아닌 전체 집계 기준입니다.</p></div></div>${renderTypeBars(combinedTypeCounts(predictionCount, quality))}</article>
        <article class="surface analysis-section"><div class="section-head"><div><h2>지역별 이상 신호</h2><p>예측 오차와 데이터 패턴 신호를 합산합니다.</p></div></div>${renderSignalRegionBars(combinedRegionCounts(analysis, quality))}</article>
      </section>
      <section class="surface analysis-section"><div class="section-head"><div><h2>대표 예측 오차 이벤트</h2><p>${escapeHtml(representativeEventCaption(summary, predictionCount, events.length))}</p></div></div>${renderPredictionEventTable(events, predictionCount)}</section>
      <section class="surface analysis-section"><div class="section-head"><div><h2>데이터 패턴 확인 대상</h2><p>선택 조건에 해당하는 품질 신호를 모두 표시합니다.</p></div></div>${renderQualitySignalTable(quality)}</section>`;
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

  function renderTypeBars(counts) {
    const ordered = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    const maximum = Math.max(...ordered.map((row) => row[1]), 1);
    return `<div class="simple-bars">${ordered.map(([type, count]) => `<div><span>${escapeHtml(type)}</span><i><b style="width:${(count / maximum * 100).toFixed(2)}%"></b></i><strong>${int.format(count)}건</strong></div>`).join("")}</div>`;
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

  function renderSignalRegionBars(counts) {
    const ordered = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    const maximum = Math.max(...ordered.map((row) => row[1]), 1);
    return `<div class="simple-bars">${ordered.map(([region, count]) => `<div><button type="button" class="text-button" data-analysis-region="${escapeHtml(region)}">${escapeHtml(displayRegion(region))}</button><i><b style="width:${(count / maximum * 100).toFixed(2)}%"></b></i><strong>${int.format(count)}건</strong></div>`).join("")}</div>`;
  }

  function severity(severityValue) {
    return ({ critical: ["긴급", "danger"], high: ["높음", "danger"], medium: ["주의", "caution"], review: ["검토", "caution"], low: ["낮음", "neutral"] })[String(severityValue || "review").toLowerCase()] || ["검토", "caution"];
  }

  function signalMetricText(signal) {
    const values = [];
    const source = signal.metrics || {};
    const absoluteError = source.absolute_residual ?? source.absolute_error ?? signal.absolute_error;
    const threshold = source.anomaly_threshold ?? source.threshold ?? signal.threshold;
    if (isFiniteValue(absoluteError)) values.push(`절대오차 ${formatNumber(absoluteError, 3)} MWh`);
    if (isFiniteValue(threshold)) values.push(`판단기준 ${formatNumber(threshold, 3)} MWh`);
    if (isFiniteValue(signal.y_true)) values.push(`실제 ${formatNumber(signal.y_true, 3)} MWh`);
    if (isFiniteValue(signal.y_pred)) values.push(`예측 ${formatNumber(signal.y_pred, 3)} MWh`);
    if (isFiniteValue(source.missing_weather_rate)) values.push(`기상 공백 ${formatPercent(source.missing_weather_rate)}`);
    if (isFiniteValue(source.daylight_zero_rate)) values.push(`주간 0발전 ${formatPercent(source.daylight_zero_rate)}`);
    if (isFiniteValue(source.hourly_coverage)) values.push(`시간 자료 ${formatPercent(source.hourly_coverage)}`);
    if (isFiniteValue(source.peer_pattern_correlation)) values.push(`지역 패턴 유사도 ${formatNumber(source.peer_pattern_correlation, 2)}`);
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

  function renderPredictionEventTable(events, selectedTotal) {
    if (!events.length) {
      return renderEmptyState(
        "표시할 대표 이벤트가 없습니다",
        selectedTotal ? "전체 집계에는 신호가 있지만 선택 조건에 해당하는 대표 이벤트는 보관되지 않았습니다." : "선택 조건의 예측 오차 신호가 없습니다.",
      );
    }
    return `<div class="table-responsive"><table class="data-table prediction-event-table">
      <thead><tr><th>모델</th><th>시각</th><th>지역</th><th>발전소</th><th class="numeric">실제(MWh)</th><th class="numeric">예측(MWh)</th><th class="numeric">절대오차</th><th class="numeric">판단기준(MWh)</th><th class="numeric">기준 대비</th><th>판단 방식</th></tr></thead>
      <tbody>${events.map((event) => `<tr><td>${escapeHtml(event.model_label || event.model)}</td><td>${escapeHtml(shortTime(event.timestamp))}</td><td>${escapeHtml(displayRegion(event.region))}</td><td>${event.plant_id ? `<button type="button" class="text-button" data-analysis-plant-id="${escapeHtml(event.plant_id)}">${escapeHtml(event.plant || event.plant_id)}</button>` : escapeHtml(event.plant || "-")}</td><td class="numeric">${formatNumber(event.y_true, 3)}</td><td class="numeric">${formatNumber(event.y_pred, 3)}</td><td class="numeric">${formatNumber(event.absolute_error, 3)}</td><td class="numeric">${formatNumber(event.threshold, 3)}</td><td class="numeric">${isFiniteValue(event.exceedance_ratio) ? `${formatNumber(event.exceedance_ratio, 2)}배` : "-"}</td><td>${escapeHtml(thresholdSourceLabel(event.threshold_source))}</td></tr>`).join("")}</tbody>
    </table></div>`;
  }

  function renderQualitySignalTable(signals) {
    if (!signals.length) return renderEmptyState("데이터 패턴 신호가 없습니다", "선택 조건에서 확인할 데이터 패턴 신호가 없습니다.");
    const rank = { critical: 4, high: 3, medium: 2, review: 2, low: 1 };
    const ordered = [...signals].sort((a, b) => (rank[String(b.severity).toLowerCase()] || 2) - (rank[String(a.severity).toLowerCase()] || 2));
    return `<div class="table-responsive"><table class="data-table quality-signal-table">
      <thead><tr><th>우선도</th><th>지역</th><th>발전소</th><th>신호</th><th>관련 지표</th></tr></thead>
      <tbody>${ordered.map((signal) => {
        const [label, kind] = severity(signal.severity);
        const plant = signal.plant_id ? `<button type="button" class="text-button" data-analysis-plant-id="${escapeHtml(signal.plant_id)}">${escapeHtml(signal.plant || signal.plant_id)}</button>` : escapeHtml(signal.plant || "-");
        return `<tr><td><span class="status-tag ${kind}">${escapeHtml(label)}</span></td><td>${escapeHtml(displayRegion(signal.region))}</td><td>${plant}</td><td>${escapeHtml(signal.summary || signalTypes(signal).join(", "))}</td><td>${escapeHtml(signalMetricText(signal))}</td></tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }
