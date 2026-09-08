// Forecast view helpers; composed into one private browser scope.
  function forecastModels(analysis) {
    const points = normalizeSeries(analysis);
    return (analysis.models || []).filter((model) =>
      points.some((point) => isFiniteValue(point.predictions?.[model.id]))
    );
  }

  function renderForecastModelOptions(analysis) {
    return forecastModels(analysis).map((model) =>
      `<option value="${escapeHtml(model.id)}"${model.id === state.forecastModelId ? " selected" : ""}>${escapeHtml(model.label)}</option>`
    ).join("");
  }

  function regionsForForecast(analysis) {
    return [...new Set(normalizeSeries(analysis).map((row) => row.region).filter(Boolean))]
      .sort((a, b) => displayRegion(a).localeCompare(displayRegion(b), "ko"));
  }

  function renderForecastRegionOptions(analysis) {
    return '<option value="all">전국</option>' + regionsForForecast(analysis).map((region) =>
      `<option value="${escapeHtml(region)}"${region === state.region ? " selected" : ""}>${escapeHtml(displayRegion(region))}</option>`
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

  function renderForecastPlantOptions(analysis) {
    return '<option value="all">발전소 선택</option>' + plantsForForecast(analysis).map((row) =>
      `<option value="${escapeHtml(row.plant_id)}"${row.plant_id === state.plantId ? " selected" : ""}>${escapeHtml(row.plant)}</option>`
    ).join("");
  }

  function renderForecastPage(analysis) {
    const models = forecastModels(analysis);
    if (!models.some((model) => model.id === state.forecastModelId)) {
      state.forecastModelId = models[0]?.id || "all";
    }
    const header = renderPageIntro(
      "테스트 구간 발전량 예측 결과",
      "정식 평가에 사용된 과거 Test 구간의 실제 발전량과 모델 예측값을 발전소별로 확인합니다.",
      evaluationContext(analysis.evaluation),
    ) +
      `<p class="forecast-note">현재·미래 운영 예측이 아니라 저장된 Test 평가 결과입니다. 예보 발행시각이 보존된 기상예보 입력이 연결되기 전에는 미래 발전량을 표시하지 않습니다.</p>` +
      (analysis.status !== "ready" && analysis.message ? `<p class="status-message">${escapeHtml(analysis.message)}</p>` : "");
    if (!models.length) {
      return header + `<section id="forecast-content" class="forecast-content">${renderEmptyState("예측 결과가 없습니다", analysis.message || "정식 모델의 Test 예측이 완료되면 발전소별 시계열을 표시합니다.")}</section>`;
    }
    return header + `<section class="analysis-toolbar forecast-toolbar" aria-label="발전량 예측 조건">
        <label class="field-label" for="forecast-region"><span>지역</span><select id="forecast-region">${renderForecastRegionOptions(analysis)}</select></label>
        <label class="field-label" for="forecast-plant"><span>발전소</span><select id="forecast-plant">${renderForecastPlantOptions(analysis)}</select></label>
        <label class="field-label" for="forecast-model"><span>예측 모델</span><select id="forecast-model">${renderForecastModelOptions(analysis)}</select></label>
      </section>
      <section id="forecast-content" class="forecast-content"></section>`;
  }

  function bindForecastEvents(analysis) {
    const content = document.getElementById("forecast-content");
    const regionSelect = document.getElementById("forecast-region");
    if (!regionSelect) return;
    const render = () => { content.innerHTML = renderForecastView(analysis); };
    regionSelect.addEventListener("change", (event) => {
      state.region = event.target.value;
      state.plantId = "all";
      document.getElementById("forecast-plant").innerHTML = renderForecastPlantOptions(analysis);
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

  function renderForecastView(analysis) {
    const models = forecastModels(analysis);
    const model = models.find((candidate) => candidate.id === state.forecastModelId);
    if (!model) {
      return renderEmptyState("예측 결과가 없습니다", analysis.message || "정식 모델의 Test 예측이 완료되면 발전소별 시계열을 표시합니다.");
    }
    if (state.plantId === "all") {
      return `<section class="surface forecast-series">${renderEmptyState("발전소를 선택해 주세요", "상단에서 지역과 발전소를 선택하면 실제 발전량과 선택 모델의 예측값을 표시합니다.")}</section>`;
    }
    const selected = plantsForForecast(analysis).find((row) => row.plant_id === state.plantId);
    const label = selected?.plant || state.plantId;
    const points = normalizeSeries(analysis).filter((row) =>
      (state.region === "all" || row.region === state.region) &&
      row.plant_id === state.plantId &&
      isFiniteValue(row.predictions?.[model.id])
    );
    if (!points.length) {
      return renderEmptyState("선택 조건의 예측 결과가 없습니다", "다른 발전소나 모델을 선택해 확인해 주세요.");
    }
    const horizon = isFiniteValue(analysis.evaluation?.horizon_hours) ? formatNumber(analysis.evaluation.horizon_hours) : "-";
    return renderSummaryStats([
      ["선택 모델", model.label || model.id],
      ["표시 시점", int.format(points.length), "건"],
      ["예측 범위", horizon, isFiniteValue(analysis.evaluation?.horizon_hours) ? "시간" : ""],
    ]) +
      `<section class="surface forecast-series"><div class="section-head"><div><h2>${escapeHtml(label)} 발전량 예측</h2><p>Test 구간에서 가장 최근의 연속 최대 168시간을 실제 발전량과 ${escapeHtml(model.label || model.id)} 예측값으로 표시합니다.</p></div></div>${renderLineChart(points, [model], label)}</section>`;
  }
