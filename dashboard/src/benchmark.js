// Optimized historical benchmark; independent from legacy anomaly analysis.
  function renderBenchmarkSection(benchmark = {}) {
    if (benchmark.status !== "ready" || !(benchmark.tasks || []).length) {
      return `<section class="surface analysis-section benchmark-section" aria-label="모델 최적화 벤치마크"><div class="section-head"><div><h2>모델 최적화 벤치마크</h2><p>${escapeHtml(benchmark.message || "완료된 정식 벤치마크 결과가 없습니다.")}</p></div></div></section>`;
    }
    return `<section class="surface analysis-section benchmark-section" aria-labelledby="benchmark-title">
      <div class="section-head"><div><h2 id="benchmark-title">각 모델 최적화 · 하이브리드 채택 결과</h2><p>실측 기상·발전량으로 평가한 과거 Test 결과입니다. 같은 예측 시간의 공통 표본끼리 비교합니다.</p></div><span class="status-tag">정식 평가</span></div>
      <p class="benchmark-method">Train 학습 → Validation에서 각 모델 설정 선택 → Calibration 앞 구간에서 Hybrid 가중치 학습 → 뒤 구간에서 채택 모델 선택 → Test 최종 평가</p>
      <label class="field-label benchmark-horizon" for="benchmark-horizon"><span>예측 시간</span><select id="benchmark-horizon">${benchmark.tasks.map((task, index) => `<option value="${index}">${escapeHtml(formatNumber(task.horizon_hours))}시간 후 발전량</option>`).join("")}</select></label>
      <div id="benchmark-content"></div>
    </section>`;
  }

  function bindBenchmarkEvents(benchmark = {}) {
    const horizon = document.getElementById("benchmark-horizon");
    if (!horizon) return;
    const render = () => {
      const task = benchmark.tasks[Number(horizon.value)];
      document.getElementById("benchmark-content").innerHTML = renderBenchmarkTask(task);
      const plantSelect = document.getElementById("benchmark-plant");
      const modelSelect = document.getElementById("benchmark-model");
      const renderSeries = () => {
        const points = task.series.filter((row) => row.plant_id === plantSelect.value);
        const models = modelSelect.value === "all" ? task.models : task.models.filter((model) => model.id === modelSelect.value);
        document.getElementById("benchmark-series").innerHTML = points.length
          ? renderLineChart(points, models, points[0].plant, "benchmark-series-chart")
          : renderEmptyState("표시할 예측값이 없습니다", "선택한 발전소의 평가 시계열을 확인할 수 없습니다.");
      };
      plantSelect.addEventListener("change", renderSeries);
      modelSelect.addEventListener("change", renderSeries);
      renderSeries();
    };
    horizon.addEventListener("change", render);
    render();
  }

  function benchmarkPeriod(period = {}) {
    return period.start && period.end ? `${formatDate(period.start)} ~ ${formatDate(period.end)}` : "-";
  }

  function renderBenchmarkTask(task) {
    const coverage = task.coverage || {};
    const plants = [...new Map((task.series || []).map((row) => [row.plant_id, row])).values()];
    const persistence = task.persistence_comparison?.test;
    return (task.provenance?.optimization_scope === "bounded_observed_pilot" ? '<p class="status-message">실측 데이터 일부 범위에서 수행한 제한된 예산의 실험입니다. 전체 데이터 최적화 결과로 해석하지 않습니다.</p>' : "") +
      `<div class="benchmark-selection"><strong>${escapeHtml(formatNumber(task.horizon_hours))}시간 후 · ${escapeHtml(task.selected_label)}</strong><span class="status-tag">채택</span><p>${escapeHtml(task.reason)}</p></div>` +
      renderSummaryStats([["최종 Test 표본", int.format(coverage.test_rows || 0), "건"], ["평가 발전소", int.format(coverage.plants || 0), "개"], ["평가 지역", int.format(coverage.regions || 0), "개"]]) +
      `<p class="benchmark-periods">가중치 학습: ${escapeHtml(benchmarkPeriod(task.periods?.gate_fit))}<br>모델 선택: ${escapeHtml(benchmarkPeriod(task.periods?.selection))}<br>최종 Test: ${escapeHtml(benchmarkPeriod(task.periods?.test))}</p>
      <h3>선택 구간과 최종 Test 오차</h3><p class="muted">발전소·시간 표본을 모아 계산한 지표입니다. 선택 MAE로 채택을 결정하며 Test 점수는 최종 보고에만 사용합니다.${isFiniteValue(task.alignment?.test?.common_fraction) ? ` Test 공통 표본 비율 ${escapeHtml(formatPercent(task.alignment.test.common_fraction))}.` : ""}</p>
      <div class="table-responsive"><table class="data-table"><thead><tr><th>모델</th><th class="numeric">선택 MAE</th><th class="numeric">Test MAE</th><th class="numeric">Test RMSE</th><th class="numeric">Test R²</th></tr></thead><tbody>${task.models.map((model) => `<tr${model.selected ? ' class="is-selected"' : ""}><td><strong>${escapeHtml(model.label)}</strong>${model.selected ? '<span class="status-tag">채택</span>' : ""}</td><td class="numeric">${escapeHtml(metricText("mae", model.selection_metrics.pooled.mae))}</td><td class="numeric">${escapeHtml(metricText("mae", model.test_metrics.pooled.mae))}</td><td class="numeric">${escapeHtml(metricText("rmse", model.test_metrics.pooled.rmse))}</td><td class="numeric">${escapeHtml(metricText("r2", model.test_metrics.pooled.r2))}</td></tr>`).join("")}</tbody></table></div>
      ${persistence && persistence.beats_persistence === false ? '<p class="status-message">채택 모델의 최종 Test MAE가 직전 관측 유지 기준선보다 낮지 않습니다. 모델 채택은 이전 선택 구간에서 결정한 결과입니다.</p>' : ""}
      ${renderBenchmarkOptimization(task.optimization)}
      <details class="benchmark-details"><summary>지역 합산 발전량 성능</summary><p>평가에 포함된 발전소들의 같은 시간 발전량을 합산해 계산합니다. 전국의 모든 설비를 대표하는 수치는 아닙니다.</p>${renderBenchmarkRegionTable(task)}</details>
      <h3>실제 발전량과 모델 예측</h3><p class="muted">Test의 가장 최근 연속 구간을 발전소별 최대 168시간 표시합니다. 전체 Test 지표는 위 표에서 확인합니다.</p>
      <div class="analysis-toolbar"><label class="field-label" for="benchmark-plant"><span>발전소</span><select id="benchmark-plant">${plants.map((row) => `<option value="${escapeHtml(row.plant_id)}">${escapeHtml(row.region)} · ${escapeHtml(row.plant)}</option>`).join("")}</select></label><label class="field-label" for="benchmark-model"><span>표시 모델</span><select id="benchmark-model"><option value="${escapeHtml(task.selected_model)}">채택 모델 · ${escapeHtml(task.selected_label)}</option><option value="all">모든 모델과 기준선</option>${task.models.filter((model) => !model.selected).map((model) => `<option value="${escapeHtml(model.id)}">${escapeHtml(model.label)}</option>`).join("")}</select></label></div>
      <div id="benchmark-series"></div>
      <details class="benchmark-details"><summary>평가 데이터 근거</summary><p>실측 데이터 기준 · 평가 범위 ${escapeHtml(benchmarkPeriod(coverage))} · ${int.format(coverage.plants || 0)}개 발전소</p><p class="benchmark-fingerprint">데이터 식별값: <code>${escapeHtml(String(task.provenance?.dataset_fingerprint || ""))}</code></p><p>이 결과는 이번 최적화 실험의 검증 결과입니다. 과거 저장 모델의 성능 재현을 의미하지 않습니다.</p></details>`;
  }

  function renderBenchmarkOptimization(optimization = {}) {
    const rows = Object.entries(optimization);
    if (!rows.length) return "";
    return `<details class="benchmark-details"><summary>Validation에서 선택한 모델 설정</summary><p>정해진 탐색 후보와 예산 안에서 선택한 설정입니다. 모델별 입력 길이와 특성, 하이퍼파라미터는 다를 수 있습니다.</p>${rows.map(([model, setting]) => `<h4>${escapeHtml(model === "xgboost" ? "XGBoost" : "CNN-BiLSTM")}</h4><p>Validation MAE: ${escapeHtml(metricText("mae", setting.validation_mae))}${setting.sequence_length ? ` · 입력 길이: ${escapeHtml(formatNumber(setting.sequence_length))}시간` : ""} · 특성 ${int.format((setting.feature_columns || []).length)}개</p><p>${escapeHtml((setting.feature_columns || []).join(", "))}</p><pre class="benchmark-parameters">${escapeHtml(JSON.stringify(setting.parameters || {}, null, 2))}</pre>`).join("")}</details>`;
  }

  function renderBenchmarkRegionTable(task) {
    const rows = task.models.flatMap((model) => [
      { region: "평가 발전소 전체 합산", model: model.label, ...model.test_metrics.national },
      ...model.test_metrics.region.map((row) => ({ ...row, model: model.label })),
    ]);
    return `<div class="table-responsive"><table class="data-table"><thead><tr><th>범위</th><th>모델</th><th class="numeric">Test MAE</th><th class="numeric">Test RMSE</th><th class="numeric">Test R²</th><th class="numeric">평가 시점</th></tr></thead><tbody>${rows.map((row) => `<tr><td>${escapeHtml(row.region)}</td><td>${escapeHtml(row.model)}</td><td class="numeric">${escapeHtml(metricText("mae", row.mae))}</td><td class="numeric">${escapeHtml(metricText("rmse", row.rmse))}</td><td class="numeric">${escapeHtml(metricText("r2", row.r2))}</td><td class="numeric">${int.format(row.n_samples)}</td></tr>`).join("")}</tbody></table></div>`;
  }
