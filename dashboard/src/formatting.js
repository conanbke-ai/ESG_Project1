// Formatting view helpers; composed into one private browser scope.
  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (character) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;",
    })[character]);
  }

  function isFiniteValue(value) {
    return value !== null && value !== "" && Number.isFinite(Number(value));
  }

  function formatNumber(value, digits = 0) {
    if (!isFiniteValue(value)) return "-";
    return Number(value).toLocaleString("ko-KR", { maximumFractionDigits: digits });
  }

  function formatPercent(value) {
    return isFiniteValue(value) ? `${formatNumber(Number(value) * 100, 1)}%` : "-";
  }

  function renderPageIntro(title, description, context = "") {
    return `<section class="page-intro"><div><h1>${escapeHtml(title)}</h1><p>${escapeHtml(description)}</p></div>${context ? `<p class="page-context">${escapeHtml(context)}</p>` : ""}</section>`;
  }

  function renderSummaryStats(items) {
    return `<section class="summary-grid">${items.map(([label, value, unit = ""]) => `<article class="summary-item"><span>${escapeHtml(label)}</span><div class="summary-value"><strong>${escapeHtml(value)}</strong>${unit ? `<small>${escapeHtml(unit)}</small>` : ""}</div></article>`).join("")}</section>`;
  }

  function renderEmptyState(title, message) {
    return `<div class="empty-state"><h2>${escapeHtml(title)}</h2><p>${escapeHtml(message)}</p></div>`;
  }

  function displayRegion(region) {
    if (!region || region === "unknown") return "지역 미확인";
    return region === "전남광주통합특별시" ? "전남·광주" : region;
  }

  function shortRegion(region) {
    return ({
      서울특별시: "서울", 부산광역시: "부산", 대구광역시: "대구", 인천광역시: "인천",
      광주광역시: "광주", 대전광역시: "대전", 울산광역시: "울산", 세종특별자치시: "세종",
      경기도: "경기", 강원특별자치도: "강원", 충청북도: "충북", 충청남도: "충남",
      전북특별자치도: "전북", 전라남도: "전남", 경상북도: "경북", 경상남도: "경남",
      제주특별자치도: "제주", unknown: "지역 미확인",
    })[region] || displayRegion(region);
  }

  function emptyAnalysis() {
    return { status: "empty", message: "비교 가능한 평가 결과가 아직 없습니다.", evaluation: {}, models: [], regions: [], plants: [], series: [], anomalies: { prediction_summary: { total: 0, returned_top_events: 0, by_model: [], by_region: [], by_plant: [] }, prediction_signals: [], data_quality_signals: [] } };
  }

  function formatDate(value) {
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? String(value || "").slice(0, 10) : new Intl.DateTimeFormat("ko-KR", { year: "numeric", month: "2-digit", day: "2-digit" }).format(parsed);
  }

  function evaluationContext(evaluation = {}) {
    const parts = [];
    if (evaluation.from && evaluation.to) parts.push(`${formatDate(evaluation.from)} ~ ${formatDate(evaluation.to)}`);
    if (isFiniteValue(evaluation.horizon_hours) && Number(evaluation.horizon_hours) > 0) parts.push(`${formatNumber(evaluation.horizon_hours)}시간 예측`);
    if (isFiniteValue(evaluation.common_samples) && Number(evaluation.common_samples) > 0) parts.push(`공통 평가 ${int.format(Number(evaluation.common_samples))}건`);
    return parts.join(" · ");
  }

  function normalizeSeries(analysis) {
    const grouped = new Map();
    (analysis.series || []).forEach((row) => {
      if (!row.timestamp || !row.plant_id) return;
      const key = [row.timestamp, row.plant_id].join("|");
      const point = grouped.get(key) || { timestamp: row.timestamp, plant_id: String(row.plant_id), region: row.region, plant: row.plant, y_true: isFiniteValue(row.y_true) ? Number(row.y_true) : null, predictions: {} };
      if (isFiniteValue(row.y_true)) point.y_true = Number(row.y_true);
      if (row.predictions && typeof row.predictions === "object") Object.entries(row.predictions).forEach(([model, value]) => { if (isFiniteValue(value)) point.predictions[model] = Number(value); });
      (analysis.models || []).forEach((model) => {
        const id = String(model.id || "");
        const lower = id.toLowerCase();
        const candidates = [row[`${id}_pred`], lower.includes("xg") ? row.xgb_pred : null, lower.includes("cnn") ? row.cnn_pred : null, lower.includes("hybrid") ? row.hybrid_pred : null, row.model === id ? row.y_pred : null];
        const value = candidates.find(isFiniteValue);
        if (value !== undefined) point.predictions[id] = Number(value);
      });
      grouped.set(key, point);
    });
    return [...grouped.values()].sort((a, b) => String(a.timestamp).localeCompare(String(b.timestamp)));
  }

  function renderLineChart(source, models, plantLabel) {
    const points = source.length <= 168 ? source : Array.from({ length: 168 }, (_, index) => source[Math.round(index * (source.length - 1) / 167)]);
    const visible = models.filter((model) => points.some((point) => isFiniteValue(point.predictions[model.id])));
    const values = points.flatMap((point) => [point.y_true, ...visible.map((model) => point.predictions[model.id])]).filter(isFiniteValue).map(Number);
    if (!values.length) return renderEmptyState("표시할 값이 없습니다", "실제값과 예측값을 확인할 수 없습니다.");
    const width = 920, height = 300, pad = { left: 58, right: 20, top: 20, bottom: 42 };
    let minimum = Math.min(...values), maximum = Math.max(...values);
    if (minimum === maximum) { minimum -= 1; maximum += 1; }
    const x = (index) => pad.left + index / Math.max(1, points.length - 1) * (width - pad.left - pad.right);
    const y = (value) => pad.top + (maximum - value) / (maximum - minimum) * (height - pad.top - pad.bottom);
    const path = (getter) => {
      let open = false;
      return points.map((point, index) => {
        const value = getter(point);
        if (!isFiniteValue(value)) { open = false; return ""; }
        const command = open ? "L" : "M";
        open = true;
        return `${command}${x(index).toFixed(2)},${y(Number(value)).toFixed(2)}`;
      }).join(" ");
    };
    const ticks = [minimum, minimum + (maximum - minimum) / 2, maximum];
    const timeIndexes = [...new Set([0, Math.floor((points.length - 1) / 2), points.length - 1])];
    return `<div class="chart-legend"><span><i class="actual"></i>실제 발전량</span>${visible.map((model, index) => `<span><i class="series-${index % 4}"></i>${escapeHtml(model.label)}</span>`).join("")}</div>
      <svg class="line-chart" viewBox="0 0 ${width} ${height}" role="img" aria-labelledby="series-chart-title series-chart-desc">
        <title id="series-chart-title">${escapeHtml(plantLabel)} 실제 발전량과 예측값</title><desc id="series-chart-desc">실제 발전량과 모델별 예측값을 비교한 선 그래프입니다.</desc>
        ${ticks.map((tick) => `<line class="chart-gridline" x1="${pad.left}" x2="${width - pad.right}" y1="${y(tick)}" y2="${y(tick)}"></line><text class="chart-axis-label" x="${pad.left - 10}" y="${y(tick) + 4}" text-anchor="end">${escapeHtml(formatNumber(tick, 2))}</text>`).join("")}
        ${timeIndexes.map((index) => `<text class="chart-axis-label" x="${x(index)}" y="${height - 14}" text-anchor="${index === 0 ? "start" : index === points.length - 1 ? "end" : "middle"}">${escapeHtml(shortTime(points[index].timestamp))}</text>`).join("")}
        <text class="chart-axis-title" x="14" y="${height / 2}" transform="rotate(-90 14 ${height / 2})">발전량(MWh)</text>
        <path class="chart-line actual" d="${path((point) => point.y_true)}"></path>${visible.map((model, index) => `<path class="chart-line series-${index % 4}" d="${path((point) => point.predictions[model.id])}"></path>`).join("")}
      </svg>`;
  }

  function shortTime(value) {
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? String(value).slice(0, 16) : new Intl.DateTimeFormat("ko-KR", { month: "2-digit", day: "2-digit", hour: "2-digit" }).format(parsed);
  }
