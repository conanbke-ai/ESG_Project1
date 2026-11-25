// Shared dashboard state, view selection, and metric labels.
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
    nationalSortDirection: "desc",
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
