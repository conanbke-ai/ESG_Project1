// Fetch public dashboard data and initialize the selected view.
  try {
    const response = await fetch("data/dashboard_data.json", { cache: "no-store" });
    if (!response.ok) throw new Error("dashboard data unavailable");
    const data = await response.json();
    if (view === "analysis") {
      const analysis = data.model_analysis || emptyAnalysis();
      app.innerHTML = renderAnalysisPage(analysis);
      bindAnalysisEvents(analysis);
    } else if (view === "forecast") {
      const analysis = data.model_analysis || emptyAnalysis();
      app.innerHTML = renderForecastPage(analysis);
      bindForecastEvents(analysis);
    } else {
      let boundaries = null;
      try {
        const boundaryResponse = await fetch("data/korea_provinces.geojson", { cache: "no-store" });
        if (boundaryResponse.ok) boundaries = await boundaryResponse.json();
      } catch (boundaryError) {
        console.warn("province boundaries unavailable", boundaryError);
      }
      app.innerHTML = renderCoveragePage(data.national_inventory);
      bindCoverageEvents(data.national_inventory, boundaries);
    }
  } catch (error) {
    console.error(error);
    app.innerHTML = `<section class="empty-state" role="alert"><h1>화면을 불러오지 못했습니다</h1><p>잠시 후 새로고침해 주세요.</p></section>`;
  }
