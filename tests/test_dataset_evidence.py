"""Gold construction must not invent feature-selection performance evidence."""

import json

import numpy as np
import pandas as pd

from solar_forecast.collectors.plant_metadata import PlantMetadataCatalog
from solar_forecast.datasets.model_dataset_builder import NationwideModelDatasetBuilder
from solar_forecast.features.asos_features import WEATHER_COLUMN_MAP


def test_rebuilt_gold_never_reuses_historical_ablation_as_current_evidence(tmp_path):
    # Small generated inputs exercise manifest provenance, not model accuracy.
    timestamps = pd.date_range("2025-06-01", periods=72, freq="h")
    weather = pd.DataFrame(
        {column: np.zeros(len(timestamps)) for column in WEATHER_COLUMN_MAP}
    )
    weather["일시"] = timestamps
    weather["지점"] = 108
    weather["지점명"] = "서울"
    weather["기온(°C)"] = 20.0
    weather["습도(%)"] = 50.0
    weather.to_csv(tmp_path / "OBS_ASOS_TIM_2025.csv", index=False)
    pd.DataFrame(
        {"지점": [108], "위도": [37.57], "경도": [126.97], "노장해발고도(m)": [85.7]}
    ).to_csv(tmp_path / "META_관측지점정보.csv", index=False)

    builder = NationwideModelDatasetBuilder(tmp_path, PlantMetadataCatalog([]))
    for plants in (1, 2):
        generation = pd.concat(
            [
                pd.DataFrame(
                    {
                        "일시": timestamps,
                        "발전구분": f"테스트태양광{index}",
                        "지역": "서울",
                        "지점번호": 108,
                        "합산발전량(MWh)": np.maximum(
                            np.sin((timestamps.hour.to_numpy() - 6) * np.pi / 12), 0
                        ),
                    }
                )
                for index in range(plants)
            ],
            ignore_index=True,
        )
        source = tmp_path / "legacy_mapping.csv"
        generation.to_csv(source, index=False)
        result = builder.build(source, tmp_path / "gold.csv")
        manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
        evidence = manifest["feature_selection_evidence"]

        assert manifest["plants"] == plants
        assert manifest["rows"] == plants * len(timestamps)
        assert evidence["status"] == "not_validated_for_current_dataset"
        assert evidence["dataset_match_verified"] is False
        assert evidence["experiment_artifact"] is None
        assert evidence["metrics"] is None
        assert evidence["historical_reference"]["current_dataset_evidence"] is False
        assert "baseline_23_mean_mae" not in evidence
        assert "selected_26_mean_mae" not in evidence
        assert "relative_mean_mae_improvement" not in evidence
