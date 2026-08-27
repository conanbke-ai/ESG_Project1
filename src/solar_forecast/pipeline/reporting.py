from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Any

import pandas as pd


class HtmlReportWriter:
    def write(self, metrics: dict[str, Any], anomalies: pd.DataFrame, source: Path, output_path: Path) -> Path:
        return build_html_report(metrics, anomalies, source, output_path)


def build_html_report(metrics: dict[str, Any], anomalies: pd.DataFrame, source: Path, output_path: Path) -> Path:
    metric_rows = "".join(f"<tr><th>{escape(str(k))}</th><td>{float(v):.6f}</td></tr>" for k, v in metrics.items())
    preview = anomalies.head(100).to_html(index=False, classes="data", border=0)
    html = f"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>태양광 AI 분석 결과</title><style>
body{{font-family:Arial,sans-serif;max-width:1100px;margin:40px auto;padding:0 20px;color:#1f2937}}
h1{{color:#166534}}table{{border-collapse:collapse;width:100%;margin:16px 0}}th,td{{padding:9px;border-bottom:1px solid #ddd;text-align:right}}th{{background:#f0fdf4}}.meta{{color:#64748b}}
</style></head><body><h1>태양광 AI 분석 결과</h1><p class="meta">입력 파일: {escape(str(source))}</p>
<h2>평가 지표</h2><table>{metric_rows}</table>
<h2>예측 및 이상징후 (최대 100건)</h2>{preview}
<p><strong>해석 한계:</strong> 공개 데이터만으로 설비 고장·정비·출력제어 여부를 확인할 수 없습니다.</p>
</body></html>"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    return output_path
