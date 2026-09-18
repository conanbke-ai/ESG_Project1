"""Dashboard commands: CLI input translation and command dispatch."""
from __future__ import annotations

import argparse
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from solar_forecast.config_loader import PROJECT_ROOT


def handle_build_dashboard_command(args: argparse.Namespace) -> None:
    from solar_forecast.reporting import DashboardBuilder

    result = DashboardBuilder(PROJECT_ROOT, Path(args.output_dir)).build()
    print(
        f"National inventory refreshed: {result.national_generator_records:,} generator records, "
        f"{result.national_capacity_mw:,.2f} MW"
    )
    print(
        f"Model analytics: {result.model_analysis_status}, "
        f"{result.data_quality_signals} solar data signals"
    )
    print(f"Solar dashboard: {result.solar_dashboard}")
    print(f"Forecast dashboard: {result.forecast_dashboard}")
    print(f"Model analytics dashboard: {result.analytics_dashboard}")


def handle_prepare_boundaries_command(args: argparse.Namespace) -> None:
    from solar_forecast.reporting import SgisBoundarySource, SgisProvinceBoundaryConverter

    converter = SgisProvinceBoundaryConverter(
        simplify_meters=args.simplify_meters,
        precision=args.precision,
    )
    payload = converter.convert(
        Path(args.source_shp),
        Path(args.output),
        source_archive_path=Path(args.source_archive),
        source=SgisBoundarySource(
            provider=args.provider,
            source_url=args.source_url,
            reference_date=args.reference_date,
            archive_sha256=args.archive_sha256,
            shapefile_sha256=args.shapefile_sha256,
        ),
    )
    print(
        f"Official province boundaries: {len(payload['features'])} regions -> "
        f"{Path(args.output)}"
    )


def handle_serve_dashboard_command(args: argparse.Namespace) -> None:
    from solar_forecast.reporting import DashboardBuilder

    output_dir = Path(args.output_dir)
    if not output_dir.is_absolute():
        output_dir = PROJECT_ROOT / output_dir
    if not args.no_refresh:
        DashboardBuilder(PROJECT_ROOT, output_dir).build()
    entrypoint = output_dir / "solar_dashboard.html"
    if not entrypoint.exists():
        raise FileNotFoundError(f"Dashboard entrypoint is missing: {entrypoint}")
    handler = partial(SimpleHTTPRequestHandler, directory=str(output_dir))
    server = ThreadingHTTPServer((args.host, args.port), handler)
    url = f"http://{args.host}:{server.server_port}/solar_dashboard.html"
    print(f"Dashboard server: {url}", flush=True)
    print("Stop with Ctrl+C", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
