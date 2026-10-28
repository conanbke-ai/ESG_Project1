"""기존 Gold로 사용자 로컬 GPU 학습·저장 모델 재예측·화면 생성을 실행한다.

패키지 설치, 데이터 재생성, CI 호출, 탐색 예산 변경은 수행하지 않는다.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import importlib
import json
import os
from pathlib import Path
import sys
import uuid


PROJECT_ROOT = Path(__file__).resolve().parents[1]
GOLD_CONTRACT = "solar-model-ready-manifest.v1"
PREPROCESSING_CONTRACT = "solar-observed-preprocessing.v2"


def _write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".part")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    os.replace(temporary, path)


def _resolve(path: Path | str, project_root: Path) -> Path:
    value = Path(path).expanduser()
    return (value if value.is_absolute() else project_root / value).resolve()


def _read_inputs(config_path: Path, data_path: Path | None, project_root: Path) -> tuple[dict, Path, dict]:
    """Validate existing Gold metadata without importing model frameworks."""
    values = json.loads(config_path.read_text(encoding="utf-8"))
    source = _resolve(data_path or values["input_dataset"], project_root)
    if not source.exists():
        raise FileNotFoundError(f"기존 Gold가 없습니다: {source}. --data로 실제 경로를 지정하세요.")
    if not source.is_dir() and not source.name.lower().endswith((".csv", ".csv.gz")):
        raise ValueError("Gold는 CSV, CSV.GZ 또는 해당 파일의 파티션 폴더여야 합니다.")
    manifest_path = source.parent / "model_ready_manifest.json"
    if not manifest_path.is_file():
        raise FileNotFoundError(f"Gold 전처리 계약을 확인할 manifest가 없습니다: {manifest_path}")
    manifest_bytes = manifest_path.read_bytes()
    manifest = json.loads(manifest_bytes)
    if manifest.get("contract") != GOLD_CONTRACT or manifest.get("schema_version") != 1:
        raise ValueError("지원하는 Gold manifest 계약이 아닙니다.")
    if manifest.get("preprocessing_contract") != PREPROCESSING_CONTRACT:
        raise ValueError("기존 Gold에 solar-observed-preprocessing.v2가 필요합니다. 데이터는 자동 재생성하지 않습니다.")
    # Preserve the selected experiment's budget and boundaries. This is an
    # execution copy; original files and historical manifest paths stay intact.
    values["input_dataset"] = str(source)
    metadata = {
        "source": str(source), "manifest_path": str(manifest_path),
        "manifest_sha256": hashlib.sha256(manifest_bytes).hexdigest(),
        "preprocessing_contract": manifest["preprocessing_contract"],
    }
    return values, source, metadata


def _configure_environment(threads: int) -> None:
    if threads < 1:
        raise ValueError("--threads는 1 이상이어야 합니다.")
    for name in ("OMP_NUM_THREADS", "MKL_NUM_THREADS", "OPENBLAS_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
        os.environ[name] = str(threads)


def _load_dependencies(threads: int, *, require_cuda: bool = True) -> dict:
    """Import the installed numeric runtime; missing or broken binaries fail early."""
    packages = {}
    for name in ("numpy", "pandas", "sklearn", "torch", "xgboost", "optuna"):
        try:
            module = importlib.import_module(name)
        except Exception as exc:
            raise RuntimeError(f"학습 의존성을 불러올 수 없습니다: {name} ({type(exc).__name__}: {exc}). 설치된 가상환경의 Python으로 실행하세요.") from exc
        packages[name] = str(getattr(module, "__version__", "unknown"))
    torch = importlib.import_module("torch")
    cuda_available = bool(torch.cuda.is_available())
    if require_cuda and not cuda_available:
        raise RuntimeError("본 학습은 사용자 로컬 GPU에서 실행해야 합니다. 이 Python 환경에서 CUDA GPU를 사용할 수 없어 중단합니다. CPU로 대체 학습하지 않습니다.")
    torch.set_num_threads(threads)
    torch.set_num_interop_threads(1)
    return {"python": sys.version.split()[0], "packages": packages,
            "torch_intraop_threads": torch.get_num_threads(),
            "torch_interop_threads": torch.get_num_interop_threads(),
            "cuda_available": cuda_available,
            "gpu_name": torch.cuda.get_device_name(0) if cuda_available else None,
            "training_device_policy": "local_cuda_required" if require_cuda else "training_not_requested",
            "cnn_training_device": "cuda" if require_cuda else None,
            "xgboost_training_device": "cpu" if require_cuda else None,
            "replay_device": "cpu"}


def run_local_benchmark(args: argparse.Namespace, *, project_root: Path = PROJECT_ROOT) -> tuple[int, Path]:
    """Run existing services; return process exit code and durable stage-report path."""
    project_root = project_root.resolve()
    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%f") + "_" + uuid.uuid4().hex[:8]
    work_dir = project_root / "artifacts/verification/local_benchmark" / run_id
    report_path = work_dir / "status.json"
    report = {"contract": "solar-local-benchmark-run.v1", "status": "running",
              "stage": "input_validation", "training_invoked": False, "replay_invoked": False,
              "preflight_only": bool(args.preflight_only), "ci_invoked": False,
              "packages_installed": False, "data_regenerated": False}
    try:
        _configure_environment(args.threads)
        replay_run = _resolve(args.replay_run, project_root) if args.replay_run else None
        config_path = replay_run / "experiment.json" if replay_run else _resolve(args.config, project_root)
        values, source, gold = _read_inputs(config_path, args.data, project_root)
        if source.is_dir() and source in work_dir.parents:
            raise ValueError("실행 산출물을 Gold 입력 폴더 안에 기록할 수 없습니다.")
        # The current dashboard discovers this exact artifact root. Reject an
        # unsupported output route before spending time on model training.
        benchmark_root = project_root / "artifacts/benchmarks"
        if not replay_run and _resolve(values.get("output_root", "artifacts/benchmarks"), project_root) != benchmark_root:
            raise ValueError("대시보드 연결을 위해 실험 output_root는 artifacts/benchmarks여야 합니다.")
        if replay_run and replay_run.parent != benchmark_root:
            raise ValueError("재예측 후 대시보드 갱신 대상은 artifacts/benchmarks 아래의 실행 폴더여야 합니다.")
        report.update(gold=gold, source_config=str(config_path), stage="dependencies")
        _write_json(report_path, report)
        report["runtime"] = _load_dependencies(args.threads, require_cuda=replay_run is None)
        sys.path.insert(0, str(project_root / "src"))
        sys.path.insert(0, str(project_root / "tools"))
        resolved_path = work_dir / "experiment.json"
        _write_json(resolved_path, values)
        report["resolved_config"] = str(resolved_path)
        if not replay_run:
            from solar_forecast.evaluation.forecast_readiness import run_forecast_readiness
            report["stage"] = "readiness"
            _write_json(report_path, report)
            readiness_path = work_dir / "forecast_readiness.json"
            readiness = run_forecast_readiness(resolved_path, data_path=source, output_path=readiness_path, project_root=project_root)
            report["readiness_report"] = str(readiness_path)
            if not readiness.get("coverage_gate_passed"):
                raise ValueError("모델별 예측 표본의 공통 커버리지 기준을 충족하지 못했습니다.")
            if args.preflight_only:
                report.update(status="preflight_passed", stage="completed",
                              prediction_or_training_performed=False,
                              accuracy_validated=False)
                _write_json(report_path, report)
                print(f"사전 점검 통과. 학습·예측은 실행하지 않았습니다. 상태: {report_path}", flush=True)
                return 0, report_path
            from solar_forecast.jobs.benchmark_job import BenchmarkService
            report.update(stage="training", training_invoked=True)
            _write_json(report_path, report)
            run_dir = BenchmarkService(project_root=project_root).run(resolved_path, smoke=False)
        else:
            run_dir = replay_run
        # Persist the exact returned path before replay, so a failed replay or
        # dashboard build can be retried with --replay-run without another fit.
        run_dir = Path(run_dir).resolve()
        report.update(benchmark_run_dir=str(run_dir), stage="saved_model_replay", replay_invoked=True)
        _write_json(report_path, report)
        from verify_benchmark_model_artifacts import verify_selected_artifacts
        replay_report = verify_selected_artifacts(run_dir, source)
        replay_path = work_dir / "stored_model_replay.json"
        _write_json(replay_path, replay_report)
        report["replay_report"] = str(replay_path)
        if replay_report.get("status") != "passed":
            raise ValueError(f"저장 모델 재예측 검증에 실패했습니다: {replay_path}")
        from solar_forecast.reporting.dashboard_builder import DashboardBuilder
        report["stage"] = "dashboard"
        _write_json(report_path, report)
        dashboard = DashboardBuilder(project_root).build()
        payload = json.loads(dashboard.data_path.read_text(encoding="utf-8"))
        if payload.get("model_benchmark", {}).get("run_id") != run_dir.name:
            raise ValueError("대시보드가 이번 검증 대상 실행을 표시하지 않습니다. 학습 결과는 보존했습니다.")
        report.update(status="completed", stage="completed", dashboard_data=str(dashboard.data_path))
        _write_json(report_path, report)
        print(f"학습 결과·저장 모델 재예측·화면 데이터 확인 완료: {run_dir}\n상태: {report_path}", flush=True)
        return 0, report_path
    except (Exception, KeyboardInterrupt) as exc:
        report.update(status="interrupted" if isinstance(exc, KeyboardInterrupt) else "failed",
                      error={"type": type(exc).__name__, "message": str(exc)})
        _write_json(report_path, report)
        print(f"실행 중단 ({report['stage']}): {type(exc).__name__}: {exc}\n상태: {report_path}", file=sys.stderr, flush=True)
        if report.get("benchmark_run_dir"):
            print(f"학습을 반복하지 않고 --replay-run \"{report['benchmark_run_dir']}\"으로 재검증할 수 있습니다.", file=sys.stderr)
        return (130 if isinstance(exc, KeyboardInterrupt) else 1), report_path


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=Path("config/experiments/observed_calendar_candidate.json"))
    parser.add_argument("--data", type=Path, help="기존 Gold CSV/CSV.GZ 또는 파티션 폴더")
    parser.add_argument("--threads", type=int, default=4, help="CPU 수치 연산 스레드 수 (기본 4)")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--preflight-only", action="store_true", help="의존성·Gold·표본만 점검; 학습·예측하지 않음")
    mode.add_argument("--replay-run", type=Path, help="완료 실행의 experiment.json으로 저장 모델만 재예측하고 화면 갱신")
    args = parser.parse_args()
    code, _ = run_local_benchmark(args)
    raise SystemExit(code)


if __name__ == "__main__":
    main()
