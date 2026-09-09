# 코드 색인

`python tools/update_code_index.py`로 재생성합니다. 파일 역할은 `config/architecture/modules.json`에서 관리합니다.
클래스·메서드·내부 함수까지 실제 AST에서 추출합니다. 함수 이름 뒤 괄호는 입력 인자이며, 설명이 있는 항목은 함수 docstring을 함께 표시합니다.

## [__init__.py](../src/solar_forecast/__init__.py)

__init__ 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [anomalies/__init__.py](../src/solar_forecast/anomalies/__init__.py)

anomalies 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [anomalies/event_batch.py](../src/solar_forecast/anomalies/event_batch.py)

운영 이상치 이벤트 ID 생성, JSONL 무결성 manifest 기록과 검증

- `OperationalEventBatch` — A verified, replayable operational anomaly-event stream.
- `OperationalEventBatch.records(self)` — Read one record at a time after the whole file passed integrity checks.
- `stable_operational_event_id(*, detector_id: str, plant_id: str, observed_at: datetime, detector_version: str, influence_factor: str, signal_type: str, detector_role: str)` — Create the stable event key shared by producers and notification consumers.
- `operational_event_id(record: Mapping[str, Any])` — Derive an event ID from either the flat or nested v1 event representation.
- `write_operational_event_batch(records: Iterable[Mapping[str, Any]], output_dir: Path, *, run_id: str, detector_version: str)` — Atomically write JSONL plus a count/hash manifest without buffering all events.
- `verify_operational_event_batch(events_path: Path, manifest_path: Path | None=None)` — Fail closed before enqueueing if lineage, count, contract or JSONL is wrong.
- `_iter_records(path: Path, *, expected_run_id: str | None=None, expected_detector_version: str | None=None)`
- `_validate_event_shell(record: Mapping[str, Any], *, line_number: int)`
- `_find_forbidden_keys(value: Any)`
- `_required_text(name: str, value: Any)`

## [anomalies/influence_policy.py](../src/solar_forecast/anomalies/influence_policy.py)

이상치 설명에 허용되는 영향 요인과 금지 항목 검증

- `validate_influence_factor(value: str)` — Reject unsupported root-cause labels before they reach reports or alerts.

## [cli/__init__.py](../src/solar_forecast/cli/__init__.py)

cli 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `main(argv: Sequence[str] | None=None)`

## [cli/__main__.py](../src/solar_forecast/cli/__main__.py)

python -m solar_forecast.cli 모듈 실행 진입점

공개 export 또는 설정 상수만 정의합니다.

## [cli/arguments.py](../src/solar_forecast/cli/arguments.py)

쉼표 목록 파싱과 CNN 시퀀스 공통 CLI 인자 구성

- `parse_csv_values(value: str | None)`
- `build_sequence_config(args: argparse.Namespace)`
- `add_sequence_arguments(parser: argparse.ArgumentParser)`

## [cli/collection_commands.py](../src/solar_forecast/cli/collection_commands.py)

collect 인자를 수집 설정으로 바꾸고 공급자 수집 실행

- `handle_collect_command(args: argparse.Namespace)`

## [cli/dashboard_commands.py](../src/solar_forecast/cli/dashboard_commands.py)

대시보드 생성·서빙과 공식 지도 경계 변환 명령 처리

- `handle_build_dashboard_command(args: argparse.Namespace)`
- `handle_prepare_boundaries_command(args: argparse.Namespace)`
- `handle_serve_dashboard_command(args: argparse.Namespace)`

## [cli/dataset_commands.py](../src/solar_forecast/cli/dataset_commands.py)

prepare-data와 후보 데이터 심사 명령 처리

- `handle_prepare_data_command(args: argparse.Namespace)`
- `handle_audit_candidate_data_command(args: argparse.Namespace)`

## [cli/job_commands.py](../src/solar_forecast/cli/job_commands.py)

실행 상태와 job 입출력 계약 조회 명령 처리

- `handle_status_command(_: argparse.Namespace)`
- `handle_jobs_command(args: argparse.Namespace)`
- `handle_job_contract_command(args: argparse.Namespace)`

## [cli/notification_commands.py](../src/solar_forecast/cli/notification_commands.py)

검증된 이상치 이벤트의 알림 dispatcher 명령 처리

- `handle_notify_anomalies_command(args: argparse.Namespace)`

## [cli/parser.py](../src/solar_forecast/cli/parser.py)

공개 CLI 명령과 옵션 선언, 명령 처리 함수 연결

- `build_parser()`

## [cli/training_commands.py](../src/solar_forecast/cli/training_commands.py)

학습, 하이브리드 평가, 특징 비교, 전체 파이프라인 명령 처리

- `handle_benchmark_command(args: argparse.Namespace)`
- `handle_pipeline_command(args: argparse.Namespace)`
- `handle_hybrid_command(args: argparse.Namespace)`
- `handle_train_command(args: argparse.Namespace)`
- `handle_evaluate_features_command(args: argparse.Namespace)`

## [cli/verification_commands.py](../src/solar_forecast/cli/verification_commands.py)

연결 검증과 저장 예측 CSV의 수치 동등성 비교 명령 처리

- `handle_compare_predictions_command(args: argparse.Namespace)` — Write parity evidence for two saved CSVs without modifying either input.
- `handle_verify_e2e_command(args: argparse.Namespace)`

## [collectors/__init__.py](../src/solar_forecast/collectors/__init__.py)

collectors 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [collectors/collection_config.py](../src/solar_forecast/collectors/collection_config.py)

수집 기간·공급자·저장 경로·ASOS 모드·호출 한도 검증

- `CollectionConfig`
- `CollectionConfig.__post_init__(self)`
- `load_source_catalog()`

## [collectors/collection_service.py](../src/solar_forecast/collectors/collection_service.py)

공급자별 실행 격리와 전체 collection manifest 기록

- `CollectionService` — Application service that isolates collectors and owns run artifacts.
- `CollectionService.__init__(self, config: CollectionConfig)`
- `CollectionService.run(self)`
- `CollectionService._collect(self, source: str, run_dir: Path)`
- `collect_all(config: CollectionConfig)` — Compatibility facade; new code should instantiate CollectionService.

## [collectors/collector_factory.py](../src/solar_forecast/collectors/collector_factory.py)

선택한 공급자 adapter만 로딩하고 ASOS API/browser 모드 결정

- `build_collector(source: str, config: CollectionConfig)` — Resolve one source, selecting ASOS API mode after local environment loading.

## [collectors/contracts.py](../src/solar_forecast/collectors/contracts.py)

수집기 인터페이스와 수집 결과 데이터 계약

- `CollectionResult`
- `Collector`
- `Collector.collect(self)`

## [collectors/download_naming.py](../src/solar_forecast/collectors/download_naming.py)

기관·자료명·수집일을 포함하는 표준 다운로드 파일명 생성

- `_clean_component(value: str)`
- `SolarDownloadName` — Canonical name for an immutable provider download (Bronze artifact).
- `SolarDownloadName.filename(self)`
- `build_solar_download_filename(organization: str, detail: str, downloaded_on: date)`
- `is_canonical_solar_download_name(path: str | Path)`

## [collectors/generation_collectors.py](../src/solar_forecast/collectors/generation_collectors.py)

KOEN·KOSPO·EWP·IWEST 공식 발전량 파일 다운로드 adapter

- `FileNormalizer`
- `FileNormalizer.write(self, source: Path, destination: Path)`
- `_csv_data_rows(path: Path)`
- `EwpAttachmentSpec`
- `EwpTrainingDataCollector` — Download EWP's nationwide solar training CSV from its public-data board.
- `EwpTrainingDataCollector.__init__(self, config: CollectionConfig, spec: EwpAttachmentSpec)`
- `EwpTrainingDataCollector.collect(self)`
- `DataGoDatasetSpec`
- `DataGoDatasetSpec.detail_url(self)`
- `DataGoFileCollector` — Download a login-free official attachment through data.go.kr's public flow.
- `DataGoFileCollector.__init__(self, spec: DataGoDatasetSpec, config: CollectionConfig, normalizer: FileNormalizer | None=None)`
- `DataGoFileCollector.collect(self)`
- `KoenHomepageCollector` — Reuse the existing official KOEN monthly-download browser automation.
- `KoenHomepageCollector.__init__(self, config: CollectionConfig)`
- `KoenHomepageCollector.collect(self)`

## [collectors/generation_normalizers.py](../src/solar_forecast/collectors/generation_normalizers.py)

공급자별 발전량 열·시간·단위를 plant-hour Silver 스키마로 변환

- `classify_energy_source(value: object)` — Map public Korean generator labels to a stable technology contract.
- `read_csv_with_fallback(path: Path, *, index_col: bool | None=None, **read_csv_kwargs: object)` — Read public CSV exports without coupling callers to one Korean encoding.
- `_empty_static_columns(frame: pd.DataFrame, source_file: str | None)`
- `KoenGenerationNormalizer` — Convert KOEN's daily wide CSV into one generation observation per hour.
- `KoenGenerationNormalizer.read(self, path: Path)`
- `KoenGenerationNormalizer.transform(self, frame: pd.DataFrame, *, source_file: str | None=None)`
- `KoenGenerationNormalizer.write(self, source: Path, destination: Path)`
- `EwpTrainingNormalizer` — Validate and standardize EWP's nationwide hourly solar training file.
- `EwpTrainingNormalizer.read(self, path: Path)`
- `EwpTrainingNormalizer.transform(self, frame: pd.DataFrame)`
- `EwpTrainingNormalizer.write(self, source: Path, destination: Path)`
- `DailyWideSchema`
- `DailyWideGenerationNormalizer` — Normalize one-row-per-day public generation files to hourly MWh.
- `DailyWideGenerationNormalizer.__init__(self, schema: DailyWideSchema)`
- `DailyWideGenerationNormalizer._capacity_factor_metrics(cls, source_values: pd.Series, capacity_mw: pd.Series, source_unit: str)`
- `DailyWideGenerationNormalizer._resolve_source_unit(self, source_values: pd.Series, capacity_mw: pd.Series)`
- `DailyWideGenerationNormalizer._reconcile_daily_totals(self, source: pd.DataFrame, resolved_hourly_unit: str)`
- `DailyWideGenerationNormalizer.read(self, path: Path)`
- `DailyWideGenerationNormalizer.transform(self, frame: pd.DataFrame, *, source_file: str | None=None)`
- `DailyWideGenerationNormalizer.write(self, source: Path, destination: Path)`
- `KrcYeongamGenerationNormalizer` — Normalize the Korean Rural Community Corporation Yeongam exports.
- `KrcYeongamGenerationNormalizer.read(self, path: Path)`
- `KrcYeongamGenerationNormalizer.transform(self, frame: pd.DataFrame, *, year: int, source_file: str | None=None)`

## [collectors/kma_api.py](../src/solar_forecast/collectors/kma_api.py)

ASOS 관측소·월별 수집과 응답 checkpoint 재사용, 연도별 기상 CSV 게시

- `iter_asos_months(start: date, end: date)` — Yield closed calendar-month intervals bounded by the requested date range.
- `validate_asos_observations(items: list[dict], station: str, start: date, end: date)` — Reject wrong stations, out-of-range/non-hourly times, and duplicate keys.
- `KmaAsosApiCollector` — Collect observations, then publish validated station/hour weather partitions.
- `KmaAsosApiCollector.__init__(self, config: CollectionConfig, *, client: AsosHourlyApiClient | None=None)`
- `KmaAsosApiCollector.collect(self)`
- `KmaAsosApiCollector._collect_partition(self, client: AsosHourlyApiClient, station: str, start: date, end: date)`
- `KmaAsosApiCollector._read_partition(manifest_path: Path, request: dict, station: str, start: date, end: date)`

## [collectors/kma_api_client.py](../src/solar_forecast/collectors/kma_api_client.py)

ASOS 인증키를 사용하는 HTTP 요청·페이지 검증·호출 예산 관리

- `AsosApiError` — A safe error whose message never contains credentials or response bodies.
- `_RejectRedirects`
- `_RejectRedirects.redirect_request(self, req, fp, code, msg, headers, newurl)`
- `fetch_asos_response(url: str, timeout: float)` — Read one response without forwarding the query-string key to redirects.
- `AsosPage`
- `parse_asos_page(raw_bytes: bytes, page_number: int)` — Validate JSON pagination, including the provider's no-data result code.
- `AsosHourlyApiClient` — Use one explicit call budget across all stations and monthly partitions.
- `AsosHourlyApiClient.__init__(self, service_key: str, *, max_calls: int, timeout: float=30, transport: Callable[[str, float], bytes]=fetch_asos_response)`
- `AsosHourlyApiClient.fetch_page(self, station_id: str, start: date, end: date, page_number: int)`

## [collectors/kma_browser.py](../src/solar_forecast/collectors/kma_browser.py)

기상청 포털의 기존 브라우저 수집과 공통 기상 CSV 병합

- `KmaAsosBrowserCollector` — Incrementally download ASOS hourly CSV files from the KMA data portal UI.
- `KmaAsosBrowserCollector.__init__(self, config: CollectionConfig)`
- `KmaAsosBrowserCollector._latest_available_date()`
- `KmaAsosBrowserCollector._existing_files(self)`
- `KmaAsosBrowserCollector._file_range(path: Path)`
- `KmaAsosBrowserCollector._latest_existing_date(self)`
- `KmaAsosBrowserCollector._chunks(start: date, end: date)` — ASOS hourly UI permits at most one year per query.
- `KmaAsosBrowserCollector._build_driver(self, download_dir: Path)`
- `KmaAsosBrowserCollector._set_search_form(driver, start: date, end: date)`
- `KmaAsosBrowserCollector._click_text(driver, text: str)`
- `KmaAsosBrowserCollector._wait_for_download(download_dir: Path, before: set[Path], timeout: int=180)`
- `KmaAsosBrowserCollector._download_chunks(self, start: date, end: date, download_dir: Path)`
- `KmaAsosBrowserCollector._read_csv(path: Path)`
- `KmaAsosBrowserCollector._merge_by_year(self, downloads: list[Path])`
- `KmaAsosBrowserCollector.collect(self)`
- `KmaAsosBrowserCollector._collect_browser_observations(self)`

## [collectors/koen_browser.py](../src/solar_forecast/collectors/koen_browser.py)

남동발전 홈페이지 화면 조작과 공식 첨부파일 다운로드

- `KoenBrowserDownloader` — Selenium adapter for KOEN's monthly generation download screen.
- `KoenBrowserDownloader.__init__(self, output_root: Path, *, downloaded_on: date | None=None, overwrite: bool=False)`
- `KoenBrowserDownloader.download_year(self, year: int, *, start_month: int=1, end_month: int=12, max_retry: int=3)`
- `KoenBrowserDownloader._build_driver(self)`
- `KoenBrowserDownloader._open(self, driver)`
- `KoenBrowserDownloader._download_month(self, driver, year_dir: Path, year: int, month: int, max_retry: int)`
- `KoenBrowserDownloader._wait_for_csv(self, before: set[Path], timeout: int=30)`
- `KoenBrowserDownloader._move_original(self, source: Path, destination: Path)` — Move provider bytes unchanged; encoding conversion belongs to Silver.
- `download_solar_data(base_path: str, year: int, max_retry: int=3, start_month: int=1, end_month: int=12)` — Compatibility facade for the former root-level automation function.

## [collectors/komipo_api.py](../src/solar_forecast/collectors/komipo_api.py)

중부발전 전용 DATA_GO_SERVICE_KEY 기반 발전량 OpenAPI 수집

- `KomipoRenewableCollector` — Incrementally stage KOMIPO renewable measurements for admission review.
- `KomipoRenewableCollector.__init__(self, config: CollectionConfig, session: requests.Session | None=None)`
- `KomipoRenewableCollector.collect(self)`
- `KomipoRenewableCollector._fetch_day(self, service_key: str, station_code: str, current: date)`
- `KomipoRenewableCollector._destination(self, station_code: str, current: date)`
- `KomipoRenewableCollector._dates(start: date, end: date)`
- `KomipoRenewableCollector._text(root: ET.Element, path: str)`

## [collectors/plant_identity.py](../src/solar_forecast/collectors/plant_identity.py)

공식 근거로 검토한 발전소·발전기 identity 교정과 충돌 거부

- `resolve_generation_identity(frame: pd.DataFrame)` — Resolve only the reviewed department/generator pair, without guessing.
- `read_generation_partition(path: Path, columns: Iterable[str])` — Read the optional generator label before projecting registry/Gold columns.

## [collectors/plant_metadata.py](../src/solar_forecast/collectors/plant_metadata.py)

공식 발전소 메타데이터의 이름·설비용량·주소 해석과 catalog

- `canonical_plant_name(value: object)` — Return a comparison key while preserving site/unit distinctions.
- `_unit_number(value: object)`
- `_first_number(value: object)`
- `_capacity_mw(value: object, default_unit: str)`
- `PlantMetadata`
- `PlantMetadata.key(self)`
- `PlantMetadataCatalog` — Normalize and match the four companies' public plant metadata tables.
- `PlantMetadataCatalog.__init__(self, records: Iterable[PlantMetadata])`
- `PlantMetadataCatalog.canonical_plant(cls, company: str, plant: str)` — Return the reviewed physical-asset name used across snapshots.
- `PlantMetadataCatalog.from_directory(cls, directory: Path)`
- `PlantMetadataCatalog._kospo(frame: pd.DataFrame)`
- `PlantMetadataCatalog._standard(frame: pd.DataFrame, company: str, capacity_column: str, capacity_unit: str, address_column: str, energy_source_column: str | None=None)`
- `PlantMetadataCatalog.lookup(self, company: str, plant: str, unit: str | None=None, energy_source: str | None=None, *, aggregate: bool=False)`
- `PlantMetadataCatalog.enrich(self, frame: pd.DataFrame, *, aggregate: bool=False)`

## [config_loader.py](../src/solar_forecast/config_loader.py)

모델 JSON 설정 검증과 프로젝트 루트 해석

- `ModelJobConfig`
- `load_model_config(path: Path)`

## [datasets/__init__.py](../src/solar_forecast/datasets/__init__.py)

datasets 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [datasets/archive_standardizer.py](../src/solar_forecast/datasets/archive_standardizer.py)

보관 원본의 공급자 스키마 판별과 표준 발전량 파티션 생성

- `StandardizedPartition`
- `StandardizationRun`
- `StandardizationRun.rows(self)`
- `GenerationSchemaRegistry` — Detect a known public CSV layout by its declared column contract.
- `GenerationSchemaRegistry.__init__(self, schemas: tuple[DailyWideSchema, ...]=ARCHIVE_WIDE_SCHEMAS)`
- `GenerationSchemaRegistry._required(schema: DailyWideSchema)`
- `GenerationSchemaRegistry.resolve(self, columns: set[str], company: str)`
- `HistoricalGenerationStandardizationService` — Convert every retained generation export into partitioned common-schema data.
- `HistoricalGenerationStandardizationService.__init__(self, input_root: Path, output_dir: Path, registry: GenerationSchemaRegistry | None=None, metadata: PlantMetadataCatalog | None=None)`
- `HistoricalGenerationStandardizationService.run(self)`
- `HistoricalGenerationStandardizationService._normalize(self, company: str, source: Path)`
- `HistoricalGenerationStandardizationService._audit(company: str, source: Path, destination: Path, frame: pd.DataFrame)`

## [datasets/asos_weather_store.py](../src/solar_forecast/datasets/asos_weather_store.py)

API·브라우저 기상 자료의 관측소·시간 키 병합과 동시 쓰기 잠금

- `exclusive_asos_collection(weather_root: Path)` — Serialize API collection and annual-file publication across local workers.
- `observation_to_weather_row(item: dict)` — Preserve QC flags and missing values while translating provider field names.
- `_read_annual_weather(path: Path)`
- `_weather_row_key(row: dict[str, str])`
- `merge_asos_observations(weather_root: Path, items: list[dict])` — Upsert station/hour keys, retaining other stations and unknown CSV columns.
- `merge_weather_rows(weather_root: Path, rows: list[dict[str, str]])` — Publish rows from either API or browser using the same keyed column merge.

## [datasets/collector_admission.py](../src/solar_forecast/datasets/collector_admission.py)

collector Silver 파일의 plant-hour 스키마 심사와 admission manifest

- `CollectedGenerationFileAdmission`
- `CollectedGenerationAdmissionResult`
- `CollectedGenerationAdmissionResult.accepted_paths(self)`
- `CollectedGenerationAdmissionResult.accepted_count(self)`
- `CollectedGenerationAdmissionResult.rejected_count(self)`
- `CollectedGenerationAdmissionResult.accepted_rows(self)`
- `CollectedGenerationAdmissionService` — Promote collector Silver files only when they satisfy the Gold contract.
- `CollectedGenerationAdmissionService.__init__(self, source_dir: Path, manifest_path: Path)`
- `CollectedGenerationAdmissionService.run(self)`
- `CollectedGenerationAdmissionService._inspect(self, path: Path)`
- `CollectedGenerationAdmissionService._read_generation_columns(path: Path)`
- `collect_generation_admissions(source_dir: Path, manifest_path: Path)`

## [datasets/krc_candidate_intake.py](../src/solar_forecast/datasets/krc_candidate_intake.py)

농어촌공사 영암 후보 파일 정규화·발전소별 품질 심사

- `CandidateAcceptancePolicy`
- `CandidateSourceFile`
- `CandidatePlantProfile`
- `CandidateIntakeResult`
- `KrcYeongamCandidateIntakeService` — Stage, standardize, and gate KRC Yeongam files before model admission.
- `KrcYeongamCandidateIntakeService.__init__(self, source_dir: Path, output_dir: Path, policy: CandidateAcceptancePolicy | None=None)`
- `KrcYeongamCandidateIntakeService.run(self)`
- `KrcYeongamCandidateIntakeService._profile_plant(self, plant_id: str, frame: pd.DataFrame, split_frames: dict[str, pd.DataFrame])`
- `KrcYeongamCandidateIntakeService._write_csv_atomic(frame: pd.DataFrame, destination: Path)`

## [datasets/model_dataset_builder.py](../src/solar_forecast/datasets/model_dataset_builder.py)

registry 승인 발전량과 ASOS·과거 특징을 결합해 Gold dataset 생성

- `ModelDatasetResult`
- `NationwideModelDatasetBuilder` — Build one common feature table from every qualified public plant partition.
- `NationwideModelDatasetBuilder.__init__(self, weather_root: Path, metadata: PlantMetadataCatalog, engineer: LeakageSafeFeatureEngineer | None=None, quality_policy: GenerationQualityPolicy | None=None)`
- `NationwideModelDatasetBuilder.read_legacy_generation(self, source_path: Path)` — Read the retained merge only as a plant/station mapping and audit source.
- `NationwideModelDatasetBuilder._read_optional_legacy_generation(self, source_path: Path)` — Use a retained legacy mapping only when the audit artifact exists.
- `NationwideModelDatasetBuilder.build(self, source_path: Path, destination: Path, *, generation_paths: Iterable[Path] | None=None)`
- `NationwideModelDatasetBuilder._from_standardized_generation(self, paths: Iterable[Path], registry: pd.DataFrame)`
- `NationwideModelDatasetBuilder._source_role(path: Path)`
- `NationwideModelDatasetBuilder._write_model_partitions(frame: pd.DataFrame, destination: Path)`

## [datasets/numeric_preprocessor.py](../src/solar_forecast/datasets/numeric_preprocessor.py)

학습 품질 필터·숫자 열 전처리·Train에서만 학습하는 결측 대체

- `require_model_quality_filter(values: Mapping[str, object])` — Fail closed when a forecasting job tries to bypass the Gold quality gate.
- `PreprocessResult`
- `NumericPreprocessor` — Stateless preprocessing policy with an explicit feature contract.
- `NumericPreprocessor.__init__(self, *, fill_missing: bool=True)`
- `NumericPreprocessor.transform(self, frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None, passthrough_columns: Optional[Sequence[str]]=None, quality_filter_column: str | None=None)`
- `TrainFittedMedianImputer` — Fit feature medians on training rows only and reuse them unchanged.
- `TrainFittedMedianImputer.__init__(self, medians: Mapping[str, float] | None=None)`
- `TrainFittedMedianImputer.fit(self, frame: pd.DataFrame, feature_columns: Sequence[str])`
- `TrainFittedMedianImputer.transform(self, frame: pd.DataFrame, feature_columns: Sequence[str])`
- `TrainFittedMedianImputer.to_dict(self)`
- `preprocess_dataset(frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None, passthrough_columns: Optional[Sequence[str]]=None, *, fill_missing: bool=True, quality_filter_column: str | None=None)` — Select numeric fields and optionally apply legacy all-data median filling.

## [datasets/plant_registry.py](../src/solar_forecast/datasets/plant_registry.py)

발전소 registry와 공식 ASOS 관측소·검토된 연결 근거 관리

- `_normalized_mapping_text(value: object)`
- `AdministrativeArea`
- `StationMatch`
- `ReviewedStationMapping`
- `ReviewedStationMapping.__post_init__(self)`
- `ReviewedStationMappingCatalog` — Version-controlled, auditable exceptions for plants without local ASOS.
- `ReviewedStationMappingCatalog.__init__(self, mappings: Iterable[ReviewedStationMapping]=())`
- `ReviewedStationMappingCatalog.from_json(cls, path: Path)`
- `ReviewedStationMappingCatalog.get(self, company: str, plant: str)`
- `parse_administrative_area(address: object)`
- `_haversine_km(lat1: float, lon1: float, lat2: float, lon2: float)`
- `KmaStationCatalog` — Resolve a plant to an ASOS station without hiding ambiguous guesses.
- `KmaStationCatalog.__init__(self, stations: pd.DataFrame)`
- `KmaStationCatalog.from_metadata(cls, path: Path)`
- `KmaStationCatalog.by_id(self, station_id: int, *, generation_start: object=None, generation_end: object=None)`
- `KmaStationCatalog.resolve(self, *, address: object, latitude: object, longitude: object, reviewed_station_id: int | None=None, generation_start: object=None, generation_end: object=None)`
- `KmaStationCatalog._covering_stations(cls, stations: pd.DataFrame, *, generation_start: object, generation_end: object)`
- `KmaStationCatalog._intervals_cover_range(valid_from: pd.Series, valid_to: pd.Series, earliest: pd.Timestamp, latest: pd.Timestamp)`
- `KmaStationCatalog._latest_station_rows(stations: pd.DataFrame)`
- `KmaStationCatalog._distance(latitude: object, longitude: object, station: pd.Series)`
- `NationwidePlantRegistryBuilder` — Create one auditable nationwide identity and station-mapping table.
- `NationwidePlantRegistryBuilder.__init__(self, metadata: PlantMetadataCatalog, stations: KmaStationCatalog, reviewed_mappings: ReviewedStationMappingCatalog | None=None)`
- `NationwidePlantRegistryBuilder.build(self, paths: Iterable[Path], destination: Path, *, legacy_mapping: pd.DataFrame | None=None)`
- `NationwidePlantRegistryBuilder._scan(self, paths: Iterable[Path])`
- `NationwidePlantRegistryBuilder._legacy_map(self, frame: pd.DataFrame | None)`

## [datasets/preparation_service.py](../src/solar_forecast/datasets/preparation_service.py)

원본 표준화·collector admission·registry·Gold·품질 감사 orchestration

- `DataPreparationResult`
- `DataPreparationService` — Application boundary for historical standardization and model-ready features.
- `DataPreparationService.__init__(self, input_root: Path, weather_root: Path, merged_source: Path, output_dir: Path, collected_generation_dir: Path | None=COLLECTOR_SILVER_ROOT)`
- `DataPreparationService.run(self)`

## [datasets/repository.py](../src/solar_forecast/datasets/repository.py)

CSV·압축 CSV·파티션 학습 자료 탐색과 메모리 제한 분할 로딩

- `_data_extension(path: Path)`
- `DatasetLoadPolicy`
- `DatasetLoadPolicy.__post_init__(self)`
- `DatasetLoadReport`
- `DatasetLoadReport.to_dict(self)`
- `discover_latest_file(input_dir: Path)` — Find the most recently modified supported dataset below ``input_dir``.
- `DatasetRepository` — Filesystem adapter for discovering and loading tabular model data.
- `DatasetRepository.__init__(self, input_dir: Path)`
- `DatasetRepository.load(self, data_path: Optional[Path]=None)`
- `DatasetRepository.load_training_frame(self, data_path: Path, *, columns: Sequence[str], numeric_columns: Sequence[str], equals_filters: dict[str, object] | None=None, truthy_filter: str | None=None, row_limit: int | None=None, policy: DatasetLoadPolicy | None=None)` — Read only model columns and push row filters into bounded CSV chunks.
- `DatasetRepository._training_files(source: Path)`
- `load_dataset(data_path: Optional[Path], input_dir: Path)` — Compatibility facade around DatasetRepository.

## [evaluation/__init__.py](../src/solar_forecast/evaluation/__init__.py)

evaluation 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [evaluation/experiment_config.py](../src/solar_forecast/evaluation/experiment_config.py)

실측 예측 실험 스키마 검증과 모델별 특징·입력 길이·탐색 예산 구성

- `load_experiment_config(path: Path, *, project_root: Path=PROJECT_ROOT)` — Validate the executable experiment, including explicitly bounded searches.
- `build_candidate_configs(values: dict, horizon: int, run_dir: Path, *, project_root: Path=PROJECT_ROOT)` — Allow independent features/lookbacks while freezing target and time splits.
- `experiment_plan(values: dict, *, project_root: Path=PROJECT_ROOT)` — Describe actual candidates without loading data or starting training.
- `_resolve_path(value: str, project_root: Path)`

## [evaluation/feature_ablation.py](../src/solar_forecast/evaluation/feature_ablation.py)

시간 순서를 지키는 rolling-origin 특징 조합 비교 실험

- `FeatureAblationResult`
- `RollingOriginFold`
- `RollingOriginFold.to_dict(self)`
- `FeatureAblationService` — Select features with purged rolling-origin folds before Calibration/Test.
- `FeatureAblationService.__init__(self, *, seed: int=42, n_estimators: int=300)`
- `FeatureAblationService.run(self, dataset_path: Path, output_dir: Path, *, n_splits: int=3, validation_window_hours: int=2160, calibration_fraction: float=0.1, test_fraction: float=0.15, gap_hours: int=168)`
- `FeatureAblationService._rolling_folds(timestamps: pd.Series, *, n_splits: int, validation_window_hours: int, calibration_fraction: float, test_fraction: float, gap_hours: int)`
- `FeatureAblationService._as_bool(series: pd.Series)`

## [evaluation/forecast_samples.py](../src/solar_forecast/evaluation/forecast_samples.py)

발전소별 관측 시작 시각과 예측 목표 정렬, 시간창 연속성 검증

- `validate_forecast_horizon(value: object)` — Reject ambiguous horizons instead of silently truncating fractional hours.
- `validate_observation_frame(frame: pd.DataFrame, *, entity_column: str, timestamp_column: str)` — Require a unique observed plant-hour before constructing forecast targets.
- `build_forecast_samples(frame: pd.DataFrame, feature_columns: Sequence[str], target_column: str, horizon_hours: int, *, entity_column: str='plant_id', timestamp_column: str='timestamp')` — Pair y(t) with features observed at exactly t-h, independently per plant.
- `forecast_window_positions(timestamps: Sequence[object], *, horizon_hours: int, sequence_length: int)` — Return target and origin positions whose input window is truly hourly.
- `forecast_evaluation_contract(prediction_task: str | None, horizon_hours: object, *, legacy_task: str)` — Describe implemented timing; legacy row models never claim a 24h horizon.

## [evaluation/model_parity.py](../src/solar_forecast/evaluation/model_parity.py)

동일 plant-hour 예측 CSV의 수치 동등성·지표 차이·입력 무결성 비교

- `_parse_hourly_timestamps(values: pd.Series, label: str)` — Parse explicit ISO hour keys without guessing a timezone for naive values.
- `_load_prediction_artifact(path: Path, label: str)` — Read the exact hashed bytes and reject ambiguous or unusable prediction rows.
- `_calculate_error_metrics(actual: np.ndarray, predicted: np.ndarray)` — Compute pooled row metrics, preserving undefined R2 as JSON null.
- `_compare_aligned_rows(frame: pd.DataFrame, *, atol: float, rtol: float)` — Measure candidate differences using baseline predictions as the reference.
- `compare_prediction_files(baseline: Path, candidate: Path, *, atol: float=1e-06, rtol: float=1e-06)` — Compare identical observed plant-hours; invalid input raises ValueError.

## [evaluation/model_selection.py](../src/solar_forecast/evaluation/model_selection.py)

별도 시간 구간에서 하이브리드 비중 학습·채택 결정·최종 Test 평가

- `_validate_frame(frame: pd.DataFrame, name: str, horizon: int)`
- `_metrics(actual: pd.Series, prediction: pd.Series)`
- `_model_metrics(frame: pd.DataFrame, column: str)`
- `_all_metrics(frame: pd.DataFrame)`
- `_period(frame: pd.DataFrame)`
- `_predict_gate(gate: ExplainableDynamicGate, frame: pd.DataFrame)`
- `_write_csv_atomic(frame: pd.DataFrame, path: Path, *, compression: str | None=None)`
- `BenchmarkModelSelector` — Fit the blend before selection, freeze the champion before reading Test errors.
- `BenchmarkModelSelector.__init__(self, output_dir: Path, minimum_relative_improvement: float=0.0, selection_gap_hours: int=0)`
- `BenchmarkModelSelector.run(self, calibration: pd.DataFrame, test: pd.DataFrame, *, evaluation_contract: dict[str, Any], provenance: dict[str, Any])` — Persist a selection decision and four-model errors on identical held-out rows.

## [evaluation/regression_metrics.py](../src/solar_forecast/evaluation/regression_metrics.py)

예측 정합성 검증과 발전소·지역·전국 회귀 오차 집계

- `validate_predictions(frame: pd.DataFrame)`
- `calculate_metrics(frame: pd.DataFrame)` — Calculate metrics from rows instead of averaging pre-calculated RMSE/R2.
- `_group_metrics(frame: pd.DataFrame, columns: Sequence[str])`
- `aggregate_metrics(frame: pd.DataFrame)` — Return exact plant, region and national metrics from aligned predictions.

## [evaluation/temporal_split.py](../src/solar_forecast/evaluation/temporal_split.py)

Train·Validation·Calibration·Test의 시간 경계와 purge gap 생성

- `TemporalSplitConfig` — Leakage-safe four-way split shared by every forecasting model.
- `TemporalSplitConfig.__post_init__(self)`
- `TemporalSplitConfig.train_fraction(self)`
- `TemporalBoundaries`
- `TemporalBoundaries.to_dict(self)`
- `TemporalFrameSplits`
- `TemporalSplitter` — Apply one set of global timestamp boundaries to all plants.
- `TemporalSplitter.__init__(self, config: TemporalSplitConfig | None=None)`
- `TemporalSplitter.boundaries(self, timestamps: pd.Series | pd.Index)`
- `TemporalSplitter.labels(self, timestamps: pd.Series | pd.Index, boundaries: TemporalBoundaries)`
- `TemporalSplitter.split_frame(self, frame: pd.DataFrame, timestamp_column: str='timestamp')`

## [features/__init__.py](../src/solar_forecast/features/__init__.py)

features 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [features/asos_features.py](../src/solar_forecast/features/asos_features.py)

ASOS 관측값 이름·타입 표준화와 관측소 좌표 결합

- `KmaAsosNormalizer` — Normalize selected ASOS fields and join stable station coordinates.
- `KmaAsosNormalizer.__init__(self, station_metadata_path: Path)`
- `KmaAsosNormalizer.read(self, paths: Iterable[Path], *, station_ids: Iterable[int] | None=None)`
- `KmaAsosNormalizer.transform(self, frame: pd.DataFrame)`

## [features/history_features.py](../src/solar_forecast/features/history_features.py)

미래 누수를 차단하는 발전량 lag·rolling·태양 위치 특징 생성

- `HistoryPolicy`
- `LeakageSafeFeatureEngineer` — Build time and history features available at a day-ahead issue time.
- `LeakageSafeFeatureEngineer.__init__(self, policy: HistoryPolicy | None=None)`
- `LeakageSafeFeatureEngineer.transform(self, frame: pd.DataFrame, *, timestamp_column: str='timestamp', entity_column: str='plant_id', target_column: str='generation_mwh')`
- `LeakageSafeFeatureEngineer._solar_geometry(frame: pd.DataFrame, timestamp_column: str)` — Approximate solar elevation and clear-sky horizontal potential for KST.

## [infrastructure/__init__.py](../src/solar_forecast/infrastructure/__init__.py)

infrastructure 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [infrastructure/artifact_store.py](../src/solar_forecast/infrastructure/artifact_store.py)

파일 hash·원자적 교체·JSON manifest·실행 폴더 생성

- `sha256_file(path: Path, chunk_bytes: int=1024 * 1024)` — Hash an artifact incrementally so lineage checks stay memory bounded.
- `replace_file_atomic(temporary: Path, destination: Path, attempts: int=8)` — Atomically replace a file, retrying short Windows scanner/indexer locks.
- `write_json_atomic(path: Path, payload: dict[str, Any])` — Replace a JSON artifact only after its complete temporary write succeeds.
- `write_manifest(path: Path, *, status: str, model: str, run_id: str, details: dict[str, Any])`
- `create_run_directory(base_dir: Path)`

## [infrastructure/csv_storage.py](../src/solar_forecast/infrastructure/csv_storage.py)

CSV 스키마·인코딩·행 수 감사와 표준 CSV 저장

- `CsvArtifactAudit`
- `CsvArtifactAudit.as_dict(self)`
- `inspect_csv_artifact(path: Path)` — Hash and identify a Korean CSV in one bounded-memory sequential scan.
- `write_standardized_csv(frame: pd.DataFrame, destination: Path)` — Atomically write a Silver CSV using Excel-compatible UTF-8 with BOM.

## [infrastructure/environment.py](../src/solar_forecast/infrastructure/environment.py)

Git에서 제외된 .env.local을 읽어 미설정 환경변수에 적용

- `load_local_env(path: Path | None=None)` — Load missing environment values from a git-ignored local file.

## [infrastructure/error_reporting.py](../src/solar_forecast/infrastructure/error_reporting.py)

단계·예외·실행 문맥을 오류 보고서에 기록

- `write_error_report(output_dir: Path, exc: BaseException, *, stage: str, context: Optional[Mapping[str, Any]]=None)` — Create an error file only after a failed run and return its path.

## [infrastructure/project_paths.py](../src/solar_forecast/infrastructure/project_paths.py)

원본·Silver·Gold·모델·평가 산출물의 기본 저장 위치 정의

공개 export 또는 설정 상수만 정의합니다.

## [jobs/__init__.py](../src/solar_forecast/jobs/__init__.py)

jobs 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [jobs/benchmark_job.py](../src/solar_forecast/jobs/benchmark_job.py)

모델별 Validation 탐색과 하이브리드 선택 및 예측 산출물 연결

- `align_prediction_frames(frames: dict[str, pd.DataFrame], *, minimum_coverage: float)` — Align an explicit cohort, failing on changed truth, IDs or low coverage.
- `BenchmarkService` — Execute independently tuned candidates and freeze selection before Test.
- `BenchmarkService.__init__(self, *, project_root: Path=PROJECT_ROOT, training_service=None)`
- `BenchmarkService.run(self, config_path: Path, *, smoke: bool=False)`
- `BenchmarkService._candidate_summary(entry: tuple, score: float, run_root: Path)`
- `BenchmarkService._matching_contract(selected: dict)`
- `BenchmarkService._merge_selected(selected: dict, split: str, values: dict)`
- `BenchmarkService._provenance(self, source: Path, values: dict)`

## [jobs/contracts.py](../src/solar_forecast/jobs/contracts.py)

독립 실행 job 목록과 교환 manifest의 버전·입출력 계약 정의

- `ArtifactContract` — A versioned artifact exchanged between independently runnable jobs.
- `ArtifactContract.as_dict(self)`
- `JobContract` — Public job boundary that can later become a worker/container entrypoint.
- `JobContract.as_dict(self)`
- `_manifest_schema(contract: str, *, required: tuple[str, ...], properties: dict[str, Any], contract_field: str='contract')`
- `_artifact(name: str, contract: str, path_pattern: str, schema: dict[str, Any], *, required_for_next_job: bool=True, notes: tuple[str, ...]=())`
- `list_job_contracts()`
- `get_job_contract(job_id: str)`
- `job_contract_catalog()`

## [jobs/notification_dispatch_job.py](../src/solar_forecast/jobs/notification_dispatch_job.py)

이벤트 무결성·수신 경로 확인 후 알림 outbox와 dispatcher 실행

- `NotificationDispatchJobConfig`
- `NotificationDispatchJobResult`
- `NotificationDispatchJob` — Validate operational anomaly events and dispatch them through an outbox.
- `NotificationDispatchJob.__init__(self, *, project_root: Path=PROJECT_ROOT, env_loader: Callable[[], None]=load_local_env)`
- `NotificationDispatchJob.run(self, config: NotificationDispatchJobConfig)`
- `NotificationDispatchJob.run.validated_alerts()`
- `NotificationDispatchJob._project_path(self, value: str | Path)`

## [jobs/training_job.py](../src/solar_forecast/jobs/training_job.py)

모델별 trainer 선택·단일 학습 잠금·실행 manifest 기록

- `TrainingService` — Runs one model strategy behind a process-wide lock and manifest boundary.
- `TrainingService.__init__(self, trainers: dict[str, tuple[str, str]] | None=None)`
- `TrainingService.run(self, config: ModelJobConfig, *, smoke: bool=False)`
- `run_training(config: ModelJobConfig, *, smoke: bool=False)` — Compatibility facade around TrainingService.

## [jobs/training_lock.py](../src/solar_forecast/jobs/training_lock.py)

동시 학습 실행을 차단하는 프로세스 잠금

- `TrainingAlreadyRunning`
- `exclusive_training_lock(lock_path: Path, model: str)` — Prevent XGBoost and CNN training from consuming resources concurrently.

## [jobs/verification_job.py](../src/solar_forecast/jobs/verification_job.py)

수집·데이터 준비·학습·대시보드 연결의 단계별 실행 결과 검증

- `VerificationConfig` — Runtime contract for a bounded end-to-end pipeline verification.
- `PipelineVerificationResult`
- `PipelineVerificationService` — Run the public data-to-model wiring in a reproducible, auditable way.
- `PipelineVerificationService.__init__(self, config: VerificationConfig)`
- `PipelineVerificationService.run(self)`
- `PipelineVerificationService._run_step(self, steps: list[dict[str, Any]], name: str, action: Callable[[], dict[str, Any]])`
- `PipelineVerificationService._inspect_collection_manifest(self)`
- `PipelineVerificationService._run_collection(self)`
- `PipelineVerificationService._run_prepare_data(self)`
- `PipelineVerificationService._run_train(self, model: str)`
- `PipelineVerificationService._run_dashboard(self)`
- `PipelineVerificationService._model_config(self, model: str)`
- `PipelineVerificationService._latest_collection_manifest(self)`
- `PipelineVerificationService._resolve(self, path: str | Path | None)`
- `PipelineVerificationService._read_json(path: Path)`
- `PipelineVerificationService._warning_count(steps: list[dict[str, Any]])`
- `PipelineVerificationService._prediction_artifacts(details: dict[str, Any])`
- `build_collection_config(*, start_date: str, end_date: str | None, sources: str, output_dir: str | Path, standardized_output_dir: str | Path, overwrite: bool, download_date: str | None, komipo_station_codes: tuple[str, ...], api_max_calls: int, station_ids: tuple[str, ...]=(), kma_mode: str='auto', weather_root: str | Path='file/KMA_data_file')` — Create a collection config for the verification CLI boundary.

## [models/__init__.py](../src/solar_forecast/models/__init__.py)

models 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [models/cnn_bilstm/__init__.py](../src/solar_forecast/models/cnn_bilstm/__init__.py)

models/cnn_bilstm 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [models/cnn_bilstm/adaptive_training.py](../src/solar_forecast/models/cnn_bilstm/adaptive_training.py)

선택적 bandit 기반 적응 학습과 중단 상태 복구

- `BanditConfig` — Configuration for the epsilon-greedy multi-armed bandit.
- `BanditState`
- `BanditState.initialize(self, actions: Sequence[float])`
- `BanditState.select(self, actions: Sequence[float], epsilon: float)`
- `BanditState.update(self, action: float, reward: float, alpha: float)`
- `ReinforcementResult`
- `run_adaptive_training(model_cfg: CnnBiLstmNetworkConfig, train_loader, val_loader, epochs: int=30, bandit_cfg: BanditConfig | None=None, device: torch.device | None=None, checkpoint_store: TrainingCheckpointStore | None=None, checkpoint_stage: str='adaptive_fit', checkpoint_signature: str='adaptive_fit_v1', initial_model: CNNBiLSTM | None=None)` — Train the model while a bandit adapts the learning rate.

## [models/cnn_bilstm/evaluation.py](../src/solar_forecast/models/cnn_bilstm/evaluation.py)

저장된 CNN 모델 평가·체크포인트 비교·고정 보정 임계값 이상치 분석

- `load_checkpoint(path: str, device: Optional[torch.device]=None)`
- `evaluate_model(model: torch.nn.Module, data_loader, device: Optional[torch.device]=None)`
- `compare_checkpoints(checkpoint_dir: str, frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None, sequence_config: Optional[SequenceConfig]=None)` — Load all checkpoints in a directory and compare their metrics.
- `detect_outliers_from_predictions(y_true: np.ndarray, y_pred: np.ndarray, contamination: float=0.05, *, calibration_residuals: np.ndarray | None=None)` — Apply a frozen calibration threshold; never rank the evaluated set itself.
- `evaluate_and_analyze(checkpoint_path: str, frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None, sequence_config: Optional[SequenceConfig]=None, contamination: float=0.05, output_dir: Optional[str]=None, entity_column: Optional[str]=None, timestamp_column: Optional[str]=None)` — Load a checkpoint, compute metrics, and perform outlier analysis.
- `save_dataframe(df: pd.DataFrame, path: str)`

## [models/cnn_bilstm/network.py](../src/solar_forecast/models/cnn_bilstm/network.py)

CNN-BiLSTM 신경망 구조·층 구성·forward와 네트워크 설정

- `CnnBiLstmNetworkConfig` — Configuration for the CNN-BiLSTM architecture.
- `CNNBiLSTM` — 1D CNN followed by a bidirectional LSTM for sequence regression.
- `CNNBiLSTM.__init__(self, config: CnnBiLstmNetworkConfig)`
- `CNNBiLSTM.forward(self, x: torch.Tensor)`
- `build_cnn_bilstm_network(config: CnnBiLstmNetworkConfig, device: Optional[torch.device]=None)` — Helper to build and place the model on a device.

## [models/cnn_bilstm/optimization.py](../src/solar_forecast/models/cnn_bilstm/optimization.py)

CNN epoch 학습·검증 루프, Optuna 탐색과 최적 설정 재학습

- `train_cnn_bilstm_epoch(model: CNNBiLSTM, loader: DataLoader, criterion: nn.Module, optimizer: torch.optim.Optimizer, device: torch.device)`
- `evaluate_cnn_bilstm_loader(model: CNNBiLSTM, loader: DataLoader, criterion: nn.Module, device: torch.device)`
- `suggest_cnn_bilstm_config(trial: Trial, n_features: int, search_space: dict[str, object] | None=None)`
- `limit_sequence_loader(loader: DataLoader, maximum_sequences: int | None, *, shuffle: bool)`
- `optimize_cnn_bilstm(frame, target_column: str, feature_columns=None, sequence_config: Optional[SequenceConfig]=None, n_trials: int=20, device: Optional[torch.device]=None, timeout: Optional[int]=None, entity_column: Optional[str]=None, timestamp_column: Optional[str]=None, trial_epochs: int=20, early_stopping_patience: int=5, maximum_train_sequences: int | None=None, maximum_validation_sequences: int | None=None, settings: OptimizationSettings | None=None, artifact_dir: Path | None=None, checkpoint_store: TrainingCheckpointStore | None=None, optimizer_parameter_space: dict[str, object] | None=None)` — Select architecture and optimizer values from Validation only.
- `optimize_cnn_bilstm.objective(trial: Trial)`
- `optimize_cnn_bilstm.cleanup_checkpoint(_study: optuna.Study, frozen_trial)`
- `train_with_best_trial(frame, target_column: str, feature_columns=None, sequence_config: Optional[SequenceConfig]=None, n_trials: int=20, device: Optional[torch.device]=None, entity_column: Optional[str]=None, timestamp_column: Optional[str]=None, epochs: int=50, trial_epochs: int=20, early_stopping_patience: int=5, maximum_train_sequences: int | None=None, maximum_validation_sequences: int | None=None, settings: OptimizationSettings | None=None, artifact_dir: Path | None=None, timeout: Optional[int]=None, checkpoint_store: TrainingCheckpointStore | None=None, optimizer_parameter_space: dict[str, object] | None=None)` — Run Optuna then train/evaluate the best model; returns artifacts.
- `save_study_results(study: optuna.Study, path: str)` — Persist study results to disk.

## [models/cnn_bilstm/sequence_config.py](../src/solar_forecast/models/cnn_bilstm/sequence_config.py)

CNN 입력 길이·시간 분할 비율·배치 로더 설정

- `SequenceConfig` — Sequence construction and leakage-safe four-way temporal split settings.
- `SequenceConfig.__post_init__(self)`

## [models/cnn_bilstm/sequence_data.py](../src/solar_forecast/models/cnn_bilstm/sequence_data.py)

발전소별 lazy 시계열 window·학습 로더·Train 결측 대체 준비

- `SequenceDataset` — Compatibility dataset for already-materialized small arrays.
- `SequenceDataset.__init__(self, X: np.ndarray, y: np.ndarray)`
- `SequenceDataset.__len__(self)`
- `SequenceDataset.__getitem__(self, idx: int)`
- `_EntitySeries`
- `LazyWindowSequenceDataset` — Slice windows on demand instead of materializing repeated 3-D arrays.
- `LazyWindowSequenceDataset.__init__(self, series: Sequence[_EntitySeries], split: str, sequence_length: int)`
- `LazyWindowSequenceDataset.__len__(self)`
- `LazyWindowSequenceDataset.__getitem__(self, idx: int)`
- `LazyWindowSequenceDataset.context_frame(self, start: int, stop: int)` — Materialize metadata only for one sequential prediction batch.
- `_build_sequences(values: np.ndarray, targets: np.ndarray, sequence_length: int)` — Compatibility helper for small callers; the main path uses lazy windows.
- `SequenceLoaders`
- `_position_labels(n_targets: int, cfg: SequenceConfig)`
- `_fit_and_transform_training_medians(series: Sequence[_EntitySeries], feature_columns: Sequence[str], *, append_missing_indicators: bool)`
- `prepare_dataset_splits(frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None, config: Optional[SequenceConfig]=None, entity_column: Optional[str]=None, timestamp_column: Optional[str]=None)` — Build one global four-way time split with lazy per-entity windows.
- `prepare_datasets(frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None, config: Optional[SequenceConfig]=None, entity_column: Optional[str]=None, timestamp_column: Optional[str]=None)` — Compatibility view returning Train/Validation/Test from four-way splits.
- `sequence_from_csv(path: str, target_column: str, feature_columns: Optional[Sequence[str]]=None, config: Optional[SequenceConfig]=None)`
- `reconstruct_targets_from_loader(loader: DataLoader)`

## [models/cnn_bilstm/trainer.py](../src/solar_forecast/models/cnn_bilstm/trainer.py)

CNN 모델 JSON 설정을 데이터 로딩과 학습 workflow에 연결

- `CnnBiLstmTrainer` — Concrete adapter from job configuration to the CNN-BiLSTM workflow.
- `CnnBiLstmTrainer.train(self, config: ModelJobConfig, run_dir: Path, smoke: bool=False)`
- `train(config: ModelJobConfig, *, run_dir: Path, smoke: bool=False)`

## [models/cnn_bilstm/training_workflow.py](../src/solar_forecast/models/cnn_bilstm/training_workflow.py)

CNN 최종 학습·체크포인트·Validation/Calibration/Test 예측 파일 기록

- `_write_prediction_artifact(model: torch.nn.Module, loader, path: Path, *, split: str, device: torch.device)` — Stream row-aligned CNN predictions without materializing a full table.
- `train_cnn_bilstm(frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None, sequence_config: Optional[SequenceConfig]=None, n_trials: int=10, output_dir: str='artifacts/models/cnn_bilstm', use_optuna: bool=True, use_reinforcement: bool=False, epochs: int=50, entity_column: Optional[str]=None, timestamp_column: Optional[str]=None, optimizer_settings: OptimizationSettings | None=None, optimizer_trial_epochs: int=20, early_stopping_patience: int=5, optimizer_max_train_sequences: int | None=None, optimizer_max_validation_sequences: int | None=None, optimizer_timeout_seconds: int | None=None, checkpoint_store: TrainingCheckpointStore | None=None, checkpoint_root: str | Path | None=None, optimizer_storage_path: str | Path | None=None, optimizer_parameter_space: Mapping[str, object] | None=None, seed: int=42)` — Train the model with Optuna and/or reinforcement learning then persist artifacts.

## [models/hybrid/__init__.py](../src/solar_forecast/models/hybrid/__init__.py)

models/hybrid 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [models/hybrid/dynamic_gate.py](../src/solar_forecast/models/hybrid/dynamic_gate.py)

Validation 근거로 CNN/XGBoost 결합 비중 학습과 계층적 fallback

- `DynamicGateConfig`
- `DynamicGateConfig.__post_init__(self)`
- `ExplainableDynamicGate` — Validation-fitted, context-aware ensemble with auditable decisions.
- `ExplainableDynamicGate.__init__(self, config: DynamicGateConfig | None=None)`
- `ExplainableDynamicGate.fit(self, validation: pd.DataFrame)`
- `ExplainableDynamicGate.predict(self, test: pd.DataFrame)`
- `ExplainableDynamicGate.fit_predict(self, validation: pd.DataFrame, test: pd.DataFrame)`
- `normalize_prediction_columns(frame: pd.DataFrame)` — Map known source labels to the hybrid contract and derive hour.
- `validate_aligned_predictions(frame: pd.DataFrame)`
- `_optimal_convex_weight(y_true: np.ndarray, xgb_pred: np.ndarray, cnn_pred: np.ndarray)`
- `fit_region_blend(validation: pd.DataFrame)` — Compatibility baseline: learn one convex weight per region.
- `predict_region_blend(test: pd.DataFrame, weights: pd.DataFrame, fallback_weight: float=0.5)` — Compatibility baseline: apply frozen regional weights.
- `_build_profiles(validation: pd.DataFrame, min_group_samples: int)` — Learn a transparent hierarchy of validation-only routing contexts.
- `_build_profiles.profile(scope: str, group: pd.DataFrame, **context: object)`
- `_apply_profiles(test: pd.DataFrame, gate: pd.DataFrame, config: DynamicGateConfig)` — Route each row by plant/time/disagreement context and retain its rationale.
- `_time_regime(hour: int)`
- `_disagreement_band(values: pd.Series, low: float, high: float)`
- `fit_dynamic_gate(validation: pd.DataFrame, min_group_samples: int=8)` — Functional compatibility wrapper around ExplainableDynamicGate.fit.
- `predict_dynamic_hybrid(test: pd.DataFrame, gate: pd.DataFrame)` — Functional compatibility wrapper for persisted gate profiles.

## [models/hybrid/experiment.py](../src/solar_forecast/models/hybrid/experiment.py)

하이브리드 예측 입력·결합·평가·실험 산출물 기록

- `HybridExperiment` — Coordinates prediction I/O, dynamic gating, metrics, and artifacts.
- `HybridExperiment.__init__(self, output_dir: Path, artifact_level: str='minimal')`
- `HybridExperiment.run(self, validation_path: Path, test_path: Path)`
- `HybridExperiment._write_artifacts(self, profiles: pd.DataFrame, predictions: pd.DataFrame, metrics: dict[str, pd.DataFrame])`
- `build_hybrid_experiment(validation_path: Path, test_path: Path, output_dir: Path, artifact_level: str='minimal')` — Compatibility facade for callers that have not migrated to HybridExperiment.

## [models/shared/__init__.py](../src/solar_forecast/models/shared/__init__.py)

models/shared 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [models/shared/checkpoint_store.py](../src/solar_forecast/models/shared/checkpoint_store.py)

공통 dataset/config fingerprint와 프레임워크별 학습 상태 저장·복구

- `stable_signature(payload: Mapping[str, Any] | list[Any] | tuple[Any, ...])`
- `dataframe_signature(frame: pd.DataFrame, columns: list[str])` — Hash selected frame values without serializing a second tabular copy.
- `dataset_signature(source: Path)` — Fingerprint one file or partitioned dataset independently of the model.
- `training_fingerprint(config: ModelJobConfig)`
- `capture_rng_state()`
- `restore_rng_state(state: Mapping[str, Any] | None)`
- `CheckpointSettings`
- `CheckpointSettings.from_config(cls, config: ModelJobConfig)`
- `TrainingCheckpointStore` — Atomic, fingerprint-scoped state for interrupted model training.
- `TrainingCheckpointStore.__init__(self, root: Path, *, model: str, fingerprint: str, enabled: bool=True, resume: bool=True, cnn_every_epochs: int=1, xgboost_every_rounds: int=50)`
- `TrainingCheckpointStore.from_config(cls, config: ModelJobConfig, *, project_root: Path=PROJECT_ROOT)`
- `TrainingCheckpointStore.directory(self)`
- `TrainingCheckpointStore.torch_path(self, stage: str)`
- `TrainingCheckpointStore.xgboost_path(self, stage: str)`
- `TrainingCheckpointStore.save_torch(self, stage: str, payload: Mapping[str, Any], *, signature: str, progress: Mapping[str, Any], completed: bool=False)`
- `TrainingCheckpointStore.load_torch(self, stage: str, *, signature: str, map_location: torch.device | str='cpu')`
- `TrainingCheckpointStore.save_xgboost(self, booster: Any, stage: str, *, signature: str, completed_rounds: int, completed: bool=False)`
- `TrainingCheckpointStore.load_xgboost(self, stage: str, *, signature: str)`
- `TrainingCheckpointStore.remove(self, stage: str, *, kind: str)`
- `TrainingCheckpointStore.describe(self)`
- `TrainingCheckpointStore._write_metadata(self, path: Path, stage: str, signature: str, progress: Mapping[str, Any], completed: bool)`
- `TrainingCheckpointStore._validate_state(self, state: Mapping[str, Any], stage: str, signature: str)`
- `TrainingCheckpointStore._metadata_path(path: Path)`
- `TrainingCheckpointStore._stage_name(stage: str)`

## [models/shared/optuna_study.py](../src/solar_forecast/models/shared/optuna_study.py)

공통 Optuna 저장소·탐색 공간·실험 예산·study 재개 관리

- `optimizer_search_space(values: Mapping[str, object])` — Return an optional Optuna search-space section from a model config.
- `suggest_parameter(trial: Trial, name: str, search_space: Mapping[str, object], default: Mapping[str, object])` — Suggest one parameter from config, falling back to a model-owned default.
- `OptimizationSettings` — Shared, bounded Optuna contract for independently trained models.
- `OptimizationSettings.scoped(self, training_fingerprint: str)` — Bind the persistent study to one data and training contract.
- `OptimizationSettings.from_values(cls, values: Mapping[str, object], *, model: str)`
- `OptimizationSettings.validate(self)`
- `OptimizationRun`
- `OptimizationRun.best_params(self)`
- `OptunaStudyService` — Run or resume one versioned study and persist auditable artifacts.
- `OptunaStudyService.__init__(self, settings: OptimizationSettings, *, project_root: Path=PROJECT_ROOT)`
- `OptunaStudyService.run(self, objective: Callable[[Trial], float], artifact_dir: Path, *, callbacks: Sequence[Callable[[Study, FrozenTrial], None]]=())`
- `OptunaStudyService._finished_trials(trials: list[FrozenTrial])`
- `OptunaStudyService._enqueue_failed_checkpoint_retries(self, study: Study)`

## [models/xgboost/__init__.py](../src/solar_forecast/models/xgboost/__init__.py)

models/xgboost 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [models/xgboost/checkpoint.py](../src/solar_forecast/models/xgboost/checkpoint.py)

XGBoost boosting round 중간 상태·조기 종료·재개 callback

- `ResumableEarlyStoppingCallback` — Early stopping whose best score and wait counter live in the Booster.
- `ResumableEarlyStoppingCallback.__init__(self, rounds: int, metric_name: str)`
- `ResumableEarlyStoppingCallback.before_training(self, model: Any)`
- `ResumableEarlyStoppingCallback.after_iteration(self, model: Any, epoch: int, evals_log: dict)`
- `XGBoostCheckpointCallback` — Persist a Booster atomically every configured number of rounds.
- `XGBoostCheckpointCallback.__init__(self, store: TrainingCheckpointStore, *, stage: str, signature: str, initial_rounds: int=0)`
- `XGBoostCheckpointCallback.after_iteration(self, model: Any, epoch: int, evals_log: dict)`
- `ResumableXGBoostFit`
- `fit_xgboost_resumable(params: dict[str, Any], train_x: Any, train_y: Any, validation_x: Any, validation_y: Any, *, store: TrainingCheckpointStore | None, stage: str, signature: str, verbose: bool=False, callbacks: Sequence[TrainingCallback]=())` — Continue boosting from a compatible periodic checkpoint when present.

## [models/xgboost/optimization.py](../src/solar_forecast/models/xgboost/optimization.py)

XGBoost 전용 Optuna 탐색·pruning·튜닝 데이터 제한

- `OptunaPruningCallback` — Report XGBoost validation MAE each boosting round to Optuna.
- `OptunaPruningCallback.__init__(self, trial: optuna.Trial)`
- `OptunaPruningCallback.after_iteration(self, model, epoch: int, evals_log: dict)`
- `XGBoostOptimizationResult`
- `XGBoostOptimizationResult.to_dict(self)`
- `XGBoostHyperparameterOptimizer` — Bounded multi-fidelity search followed by full-data model fitting.
- `XGBoostHyperparameterOptimizer.__init__(self, settings: OptimizationSettings, values: Mapping[str, object], checkpoint_store: TrainingCheckpointStore | None=None)`
- `XGBoostHyperparameterOptimizer.optimize(self, train_frame: pd.DataFrame, validation_frame: pd.DataFrame, *, feature_columns: Sequence[str], target_column: str, artifact_dir: Path)`
- `XGBoostHyperparameterOptimizer.optimize.objective(trial: optuna.Trial)`
- `XGBoostHyperparameterOptimizer.optimize.cleanup_checkpoint(_study: optuna.Study, frozen_trial)`
- `XGBoostHyperparameterOptimizer._bounded_rows(frame: pd.DataFrame, maximum: int)`

## [models/xgboost/trainer.py](../src/solar_forecast/models/xgboost/trainer.py)

XGBoost 데이터 준비·학습·예측·모델 산출물 계약 구현

- `XGBoostTrainer` — Concrete training strategy that owns XGBoost-specific persistence.
- `XGBoostTrainer.train(self, config: ModelJobConfig, run_dir: Path, smoke: bool=False)`
- `XGBoostTrainer._chronological_split(frame: pd.DataFrame, *, validation_fraction: float, calibration_fraction: float, test_fraction: float, purge_gap_hours: int, calendar_timestamps: pd.Series | None=None, prediction_task: str='observed_conditions_estimation')`
- `XGBoostTrainer._load(config: ModelJobConfig, *, columns: list[str], numeric_columns: list[str], energy_source: str | None, smoke: bool)`
- `XGBoostTrainer._prediction_frame(context: pd.DataFrame, actual: np.ndarray, predicted: np.ndarray, *, split: str)`
- `train(config: ModelJobConfig, *, run_dir: Path, smoke: bool=False)`

## [notifications/__init__.py](../src/solar_forecast/notifications/__init__.py)

notifications 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [notifications/contracts.py](../src/solar_forecast/notifications/contracts.py)

알림 이벤트·채널·상태·수신자 데이터 구조와 식별자 계약

- `NotificationChannel`
- `NotificationStatus`
- `Severity`
- `NotificationAudience`
- `utc_now()`
- `ensure_utc(value: datetime)`
- `stable_event_id(*, model: str, plant_id: str, observed_at: datetime, detector_version: str, influence_factor: str, signal_type: str, detector_role: str)` — Build a deterministic detector event ID suitable for idempotent reruns.
- `AnomalyAlert`
- `AnomalyAlert.__post_init__(self)`
- `AnomalyAlert.from_operational_event(cls, event: Mapping[str, Any], *, detector_version: str | None=None, severity: Severity | str | None=None)` — Map a full operational anomaly event into the alert contract.
- `RecipientRoute` — A manager-directory reference and its ordered delivery fallback chain.
- `RecipientRoute.__post_init__(self)`
- `NotificationEnvelope`
- `NotificationEnvelope.for_anomaly(cls, alert: AnomalyAlert, route: RecipientRoute, *, template_id: str)`
- `DeliveryResult`
- `DeliveryResult.__post_init__(self)`
- `normalize_channels(values: Sequence[str | NotificationChannel])`
- `redact_sensitive(value: str)` — Mask common Korean phone-number and bearer-token forms before audit writes.
- `_required_text(name: str, value: Any)`

## [notifications/delivery_provider.py](../src/solar_forecast/notifications/delivery_provider.py)

알림 공급자 인터페이스·SOLAPI 전송·전화번호 검증

- `TransientNotificationError` — A temporary provider failure that is safe to retry.
- `PermanentNotificationError` — A request failure that should fall back or enter the dead-letter queue.
- `AmbiguousNotificationError` — The provider may have accepted the request, so automatic retry is unsafe.
- `ProviderMessage`
- `NotificationProvider`
- `NotificationProvider.send(self, message: ProviderMessage)`
- `ContactDirectory`
- `ContactDirectory.resolve(self, recipient_id: str, channel: NotificationChannel)` — Resolve a manager reference at dispatch time without persisting phone numbers.
- `MappingContactDirectory` — In-memory adapter for a runtime-injected secret-manager/directory snapshot.
- `MappingContactDirectory.__init__(self, contacts: Mapping[tuple[str, NotificationChannel | str], str])`
- `MappingContactDirectory.resolve(self, recipient_id: str, channel: NotificationChannel)`
- `SolapiProvider` — SOLAPI v4 adapter for approved Kakao AlimTalk templates and SMS/LMS.
- `SolapiProvider.__init__(self, channel: NotificationChannel, settings: SolapiSettings, *, session: Any | None=None)`
- `SolapiProvider.send(self, message: ProviderMessage)`
- `SolapiProvider._template_variables(message: ProviderMessage)`
- `_recipient_phone(value: str)`
- `_sender_phone(value: str)`

## [notifications/notification_config.py](../src/solar_forecast/notifications/notification_config.py)

알림 환경변수·재시도·dry-run·실발송 설정 검증

- `_parse_bool(name: str, value: str)`
- `_positive_int(name: str, value: str | int)`
- `NotificationSettings` — Operational controls; external delivery is fail-safe off by default.
- `NotificationSettings.__post_init__(self)`
- `NotificationSettings.external_delivery_allowed(self)`
- `NotificationSettings.confirm_live(self, *, confirmed: bool)` — Apply the separate CLI/API live gate; environment variables cannot set it.
- `NotificationSettings.from_env(cls, environ: Mapping[str, str] | None=None, *, prefix: str='SOLAR_NOTIFY_')`
- `NotificationSettings.from_env.item(name: str, default: str)`
- `NotificationSettings.describe(self)`
- `SolapiSettings` — SOLAPI credentials are runtime-only and hidden from repr/output.
- `SolapiSettings.__post_init__(self)`
- `SolapiSettings.from_env(cls, environ: Mapping[str, str] | None=None, *, prefix: str='SOLAR_NOTIFY_')`
- `SolapiSettings.from_env.required(name: str)`
- `SolapiSettings.describe(self)`
- `SolapiSettings.authorization_header(self, *, now: datetime | None=None, salt: str | None=None)`

## [notifications/notification_service.py](../src/solar_forecast/notifications/notification_service.py)

이벤트 enqueue와 재시도·대체 채널·전송 결과 전이 처리

- `DispatchSummary`
- `NotificationService` — Convert validated operational anomalies into deduplicated outbox work.
- `NotificationService.__init__(self, outbox: NotificationOutbox, *, template_id: str)`
- `NotificationService.enqueue_anomaly(self, alert: AnomalyAlert, routes: Sequence[RecipientRoute], *, now: datetime | None=None)`
- `NotificationDispatcher` — Claim outbox work and dispatch it through runtime-only contacts/providers.
- `NotificationDispatcher.__init__(self, settings: NotificationSettings, outbox: NotificationOutbox, contact_directory: ContactDirectory, providers: Mapping[NotificationChannel | str, NotificationProvider])`
- `NotificationDispatcher.run_once(self, *, now: datetime | None=None)`
- `build_solapi_dispatcher(settings: NotificationSettings, outbox: NotificationOutbox, contact_directory: ContactDirectory, solapi: SolapiSettings, *, session: object | None=None)` — Construct the two-channel SOLAPI adapter after the caller passes --live.

## [notifications/outbox_repository.py](../src/solar_forecast/notifications/outbox_repository.py)

SQLite outbox의 멱등성·lease·재시도·완료 상태 영속화

- `EnqueueResult`
- `OutboxRecord`
- `OutboxRecord.channel(self)`
- `FailureTransition`
- `NotificationOutbox` — Transactional SQLite outbox with leases, fallback and append-only audit.
- `NotificationOutbox.__init__(self, path: Path)`
- `NotificationOutbox.enqueue(self, envelope: NotificationEnvelope, *, now: datetime | None=None)`
- `NotificationOutbox.claim_due(self, *, limit: int, lease_seconds: int, now: datetime | None=None)`
- `NotificationOutbox.mark_accepted(self, notification_id: int, *, expected_attempt: int, provider_message_id: str, now: datetime | None=None)`
- `NotificationOutbox.mark_dry_run(self, notification_id: int, *, expected_attempt: int, now: datetime | None=None)`
- `NotificationOutbox.mark_in_doubt(self, notification_id: int, *, expected_attempt: int, reason: str, now: datetime | None=None)` — Quarantine an ambiguous provider outcome without automatic resend.
- `NotificationOutbox.mark_failed(self, notification_id: int, *, expected_attempt: int, error: str, transient: bool, max_attempts_per_channel: int, retry_base_seconds: int, retry_max_seconds: int, now: datetime | None=None)`
- `NotificationOutbox.resolve_in_doubt(self, notification_id: int, *, resolution_note: str, provider_message_id: str | None=None, retry: bool=False, now: datetime | None=None)`
- `NotificationOutbox.get(self, notification_id: int)`
- `NotificationOutbox.audit_events(self, notification_id: int)`
- `NotificationOutbox.status_counts(self)`
- `NotificationOutbox._finish(self, notification_id: int, *, expected_attempt: int, target: NotificationStatus, action: str, provider_message_id: str | None, now: datetime | None)`
- `NotificationOutbox._locked_in_progress(connection: sqlite3.Connection, notification_id: int, expected_attempt: int)`
- `NotificationOutbox._initialize(self)`
- `NotificationOutbox._connection(self)`
- `NotificationOutbox._record(row: sqlite3.Row)`
- `NotificationOutbox._audit(connection: sqlite3.Connection, notification_id: int, *, occurred_at: str, action: str, from_status: NotificationStatus | None, to_status: NotificationStatus, attempt: int, detail: Mapping[str, Any])`
- `NotificationOutbox._routing_key(event_id: str, recipient_id: str)`
- `NotificationOutbox._delivery_key(event_id: str, recipient_id: str, channel: NotificationChannel)`
- `_format_utc_timestamp(value: datetime)`
- `_parse_utc_timestamp(value: str)`
- `_serialize_outbox_payload(value: Mapping[str, Any] | Sequence[Any])`

## [notifications/recipient_routes.py](../src/solar_forecast/notifications/recipient_routes.py)

로컬 수신 경로에서 연락처 환경변수 이름과 대상 매핑 해석

- `RuntimeRouteDirectory` — Routes public event keys to runtime-only manager phone environment values.
- `RuntimeRouteDirectory.from_file(cls, path: Path, *, environ: Mapping[str, str] | None=None, require_contacts: bool)`
- `RuntimeRouteDirectory.routes_for(self, route_key: str, audience: NotificationAudience, plant_id: str)`
- `_required_text(name: str, value: object)`

## [pipeline/__init__.py](../src/solar_forecast/pipeline/__init__.py)

pipeline 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)`

## [pipeline/cnn_training_adapter.py](../src/solar_forecast/pipeline/cnn_training_adapter.py)

전체 파이프라인에 CNN 학습·평가 결과를 전달하는 adapter

- `CnnTrainingAdapter` — Adapter exposing the CNN workflow to the application pipeline.
- `CnnTrainingAdapter.execute(self, frame: pd.DataFrame, features: list[str], config: PipelineConfig, run_dir: Path)`
- `train_and_evaluate(frame: pd.DataFrame, features: list[str], config: PipelineConfig, run_dir: Path)`

## [pipeline/contracts.py](../src/solar_forecast/pipeline/contracts.py)

데이터 로딩·전처리·학습·리포트 adapter 인터페이스

- `DatasetPort`
- `DatasetPort.load(self, data_path: Optional[Path]=None)`
- `PreprocessorPort`
- `PreprocessorPort.transform(self, frame: pd.DataFrame, target_column: str, feature_columns: Optional[Sequence[str]]=None)`
- `TrainingPort`
- `TrainingPort.execute(self, frame: pd.DataFrame, features: list[str], config: PipelineConfig, run_dir: Path)`
- `ReportPort`
- `ReportPort.write(self, metrics: dict[str, Any], anomalies: pd.DataFrame, source: Path, output_path: Path)`

## [pipeline/forecast_pipeline.py](../src/solar_forecast/pipeline/forecast_pipeline.py)

데이터→전처리→학습→분석→HTML 보고서 실행 조율

- `PipelineResult`
- `ForecastPipeline` — Application service composed from replaceable infrastructure adapters.
- `ForecastPipeline.__init__(self, config: PipelineConfig, repository: DatasetPort | None=None, preprocessor: PreprocessorPort | None=None, trainer: TrainingPort | None=None, reporter: ReportPort | None=None)`
- `ForecastPipeline.run(self)`
- `run_pipeline(config: PipelineConfig)` — Compatibility facade around ForecastPipeline.

## [pipeline/html_report_adapter.py](../src/solar_forecast/pipeline/html_report_adapter.py)

파이프라인 실행 결과를 Plotly HTML 보고서로 변환

- `HtmlReportWriter`
- `HtmlReportWriter.write(self, metrics: dict[str, Any], anomalies: pd.DataFrame, source: Path, output_path: Path)`
- `build_html_report(metrics: dict[str, Any], anomalies: pd.DataFrame, source: Path, output_path: Path)`

## [pipeline/pipeline_config.py](../src/solar_forecast/pipeline/pipeline_config.py)

전체 예측 파이프라인의 입력·학습·보고서 실행 설정

- `PipelineConfig` — Configuration shared by every pipeline stage.
- `PipelineConfig.__post_init__(self)`

## [quality/__init__.py](../src/solar_forecast/quality/__init__.py)

quality 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

공개 export 또는 설정 상수만 정의합니다.

## [quality/generation_quality.py](../src/solar_forecast/quality/generation_quality.py)

물리적 발전량 품질 판정·발전소별 통계·학습 적격 여부 감사

- `PhysicalQualityConfig` — Conservative rules that separate impossible values from suspicious context.
- `PhysicalQualityConfig.__post_init__(self)`
- `GenerationQualityPolicy` — Add auditable physical/contextual flags without fabricating target values.
- `GenerationQualityPolicy.__init__(self, config: PhysicalQualityConfig | None=None)`
- `GenerationQualityPolicy.apply(self, frame: pd.DataFrame)`
- `GenerationQualityPolicy._flatline_flags(self, frame: pd.DataFrame, daylight: pd.Series)`
- `GenerationQualityPolicy._daily_aggregate_profile_flags(self, frame: pd.DataFrame)` — Flag solar plants whose daily totals were placed in one night bucket.
- `GenerationQualityPolicy._daily_aggregate_profile_evidence(self, frame: pd.DataFrame)`
- `GenerationQualityPolicy._quality_codes(frame: pd.DataFrame)`
- `PlantQualityProfiler` — Summarize sensor risk per plant without declaring equipment failure.
- `PlantQualityProfiler.__init__(self, policy: GenerationQualityPolicy | None=None)`
- `PlantQualityProfiler.profile(self, frame: pd.DataFrame)`
- `PlantQualityProfiler._positive_flatline_rate(self, group: pd.DataFrame)`
- `PlantQualityProfiler._conditional_rate(flag: pd.Series, condition: pd.Series)`
- `PlantQualityProfiler._temporal_profile_consistency(frame: pd.DataFrame)`
- `PlantQualityProfiler._peer_pattern_consistency(frame: pd.DataFrame)`
- `PlantQualityProfiler._sensor_risk(row: dict[str, object])`
- `PlantQualityProfiler._recommendation(row: dict[str, object])`
- `QualityAuditResult`
- `QualityAuditService` — Application service that persists plant diagnostics and their thresholds.
- `QualityAuditService.__init__(self, profiler: PlantQualityProfiler | None=None)`
- `QualityAuditService.run(self, frame: pd.DataFrame, output_dir: Path, *, reference_paths: Iterable[Path] | None=None, artifact_name: str='plant_quality')`
- `QualityAuditService._reference_flatline_evidence(self, report: pd.DataFrame, reference_paths: Iterable[Path])`

## [reporting/__init__.py](../src/solar_forecast/reporting/__init__.py)

reporting 패키지의 공개 import 경계; 실행은 명시적 명령에서 시작

- `__getattr__(name: str)` — Load a public symbol only when its owning feature is requested.

## [reporting/benchmark_analytics.py](../src/solar_forecast/reporting/benchmark_analytics.py)

완료된 벤치마크의 무결성을 확인하고 선택 근거와 예측을 화면에 전달

- `BenchmarkAnalyticsService` — Keep benchmark selection evidence separate from legacy anomaly policies.
- `BenchmarkAnalyticsService.__init__(self, project_root: Path)`
- `BenchmarkAnalyticsService.build(self)`
- `BenchmarkAnalyticsService._task(self, run_dir: Path, manifest: dict[str, Any], task: dict[str, Any])`
- `_read_json(path: Path)`
- `_contained_path(root: Path, value: str)`
- `_validate_predictions(frame: pd.DataFrame, horizon: int, selected: str, periods: dict[str, Any])`
- `_validate_metric_tree(metrics: dict[str, Any])`
- `_verify_metric_values(actual: pd.DataFrame, column: str, expected: dict[str, Any])`
- `_verify_test_metrics(frame: pd.DataFrame, column: str, metrics: dict[str, Any])`
- `_recent_series(frame: pd.DataFrame)`

## [reporting/dashboard_builder.py](../src/solar_forecast/reporting/dashboard_builder.py)

registry·품질·모델 결과를 대시보드 JSON과 정적 게시 파일로 구성

- `DashboardBuildResult`
- `DashboardBuildResult.mapping_report(self)` — Compatibility alias for callers using the retired report name.
- `DashboardBuilder` — Publish nationwide inventory and model analytics as a clean public view.
- `DashboardBuilder.__init__(self, project_root: Path, output_dir: Path | None=None)`
- `DashboardBuilder.build(self)`
- `DashboardBuilder._public_inventory(inventory: dict[str, Any])` — Remove hashes, local paths, encoding and audit internals from the UI payload.
- `DashboardBuilder._write_payload(self, payload: dict[str, Any])`
- `DashboardBuilder._publish_static_assets(self)` — Copy the canonical dashboard shell when publishing elsewhere.
- `DashboardBuilder._copy_province_boundaries(self, national_inventory: dict[str, Any])`

## [reporting/model_analytics.py](../src/solar_forecast/reporting/model_analytics.py)

완료된 적격 모델 실행의 비교 지표·예측 시계열·운영 이상치 투영

- `_ModelRun`
- `_MetricAccumulator`
- `_MetricAccumulator.add(self, actual: float, predicted: float, capacity_mw: float | None)`
- `_MetricAccumulator.metrics(self)`
- `_CalibrationPolicy` — Frozen calibration thresholds with plant-aware, unit-safe fallbacks.
- `_CalibrationPolicy.evaluate(self, *, plant_id: str, absolute_error: float, capacity_mw: float | None)`
- `_CalibrationPolicy._absolute_evaluation(self, *, plant_id: str, absolute_error: float, source_suffix: str, allow_global: bool=True)`
- `ModelAnalyticsService` — Build a compact public contract from approved evaluation artifacts.
- `ModelAnalyticsService.__init__(self, project_root: Path)`
- `ModelAnalyticsService.build(self)`
- `ModelAnalyticsService._aligned_analysis(self, runs: list[_ModelRun])`
- `ModelAnalyticsService._comparison_signature(run: _ModelRun)`
- `ModelAnalyticsService._artifact_path(run: _ModelRun, key: str)`
- `ModelAnalyticsService._insert_predictions(self, connection: sqlite3.Connection, path: Path, model_id: str)`
- `ModelAnalyticsService._calibration_threshold(self, connection: sqlite3.Connection, path: Path, model_id: str, capacity: dict[str, float], contamination: float=CALIBRATION_CONTAMINATION)`
- `ModelAnalyticsService._calibration_threshold.records()`
- `ModelAnalyticsService._residual_quantile(connection: sqlite3.Connection, *, model_id: str, basis: str, contamination: float, minimum_samples: int, plant_id: str | None=None)`
- `ModelAnalyticsService._plant_residual_quantiles(self, connection: sqlite3.Connection, *, model_id: str, basis: str, contamination: float)`
- `ModelAnalyticsService._aggregate_aligned_rows(self, connection: sqlite3.Connection, reference_run: _ModelRun, thresholds: dict[str, _CalibrationPolicy], capacity: dict[str, float])`
- `ModelAnalyticsService._prediction_summary(*, evaluated_by_model: Counter[str], evaluated_by_region: Counter[str], evaluated_by_plant: Counter[tuple[str, str, str]], signals_by_model: Counter[str], signals_by_region: Counter[str], signals_by_plant: Counter[tuple[str, str, str]], returned_top_events: int)`
- `ModelAnalyticsService._prediction_summary.rate(signals: int, evaluated: int)`
- `ModelAnalyticsService._empty_prediction_summary()`
- `ModelAnalyticsService._capacity_lookup(self)`
- `ModelAnalyticsService._prediction_column(columns: Iterable[str], model_id: str)`
- `ModelAnalyticsService._normalize_timestamps(values: pd.Series)`
- `ModelAnalyticsService._full_runs(self)`
- `ModelAnalyticsService._latest_runs(runs: list[_ModelRun])`
- `ModelAnalyticsService._latest_compatible_pair(self, runs: list[_ModelRun])`
- `ModelAnalyticsService._is_full_run(details: dict[str, Any])`
- `ModelAnalyticsService._model_summary(self, run: _ModelRun)`
- `ModelAnalyticsService._evaluation_summary(self, runs: list[_ModelRun], comparable: set[str])`
- `ModelAnalyticsService._data_quality_signals(self)`
- `ModelAnalyticsService._quality_signal(self, row: pd.Series)`
- `ModelAnalyticsService._read_json(path: Path)`
- `ModelAnalyticsService._text(value: object)`
- `ModelAnalyticsService._number(value: object)`
- `ModelAnalyticsService._truthy(value: object)`
- `ModelAnalyticsService._integer(value: object)`
- `ModelAnalyticsService._at_least(value: float | None, threshold: float)`
- `ModelAnalyticsService._below(value: float | None, threshold: float)`

## [reporting/national_solar_inventory.py](../src/solar_forecast/reporting/national_solar_inventory.py)

공식 전국 태양광 원장의 무결성·행정구역·용량·중복 집계

- `InventoryConfigurationError` — Raised when source configuration is incomplete or inconsistent.
- `InventorySchemaError` — Raised when the EPSIS CSV no longer satisfies its column contract.
- `InventoryIntegrityError` — Raised when a local artifact does not match its configured digest.
- `InventoryContentError` — Raised when rows violate the configured solar-inventory contract.
- `AdministrativeRegionReference` — Effective-dated administrative names used independently of EPSIS facts.
- `AdministrativeRegionReference.default(cls)`
- `AdministrativeRegionReference.from_json(cls, path: str | Path)`
- `NationalInventoryConfig` — Resolved source metadata used by the repository and output lineage.
- `NationalInventoryConfig.from_mapping(cls, values: Mapping[str, Any], *, project_root: str | Path='.')`
- `NationalInventoryConfig.from_json(cls, path: str | Path, *, project_root: str | Path | None=None)`
- `InventoryCsvRecord`
- `ReviewedLocationOverride` — A source-hash-bound, human-reviewed correction for immutable CSV rows.
- `ReviewedLocationOverrideSet` — Validate and audit reviewed row corrections before regional aggregation.
- `ReviewedLocationOverrideSet.__init__(self, *, source_path: Path, source_path_label: str, entries: Sequence[ReviewedLocationOverride])`
- `ReviewedLocationOverrideSet.from_json(cls, path: str | Path, *, path_label: str, config: NationalInventoryConfig, actual_source_sha256: str, region_reference: AdministrativeRegionReference)`
- `ReviewedLocationOverrideSet.apply(self, record: InventoryCsvRecord, capacity_mw: Decimal | None)`
- `ReviewedLocationOverrideSet.validate_complete(self)`
- `KpxEpsisSolarInventoryRepository` — Verify and stream an EPSIS CSV without materialising it as a DataFrame.
- `KpxEpsisSolarInventoryRepository.__init__(self, config: NationalInventoryConfig)`
- `KpxEpsisSolarInventoryRepository.verify_sha256(self)`
- `KpxEpsisSolarInventoryRepository.stream(self)` — Yield validated columns and a lazy row iterator while the file is open.
- `KpxEpsisSolarInventoryRepository.stream.records()`
- `_Aggregate`
- `_Aggregate.add(self, capacity_mw: Decimal | None)`
- `_LocationAggregate`
- `_DiskDuplicateTracker` — Exact row tracking backed by temporary SQLite, keeping RAM bounded.
- `_DiskDuplicateTracker.__enter__(self)`
- `_DiskDuplicateTracker.is_duplicate(self, values: tuple[str, ...])`
- `_DiskDuplicateTracker.__exit__(self, exc_type: Any, exc: Any, traceback: Any)`
- `NationalInventoryService` — Aggregate national/regional totals and attach auditable map coordinates.
- `NationalInventoryService.__init__(self, repository: KpxEpsisSolarInventoryRepository, *, coordinate_cache: Mapping[str, Sequence[float]] | None=None, region_reference: AdministrativeRegionReference | None=None)`
- `NationalInventoryService.from_config(cls, config: Mapping[str, Any] | str | Path, *, project_root: str | Path='.', coordinate_cache: Mapping[str, Sequence[float]] | None=None, region_reference: AdministrativeRegionReference | None=None)`
- `NationalInventoryService.build(self)`
- `NationalInventoryService._accumulate_record(self, record: InventoryCsvRecord, *, total: _Aggregate, region_totals: dict[str, _Aggregate], location_totals: dict[tuple[str, str], _LocationAggregate], missing_cells: Counter[str], duplicates: _DiskDuplicateTracker, location_overrides: ReviewedLocationOverrideSet | None)`
- `NationalInventoryService._load_location_overrides(self, actual_source_sha256: str)`
- `NationalInventoryService._load_coordinate_cache(self)`
- `NationalInventoryService._is_footer(record: InventoryCsvRecord)`
- `NationalInventoryService._footer_quality(footer: InventoryCsvRecord | None, total: _Aggregate)`
- `NationalInventoryService._region_records(totals: Mapping[str, _Aggregate], canonical_regions: Sequence[str])`
- `NationalInventoryService._location_records(self, totals: Mapping[tuple[str, str], _LocationAggregate], cache: Mapping[str, Sequence[float]])`
- `build_national_inventory(project_root: str | Path, *, config_path: str | Path='config/national_solar_inventory.json')` — Convenience entry point for dashboard builders and command-line jobs.
- `canonical_region(value: object, *, aliases: Mapping[str, str]=REGION_ALIASES)`
- `canonical_location(value: object, region: str, *, aliases: Mapping[str, str]=REGION_ALIASES, canonical_regions: Sequence[str]=CANONICAL_REGIONS)`
- `normalize_location_key(value: object, region: str | None=None, *, aliases: Mapping[str, str]=REGION_ALIASES, canonical_regions: Sequence[str]=CANONICAL_REGIONS)`
- `source_region_conflict(value: object, region: str, *, aliases: Mapping[str, str]=REGION_ALIASES)` — Return true only when both source fields name different provinces.
- `_first_location_token(value: object)`
- `_normalize_inventory_text(value: object)`
- `_parse_inventory_decimal(value: object)`
- `_parse_inventory_integer(value: object)`
- `_decimal_to_float(value: Decimal | None)`
- `_parse_korean_coordinates(value: Sequence[float] | object)`

## [reporting/sgis_boundaries.py](../src/solar_forecast/reporting/sgis_boundaries.py)

공식 SGIS 원본 hash 검증과 지도 GeoJSON 경계 변환

- `ProvinceBoundaryError` — Raised when a province boundary source or output is unsafe to publish.
- `BoundaryLandmark` — A stable interior point used to guard administrative ownership.
- `SgisBoundarySource` — Auditable metadata embedded in a converted SGIS boundary artifact.
- `SgisBoundarySource.__post_init__(self)`
- `SgisProvinceBoundaryConverter` — Convert the official SGIS province Shapefile to compact WGS84 GeoJSON.
- `SgisProvinceBoundaryConverter.__init__(self, *, simplify_meters: float=150.0, precision: int=6)`
- `SgisProvinceBoundaryConverter.convert(self, source_path: str | Path, output_path: str | Path, *, source_archive_path: str | Path, source: SgisBoundarySource)` — Convert, validate and atomically write one province FeatureCollection.
- `SgisProvinceBoundaryConverter._geo_dependencies()`
- `verify_sgis_source_bundle(source_path: str | Path, source_archive_path: str | Path, *, source: SgisBoundarySource)` — Verify the official archive and every Shapefile component used.
- `validate_province_boundaries(payload: Mapping[str, Any], *, require_complete: bool=True)` — Validate GeoJSON structure, unique IDs and representative island owners.
- `geometry_contains_point(geometry: Mapping[str, Any], longitude: float, latitude: float)` — Return whether a WGS84 Polygon/MultiPolygon contains a point.
- `_polygon_contains_point(polygon: Sequence[Any], longitude: float, latitude: float)`
- `_ring_contains_point(ring: Sequence[Sequence[float]], longitude: float, latitude: float)`
- `_polygonal_geometry(geometry: Any, shapely_ops: Any)`
- `_rounded_mapping(value: Any, precision: int)`
- `_file_sha256(path: Path)`
- `_zip_member_sha256(archive: ZipFile, member: Any)`

## 대시보드 원본

`dashboard/src/`를 수정하고 빌드 도구로 `dashboard/assets/dashboard.js`를 생성합니다.

### [dashboard/src/analysis.js](../dashboard/src/analysis.js)

`renderAnalysisPage`, `predictionEvents`, `qualitySignals`, `normalizedSummaryRows`, `predictionSummary`, `regionsForAnalysis`, `plantsForAnalysis`, `renderAnalysisRegionOptions`, `renderAnalysisPlantOptions`, `bindAnalysisEvents`, `bindAnalysisLinks`, `metricValue`, `metricText`, `metricHint`, `renderComparisonView`, `renderMetricPanel`, `renderModelBars`, `renderModelTable`, `renderPerformanceView`, `renderRegionMatrix`, `renderMetricQuartet`, `renderPlantTable`, `renderAnomalyView`, `signalTypes`, `filteredPredictionEvents`, `filteredQualitySignals`, `selectedPredictionCount`, `combinedTypeCounts`, `renderTypeBars`, `combinedRegionCounts`, `renderSignalRegionBars`, `severity`, `signalMetricText`, `representativeEventCaption`, `thresholdSourceLabel`, `renderPredictionEventTable`, `renderQualitySignalTable`

### [dashboard/src/benchmark.js](../dashboard/src/benchmark.js)

`renderBenchmarkSection`, `bindBenchmarkEvents`, `benchmarkPeriod`, `renderBenchmarkTask`, `renderBenchmarkOptimization`, `renderBenchmarkRegionTable`

### [dashboard/src/bootstrap.js](../dashboard/src/bootstrap.js)

화면 상태 초기화 또는 데이터 로딩 진입부.

### [dashboard/src/context.js](../dashboard/src/context.js)

화면 상태 초기화 또는 데이터 로딩 진입부.

### [dashboard/src/coverage.js](../dashboard/src/coverage.js)

`renderCoveragePage`, `renderRegionOptions`, `nationalValue`, `nationalText`, `regionList`, `cleanSubregion`, `regionSearchTerms`, `normalizedSearch`, `administrativeAliasMatch`, `locationSearchTerms`, `renderDetailDirectionOptions`, `detailNumericValue`, `compareDetailRows`, `detailSortDescription`, `detailAriaSort`, `detailSortHeader`, `detailRows`, `renderDetailTable`, `bindCoverageEvents`, `bindRows`, `renderList`, `renderDetail`, `selectRegion`, `scheduleSearchRender`, `drawMap`

### [dashboard/src/forecast.js](../dashboard/src/forecast.js)

`forecastModels`, `renderForecastModelOptions`, `regionsForForecast`, `renderForecastRegionOptions`, `plantsForForecast`, `renderForecastPlantOptions`, `renderForecastPage`, `bindForecastEvents`, `renderForecastView`

### [dashboard/src/formatting.js](../dashboard/src/formatting.js)

`escapeHtml`, `isFiniteValue`, `formatNumber`, `formatPercent`, `renderPageIntro`, `renderSummaryStats`, `renderEmptyState`, `displayRegion`, `shortRegion`, `emptyAnalysis`, `formatDate`, `evaluationContext`, `normalizeSeries`, `renderLineChart`, `shortTime`

## 개발 도구와 검증 파일

### [tools/build_dashboard_assets.py](../tools/build_dashboard_assets.py)

Build the browser bundle from named view sources, without a Node dependency.

`build_dashboard_bundle`, `main`

### [tools/check_structure.py](../tools/check_structure.py)

Check module ownership, naming, imports, generated assets, and source/artifact separation.

`module_exports`, `check_structure`, `main`

### [tools/run_observed_benchmark_pilot.py](../tools/run_observed_benchmark_pilot.py)

Run a bounded benchmark on a measured plant-year from retained public archives.

`sha256_file`, `write_json`, `continuous_evaluation_coverage`, `select_observed_plant_year`, `prepare_official_gold`, `capture_input_provenance`, `configure_cpu_runtime`, `runtime_versions`, `build_pilot_config`, `main`

### [tools/update_code_index.py](../tools/update_code_index.py)

Generate the complete Python symbol and dashboard source index from real files.

`iter_symbols`, `render_code_index`, `main`

### [tools/verify_benchmark_model_artifacts.py](../tools/verify_benchmark_model_artifacts.py)

Replay selected benchmark checkpoints on observed Test rows without training.

`_sha256`, `_json`, `_resolve`, `_load_observations`, `_test_start`, `_indexed`, `_match_truth`, `_xgboost_replay`, `_cnn_replay`, `verify_selected_artifacts`, `main`

### [tools/verify_model_reorganization.py](../tools/verify_model_reorganization.py)

Compare two checkouts using frozen data and isolated, small CPU training runs.

`write_json`, `sha256_file`, `create_output_directory`, `create_synthetic_fixture`, `experiment_config`, `revision`, `source_identity`, `runtime_versions`, `check_artifact_compatibility`, `run_worker`, `execute_worker`, `run_comparison`, `main`

### [tests/test_anomaly_policy.py](../tests/test_anomaly_policy.py)

자동 검증 파일.

`test_supported_influence_factor_is_accepted`, `test_equipment_failure_claim_is_rejected`

### [tests/test_benchmark_dashboard.py](../tests/test_benchmark_dashboard.py)

Projection integrity tests; fixture scores are not PV accuracy evidence.

`BenchmarkDashboardTests`, `BenchmarkDashboardTests.setUp`, `BenchmarkDashboardTests.predictions`, `BenchmarkDashboardTests.write_manifest`, `BenchmarkDashboardTests.write_selection`, `BenchmarkDashboardTests.result`, `BenchmarkDashboardTests.prediction_path`, `BenchmarkDashboardTests.test_completed_artifact_projects_selection_test_series_and_scope`, `BenchmarkDashboardTests.test_smoke_and_incomplete_runs_are_excluded`, `BenchmarkDashboardTests.test_changed_prediction_file_is_excluded`, `BenchmarkDashboardTests.test_dataset_horizon_and_task_contracts_must_match`, `BenchmarkDashboardTests.test_wrong_origin_and_selected_predictions_rejected_even_with_updated_hash`, `BenchmarkDashboardTests.test_misreported_final_metrics_are_rejected`, `BenchmarkDashboardTests.test_selection_path_cannot_escape_run_directory`, `BenchmarkDashboardTests.test_complete_benchmark_can_be_relocated_without_rewriting_artifacts`, `BenchmarkDashboardTests.test_overlapping_selection_and_test_periods_are_rejected`, `BenchmarkDashboardTests.test_legacy_comparisons_require_matching_task_and_information`

### [tests/test_benchmark_job.py](../tests/test_benchmark_job.py)

Benchmark orchestration tests use labelled fixtures, never accuracy evidence.

`_predictions`, `_FixtureTrainingService`, `_FixtureTrainingService.run`, `BenchmarkJobTests`, `BenchmarkJobTests.test_base_selection_ignores_reversed_test_ranking`, `BenchmarkJobTests.test_alignment_rejects_truth_change_and_low_common_coverage`, `BenchmarkJobTests.test_alignment_reports_every_dropped_row`, `BenchmarkJobTests.test_legacy_experiment_schema_is_not_silently_executed`

### [tests/test_benchmark_model_selection.py](../tests/test_benchmark_model_selection.py)

Decision-contract tests use tiny fixtures, not real-model accuracy evidence.

`predictions`, `artifact_path`, `BenchmarkModelSelectionTests`, `BenchmarkModelSelectionTests.setUp`, `BenchmarkModelSelectionTests.run_selector`, `BenchmarkModelSelectionTests.test_hybrid_is_selected_on_later_calibration_and_test_cannot_reverse_it`, `BenchmarkModelSelectionTests.test_test_truth_changes_only_report_not_gate_or_selection`, `BenchmarkModelSelectionTests.test_gate_fit_excludes_selection_and_purged_hours`, `BenchmarkModelSelectionTests.test_gate_fit_excludes_selection_and_purged_hours.record_fit`, `BenchmarkModelSelectionTests.test_ties_keep_base_model_and_persistence_is_not_a_champion`, `BenchmarkModelSelectionTests.test_minimum_improvement_is_required`, `BenchmarkModelSelectionTests.test_group_metrics_sum_generation_and_do_not_clip_negative_r2`, `BenchmarkModelSelectionTests.test_constant_target_r2_is_json_null`, `BenchmarkModelSelectionTests.test_distinct_registry_ids_can_share_a_display_name`, `BenchmarkModelSelectionTests.test_output_contains_frozen_predictions_hashes_and_exact_provenance`, `BenchmarkModelSelectionTests.test_duplicate_or_nonfinite_or_missing_predictions_are_rejected`, `BenchmarkModelSelectionTests.test_horizon_and_truth_mismatches_are_rejected`, `BenchmarkModelSelectionTests.test_calibration_and_test_origin_overlap_is_rejected`, `BenchmarkModelSelectionTests.test_short_calibration_cannot_skip_the_purge`, `BenchmarkModelSelectionTests.test_existing_evidence_is_not_overwritten`

### [tests/test_candidate_intake.py](../tests/test_candidate_intake.py)

자동 검증 파일.

`_krc_row`, `test_krc_normalizer_maps_entity_capacity_and_hourly_mwh`, `test_krc_normalizer_quarantines_files_without_entity_column`, `test_candidate_intake_separates_generation_gate_from_weather_admission`

### [tests/test_cnn_bilstm_workflow.py](../tests/test_cnn_bilstm_workflow.py)

자동 검증 파일.

`_dummy_frame`, `_short_seq_config`, `test_train_and_save_creates_timestamped_dir`, `test_compare_checkpoints_reads_nested_runs`, `test_evaluate_and_analyze_saves_outputs`, `test_entity_sequences_never_cross_plants_and_split_chronologically`, `test_anomaly_threshold_is_frozen_from_calibration_not_test_ranking`, `test_lazy_windows_keep_all_missing_train_feature_as_zero_plus_mask`, `test_imputation_owns_buffer_and_preserves_input`, `test_historical_cnn_context_matches_tabular_forecast_and_excludes_future_inputs`, `test_historical_lookbacks_keep_common_split_calendar`

### [tests/test_collector_admission.py](../tests/test_collector_admission.py)

자동 검증 파일.

`test_collected_generation_admission_accepts_only_plant_hour_contract`

### [tests/test_dashboard_builder.py](../tests/test_dashboard_builder.py)

자동 검증 파일.

`_write_national_inventory_fixture`, `_write_static_dashboard`, `test_dashboard_builder_fails_fast_when_required_map_boundary_is_missing`, `_write_model_run`, `test_serve_dashboard_defaults_to_expected_local_url`, `test_dashboard_builder_publishes_clean_user_contract_without_registry`, `test_model_analytics_excludes_smoke_and_requires_matching_contract`, `test_model_analytics_excludes_smoke_and_requires_matching_contract.manifest`, `test_model_analytics_uses_plant_calibration_and_unit_safe_event_ranking`, `test_model_analytics_falls_back_to_global_capacity_normalized_threshold`, `test_model_analytics_caps_events_but_counts_all_and_keeps_latest_hourly_segment`, `test_model_analytics_selects_latest_complete_compatible_signature`, `test_dashboard_frontend_has_no_retired_developer_quality_view`, `test_dashboard_frontend_compares_all_metrics_and_supports_national_search`

### [tests/test_dataset_evidence.py](../tests/test_dataset_evidence.py)

Gold construction must not invent feature-selection performance evidence.

`test_rebuilt_gold_never_reuses_historical_ablation_as_current_evidence`

### [tests/test_ensemble.py](../tests/test_ensemble.py)

자동 검증 파일.

`_predictions`, `test_region_blend_learns_different_model_by_region`, `test_aggregation_recomputes_region_rmse_from_rows`, `test_dynamic_gate_records_context_and_reason`, `test_korean_source_columns_are_normalized`, `test_xgboost_prediction_artifact_keeps_stable_alignment_key`

### [tests/test_error_report.py](../tests/test_error_report.py)

자동 검증 파일.

`test_error_report_contains_stage_context_and_traceback`

### [tests/test_forecast_pipeline.py](../tests/test_forecast_pipeline.py)

자동 검증 파일.

`test_preprocessing_selects_numeric_and_cleans_rows`, `test_train_fitted_imputer_does_not_use_future_values`, `test_latest_file_discovery`, `test_training_loader_pushes_filters_and_float32_conversion_into_chunks`, `test_forecast_model_quality_gate_cannot_be_disabled_or_replaced`, `test_feature_ablation_fails_closed_without_training_eligibility`, `test_full_pipeline_creates_report`

### [tests/test_forecast_samples.py](../tests/test_forecast_samples.py)

자동 검증 파일.

`observations`, `test_forecast_uses_exact_origin_weather_and_target_generation_per_plant`, `test_missing_origin_is_excluded_without_row_shift_or_cross_plant_fill`, `test_future_weather_mutation_cannot_change_an_earlier_forecast_input`, `test_invalid_horizon_is_rejected`, `test_duplicate_plant_hour_is_rejected`, `test_cnn_window_ends_at_origin_inclusive_and_excludes_gapped_history`, `test_contract_distinguishes_real_horizon_from_legacy_row_estimation`

### [tests/test_generation_collectors.py](../tests/test_generation_collectors.py)

자동 검증 파일.

`test_collection_config_rejects_reverse_date_range`, `test_source_catalog_uses_the_four_given_official_homepages`, `test_download_filename_uses_canonical_korean_rule`, `test_download_filename_sanitizes_path_characters`, `test_standardized_csv_is_atomic_utf8_sig`, `test_koen_bronze_move_preserves_provider_bytes`, `test_kma_hourly_chunks_never_exceed_one_year`, `test_kma_chunks_support_leap_day`, `test_koen_rejects_invalid_month_range`, `test_koen_normalizer_converts_24_hour_columns_to_rows`, `test_koen_normalizer_converts_downloaded_kwh_unit_to_mwh`, `test_koen_normalizer_corrects_legacy_mislabeled_mwh_values`, `test_ewp_normalizer_keeps_latest_duplicate_and_builds_time_features`, `test_daily_wide_normalizer_converts_source_units`, `test_daily_wide_normalizer_keeps_static_capacity_and_global_plant_id`, `test_kospo_interval_normalizer_corrects_decisive_wh_scale_header_defect`, `test_kospo_interval_normalizer_preserves_physically_plausible_kwh_values`, `test_daily_total_reconciliation_fails_closed_on_incomplete_source_row`, `test_actual_dasong_export_reconciles_both_daily_total_unit_segments`, `test_generation_manifest_records_capacity_based_unit_resolution`, `test_iwest_renewable_standardization_retains_wind_with_energy_source`, `test_public_coordinate_columns_are_corrected_only_for_korea_range_inversion`, `_XmlResponse`, `_XmlResponse.__init__`, `_XmlResponse.raise_for_status`, `_KomipoSession`, `_KomipoSession.__init__`, `_KomipoSession.get`, `test_komipo_collector_writes_resumable_bounded_bronze_partition`, `test_komipo_collector_preflights_daily_call_budget`

### [tests/test_generation_quality.py](../tests/test_generation_quality.py)

자동 검증 파일.

`_quality_frame`, `test_quality_policy_flags_negative_without_turning_it_into_zero`, `test_quality_profiler_reports_risk_without_claiming_failure`, `test_quality_audit_distinguishes_pipeline_flatline_from_official_raw`, `_single_bucket_solar`, `test_quality_policy_excludes_daily_total_placed_in_one_night_bucket`, `test_quality_policy_does_not_flag_single_daylight_bucket_as_daily_aggregate`, `test_daily_aggregate_gate_does_not_cross_energy_sources_with_same_plant_id`, `test_quality_policy_requires_enough_consistent_active_days_for_aggregate_gate`, `test_quality_manifest_records_daily_aggregate_hard_gate`

### [tests/test_history_features.py](../tests/test_history_features.py)

자동 검증 파일.

`_hourly_frame`, `test_history_features_use_only_values_available_24_hours_before_issue_time`, `test_lag_is_timestamp_based_when_an_hour_is_missing`, `test_history_engineering_rejects_duplicate_entity_timestamps`, `test_model_configs_match_the_selected_generated_columns`, `test_controlled_experiment_forbids_unavailable_subday_lags`

### [tests/test_job_contracts.py](../tests/test_job_contracts.py)

자동 검증 파일.

`test_job_contract_catalog_documents_modular_monolith_decision`, `test_job_contract_cli_prints_worker_ready_boundary`, `test_jobs_cli_lists_boundaries`, `test_training_manifest_includes_versioned_contract`, `test_collection_manifest_includes_versioned_contract`, `test_prepare_data_contract_declares_gold_model_ready_schema`

### [tests/test_kma_asos_api.py](../tests/test_kma_asos_api.py)

Offline ASOS boundary tests; synthetic observations never represent live coverage.

`observation`, `response_bytes`, `AsosApiTests`, `AsosApiTests.setUp`, `AsosApiTests.client`, `AsosApiTests.client.transport`, `AsosApiTests.test_decoding_key_is_url_encoded_once_and_query_is_explicit`, `AsosApiTests.test_paginated_collection_reuses_complete_hash_checked_partition`, `AsosApiTests.test_partial_observation_coverage_is_queried_again`, `AsosApiTests.test_budget_failure_does_not_publish_partial_weather_or_complete_manifest`, `AsosApiTests.test_cache_tampering_is_rejected_without_network_or_csv_changes`, `AsosApiTests.test_no_data_does_not_create_complete_manifest`, `AsosApiTests.test_wrong_station_duplicate_timestamp_and_changing_total_are_rejected`, `AsosApiTests.test_out_of_range_and_non_hourly_timestamps_are_rejected`, `AsosApiTests.test_invalid_provider_payloads_are_rejected`, `AsosApiTests.test_transport_traceback_does_not_expose_key_or_query`, `AsosApiTests.test_transport_traceback_does_not_expose_key_or_query.transport`, `AsosApiTests.test_annual_merge_preserves_other_stations_unknown_columns_and_qc`, `AsosApiTests.test_missing_api_fields_do_not_erase_existing_values`, `AsosApiTests.test_api_and_browser_use_same_column_aware_merge`, `AsosApiTests.test_concurrent_weather_writer_is_rejected`, `AsosApiTests.test_mode_selection_and_missing_key_have_no_implicit_fallback`, `AsosApiTests.test_cli_collect_and_verification_expose_the_same_station_and_mode_options`, `AsosApiTests.test_collection_manifest_marks_annual_weather_as_silver`, `AsosApiTests.test_month_chunks_handle_leap_year_and_cross_year`, `AsosApiTests.test_asos_response_to_admission_and_gold_retains_observation_values`

### [tests/test_kospo_identity.py](../tests/test_kospo_identity.py)

Synthetic regression cases for collector identity; not production row counts.

`KospoIdentityTests`, `KospoIdentityTests.setUp`, `KospoIdentityTests.raw_frame`, `KospoIdentityTests.silver`, `KospoIdentityTests.write_silver`, `KospoIdentityTests.metadata`, `KospoIdentityTests.station_frame`, `KospoIdentityTests.reviewed`, `KospoIdentityTests.registry`, `KospoIdentityTests.test_new_collector_uses_generator_identity_without_changing_values`, `KospoIdentityTests.test_other_generators_are_not_mapped_by_department`, `KospoIdentityTests.test_generator_whitespace_is_normalized_but_suffixes_are_not_guessed`, `KospoIdentityTests.test_company_and_department_are_both_required`, `KospoIdentityTests.test_repeated_resolution_is_idempotent_and_preserves_meters`, `KospoIdentityTests.test_legacy_admission_records_rule_and_preserves_file_hash`, `KospoIdentityTests.test_conflicting_id_is_rejected_by_admission`, `KospoIdentityTests.test_identity_aliases_cannot_double_count_one_meter_in_one_file`, `KospoIdentityTests.test_stale_generator_id_cannot_be_silently_reassigned`, `KospoIdentityTests.test_historical_partitions_without_unit_are_unchanged`, `KospoIdentityTests.test_read_projection_still_rejects_missing_required_columns`, `KospoIdentityTests.test_real_address_does_not_bypass_weather_gate_or_infer_capacity`, `KospoIdentityTests.test_approved_legacy_and_new_snapshots_reconcile_without_double_count`, `KospoIdentityTests.test_synthetic_admission_to_gold_build_with_explicit_weather_review`

### [tests/test_model_checkpoints.py](../tests/test_model_checkpoints.py)

자동 검증 파일.

`_store`, `_frame`, `test_torch_checkpoint_is_atomic_and_rejects_wrong_signature`, `test_training_fingerprint_changes_with_data_but_not_runtime_budget`, `test_cnn_fixed_training_resumes_after_interrupted_epoch`, `test_cnn_fixed_training_resumes_after_interrupted_epoch.interrupted`, `test_cnn_fixed_training_resumes_after_interrupted_epoch.resumed`, `test_cnn_legacy_workflow_persists_study_and_final_state`, `test_adaptive_training_restores_bandit_and_epoch_progress`, `test_adaptive_training_restores_bandit_and_epoch_progress.interrupted`, `test_adaptive_training_restores_bandit_and_epoch_progress.resumed`, `test_xgboost_continues_from_saved_boosting_round`, `test_xgboost_early_stopping_state_survives_completed_resume`, `test_xgboost_restores_early_stopping_wait_after_interruption`

### [tests/test_model_optimization.py](../tests/test_model_optimization.py)

자동 검증 파일.

`_settings`, `test_optuna_study_resumes_without_exceeding_max_total_trials`, `test_optuna_study_resumes_without_exceeding_max_total_trials.objective`, `test_optimizer_study_name_is_scoped_to_training_fingerprint`, `test_failed_checkpointed_trial_is_retried_with_the_same_parameters`, `test_failed_checkpointed_trial_is_retried_with_the_same_parameters.objective`, `test_xgboost_optimizer_uses_validation_and_persists_artifacts`, `test_cnn_optimizer_handles_missing_mask_dimension_and_saves_study`

### [tests/test_model_parity.py](../tests/test_model_parity.py)

Strict artifact comparisons must never manufacture model performance evidence.

`PredictionParityTests`, `PredictionParityTests.setUp`, `PredictionParityTests.write_candidate`, `PredictionParityTests.compare`, `PredictionParityTests.test_reordered_identical_rows_pass_and_preserve_string_ids`, `PredictionParityTests.test_changed_predictions_fail_with_pooled_metrics_and_group_deltas`, `PredictionParityTests.test_tolerance_is_applied_against_baseline`, `PredictionParityTests.test_relative_tolerance_cannot_use_candidate_as_reference`, `PredictionParityTests.test_invalid_tolerances_are_rejected`, `PredictionParityTests.test_duplicate_keys_are_rejected`, `PredictionParityTests.test_normalized_duplicate_timestamps_are_rejected`, `PredictionParityTests.test_missing_rows_and_same_count_different_keys_are_rejected`, `PredictionParityTests.test_different_observations_are_rejected_even_with_loose_tolerance`, `PredictionParityTests.test_nonfinite_or_nonnumeric_targets_and_predictions_are_rejected`, `PredictionParityTests.test_region_disagreement_or_missing_column_is_rejected`, `PredictionParityTests.test_absent_regions_are_explicitly_unavailable`, `PredictionParityTests.test_split_mismatch_is_rejected`, `PredictionParityTests.test_malformed_and_non_hourly_timestamps_are_rejected`, `PredictionParityTests.test_equivalent_aware_timestamps_align_without_guessing_naive_timezone`, `PredictionParityTests.test_mixed_timezone_modes_are_rejected`, `PredictionParityTests.test_constant_and_single_row_targets_have_null_r2`, `PredictionParityTests.test_missing_columns_empty_files_and_duplicate_headers_are_rejected`, `PredictionParityTests.test_empty_identifiers_are_rejected`

### [tests/test_national_inventory.py](../tests/test_national_inventory.py)

자동 검증 파일.

`_write_csv`, `_config`, `_write_location_overrides`, `_approved_override`, `test_cp949_inventory_keeps_duplicates_and_validates_footer_and_coordinates`, `test_source_region_conflict_is_flagged_without_silent_reassignment`, `test_reviewed_location_override_is_applied_before_aggregation`, `test_reviewed_location_override_fails_closed_on_contract_drift`, `test_reviewed_location_override_rejects_a_different_source_hash`, `test_production_location_review_contract_preserves_totals_and_nesting`, `test_effective_dated_reference_merges_legacy_gwangju_jeonnam_names`, `test_natural_place_aliases_do_not_reassign_generator_name_substrings`, `test_utf8_replacement_character_metric_and_sha256_failure`, `test_missing_required_column_is_rejected`, `test_non_solar_row_is_rejected_instead_of_polluting_national_totals`

### [tests/test_notifications.py](../tests/test_notifications.py)

자동 검증 파일.

`_event`, `_alert`, `_route`, `_Directory`, `_Directory.__init__`, `_Directory.resolve`, `_Provider`, `_Provider.__init__`, `_Provider.send`, `_queue`, `_live_settings`, `test_external_delivery_requires_environment_and_separate_live_gate`, `test_dashboard_evaluation_and_unreviewed_events_are_rejected`, `test_stable_event_id_distinguishes_multiple_signals_at_same_timestamp`, `test_data_quality_alert_cannot_be_routed_to_plant_manager`, `test_plant_manager_route_cannot_cross_a_plant_boundary`, `test_enqueue_is_idempotent_and_persists_no_phone_or_api_secret`, `test_dry_run_suppresses_without_resolving_contact_or_calling_provider`, `test_live_dispatch_records_provider_acceptance_and_not_delivery`, `test_transient_failure_uses_exponential_retry_then_accepts`, `test_ambiguous_provider_outcome_is_quarantined_without_automatic_retry`, `test_permanent_kakao_failure_falls_back_to_sms_with_distinct_key`, `test_expired_dispatch_lease_is_in_doubt_not_automatically_resent`, `test_failure_audit_redacts_phone_and_bearer_token`, `_Response`, `_Response.json`, `_Session`, `_Session.__init__`, `_Session.post`, `_solapi_settings`, `_provider_message`, `test_solapi_uses_v4_hmac_and_approved_alimtalk_template_contract`, `test_solapi_credentials_cannot_be_redirected_to_another_host`, `test_solapi_long_sms_fallback_is_sent_as_lms_text_message`, `test_solapi_ambiguous_transport_outcomes_are_not_classified_as_retryable`, `test_solapi_ambiguous_transport_outcomes_are_not_classified_as_retryable.Session`, `test_solapi_ambiguous_transport_outcomes_are_not_classified_as_retryable.Session.post`, `test_solapi_accepts_registered_service_sender_but_requires_mobile_recipient`, `test_mapping_directory_is_runtime_only_and_missing_contact_falls_back`, `test_operational_event_rejects_other_energy_sources_and_fault_claims`, `test_operational_event_accepts_self_contained_nested_detector_and_evidence`, `test_operational_notification_requires_clean_physical_units_and_ranges`, `test_operational_event_batch_verifies_hash_count_and_excludes_contacts`, `_write_routes`, `test_runtime_route_directory_resolves_contacts_only_for_live_dispatch`, `test_notify_anomalies_cli_defaults_to_preview_and_has_explicit_live_gate`, `test_notify_anomalies_cli_dry_run_validates_manifest_and_uses_separate_outbox`, `test_notify_anomalies_cli_validates_all_rows_before_mutating_outbox`, `test_notify_anomalies_cli_live_is_blocked_without_environment_gate`, `test_notify_anomalies_cli_rejects_any_custom_preview_outbox`

### [tests/test_plant_registry.py](../tests/test_plant_registry.py)

자동 검증 파일.

`_stations`, `test_administrative_region_is_separate_from_weather_station_name`, `test_station_catalog_quarantines_an_address_without_a_defensible_match`, `test_reviewed_alias_keeps_one_physical_plant_identity_and_filters_ess_capacity`, `test_reviewed_station_mapping_catalog_is_explicit_and_auditable`, `test_reviewed_station_mapping_requires_evidence_and_rationale`, `test_legacy_station_seed_is_audit_only_and_cannot_make_a_plant_eligible`, `test_reviewed_mapping_requires_one_station_record_to_cover_generation_dates`, `test_nationwide_builder_does_not_filter_to_legacy_company_or_date`, `test_nationwide_builder_rejects_unapproved_mapping_methods`, `test_nationwide_builder_rejects_null_reviewed_evidence`, `test_official_partitions_do_not_require_a_legacy_mapping_file`, `test_nationwide_builder_replaces_cumulative_revision_instead_of_summing_it`, `test_nationwide_builder_prefers_explicit_latest_snapshot_date`, `test_model_ready_partitions_are_replaced_by_company_and_year`

### [tests/test_prediction_comparison_cli.py](../tests/test_prediction_comparison_cli.py)

Exercise prediction comparison as a real process, including artifact safety.

`PredictionComparisonCliTests`, `PredictionComparisonCliTests.setUp`, `PredictionComparisonCliTests.write_predictions`, `PredictionComparisonCliTests.invoke`, `PredictionComparisonCliTests.test_success_writes_matching_stdout_and_report_without_changing_inputs`, `PredictionComparisonCliTests.test_failed_parity_exits_one_and_preserves_failure_evidence`, `PredictionComparisonCliTests.test_invalid_artifact_exits_nonzero_without_publishing_success`, `PredictionComparisonCliTests.test_report_cannot_overwrite_either_input`, `PredictionComparisonCliTests.test_report_symlink_cannot_alias_an_input`, `PredictionComparisonCliTests.test_report_temporary_path_cannot_overwrite_an_input`, `PredictionComparisonCliTests.test_invalid_tolerance_exits_nonzero`, `PredictionComparisonCliTests.test_cli_help_does_not_import_optional_training_dependencies`

### [tests/test_reorganization_harness.py](../tests/test_reorganization_harness.py)

Safety and fixture-contract checks independent of ML dependencies.

`ReorganizationHarnessTests`, `ReorganizationHarnessTests.test_source_identity_detects_uncommitted_code_and_configuration_changes`, `ReorganizationHarnessTests.test_existing_evidence_is_never_overwritten`, `ReorganizationHarnessTests.test_fixture_is_frozen_and_has_disjoint_entities_and_missingness`, `ReorganizationHarnessTests.test_small_budget_does_not_modify_original_configuration`

### [tests/test_sgis_boundaries.py](../tests/test_sgis_boundaries.py)

자동 검증 파일.

`_official_boundaries`, `test_official_sgis_boundaries_have_complete_unique_regions_and_island_owners`, `test_boundary_validation_rejects_duplicate_region_ids`, `test_prepare_boundaries_cli_requires_auditable_source_digests`, `test_sgis_source_bundle_verifies_archive_and_every_sidecar`

### [tests/test_temporal_split.py](../tests/test_temporal_split.py)

자동 검증 파일.

`test_global_timestamp_boundaries_are_shared_by_irregular_entities`, `test_purge_gap_removes_boundary_hours_from_all_entities`

### [tests/test_training_lock.py](../tests/test_training_lock.py)

자동 검증 파일.

`test_training_lock_blocks_second_model`, `test_completed_training_manifest_preserves_execution_mode`

### [tests/test_verification_job.py](../tests/test_verification_job.py)

자동 검증 파일.

`test_verify_e2e_cli_defaults_to_smoke_and_requires_collect_dates`, `test_verification_service_writes_auditable_report`, `test_verification_service_writes_auditable_report.FakePreparationService`, `test_verification_service_writes_auditable_report.FakePreparationService.__init__`, `test_verification_service_writes_auditable_report.FakePreparationService.run`, `test_verification_service_writes_auditable_report.FakeTrainingService`, `test_verification_service_writes_auditable_report.FakeTrainingService.run`, `test_verification_service_writes_auditable_report.FakeDashboardBuilder`, `test_verification_service_writes_auditable_report.FakeDashboardBuilder.__init__`, `test_verification_service_writes_auditable_report.FakeDashboardBuilder.build`
