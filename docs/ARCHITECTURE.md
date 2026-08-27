# Architecture

제품 코드는 `src/solar_forecast` 하나의 설치형 패키지로 관리합니다. 루트의 `app.py`는
개발 환경용 얇은 진입점이며 비즈니스 로직을 포함하지 않습니다.

## 설계 원칙

- Application service: 한 유스케이스의 흐름과 산출물 경계를 객체 하나가 소유합니다.
- Dependency injection: 파이프라인의 저장소·전처리·학습·보고 구현을 생성자에서 교체할 수 있습니다.
- Ports and adapters: 공식 HTTP 첨부, 브라우저 자동화, 파일, 모델 프레임워크 의존성을 서비스 경계 밖으로 격리합니다.
- Configuration as data: 모델과 실험 조건은 `config/` JSON에서 관리합니다.
- Reproducible artifacts: 실행별 디렉터리와 manifest에 입력·설정·결과를 기록합니다.
- Lineage and failure safety: 원본·파티션 SHA-256과 byte를 기록하고 임시 파일을 완성한 뒤
  atomic replace하여 중간 실패가 완성 산출물처럼 보이지 않게 합니다.
- Quarantine before admission: 발전소 식별자·단위·기상 매핑이 불확실한 추가 파일은 격리하고,
  발전량 품질 게이트와 모델 편입 승인을 분리합니다.
- Leakage safety: 모든 발전소에 같은 전역 시간 경계를 적용하고 Train/Validation/Calibration/Test를
  분리합니다. 결측 통계는 Train, Hybrid 판단 근거는 Validation, 잔차 임계값은 Calibration만
  사용하며 Test 실제값은 마지막 평가에만 사용합니다.
- Technology separation: 원본의 태양광·풍력·수력을 모두 보존하고 모델 유스케이스에서 발전원을 선택합니다.
- Quality before imputation: 물리 위반과 문맥상 의심을 flag로 분리한 뒤 필요한 구간만 복원합니다.

## 패키지 구조

```text
src/solar_forecast/
├─ cli.py                   # 단일 CLI
├─ pipeline/                # ForecastPipeline과 단계별 어댑터
├─ ensemble/                # ExplainableDynamicGate, HybridExperiment
├─ collectors/              # Collector 구현과 CollectionService
├─ features/                # ASOS 표준화와 누수 없는 시간·이력 피처
├─ quality/                 # 물리 규칙, 품질 flag, 발전소별 센서 위험 진단
├─ evaluation/              # 공통 TemporalSplitter와 rolling-origin 피처 ablation
├─ preparation.py           # 보관 원본 전체 표준화/학습파일 application service
├─ models/
│  ├─ cnn/                  # CNN-BiLSTM 학습 엔진
│  ├─ cnn_bilstm.py         # 독립 학습 어댑터
│  └─ xgboost.py            # 독립 학습 어댑터
├─ anomalies/               # 이상징후 해석 정책
├─ artifacts/               # manifest 저장
├─ infrastructure/          # 환경·오류·로깅 어댑터
└─ jobs/                    # 프로세스 간 학습 잠금
```

의존성은 CLI에서 application service로, service에서 명시적 adapter로 흐릅니다. 데이터 I/O나
파일 저장은 모델의 핵심 계산과 섞지 않습니다. 기존 호출자를 위한 함수형 API는 얇은 호환
facade로만 유지합니다.

## 주요 객체

- `ForecastPipeline`: 데이터 로드 → 전처리 → CNN 학습/평가 → 보고서 생성
- `CollectionService`: 출처별 collector 격리 실행 및 수집 manifest 생성
- `DataGoFileCollector`/`EwpTrainingDataCollector`/`KoenHomepageCollector`: 공식 게시물의 전송 방식별 adapter
- `DailyWideGenerationNormalizer`: 발전사별 wide CSV 스키마와 원본 단위를 주입받아 공통 시간별 MWh 계약으로 변환
- `GenerationSchemaRegistry`: 파일 컬럼 집합으로 여섯 가지 보관 원본 스키마를 명시적으로 판별
- `HistoricalGenerationStandardizationService`: 4개사 원본을 파일 단위 gzip 파티션과 검증 manifest로 변환
- `NationwidePlantRegistryBuilder`: 기관·발전원별 안정 키, 행정구역, 기상 매핑 근거와 quarantine 사유를 관리
- `NationwideModelDatasetBuilder`: 승인된 모든 기관·발전소를 공통 피처로 결합하고 누적 개정본을 최신
  snapshot 우선으로 조정한 뒤 회사×연도 Gold 파티션 생성
- `KrcYeongamCandidateIntakeService`: 추가 발전소 원본의 hash·스키마·연속성·4구간 분할을 검증하고
  개체 식별이 불가능한 과거파일을 quarantine
- `PlantMetadataCatalog`: 공식 설비현황을 발전소/호기 단위로 매칭하되 불확실한 총용량 복제를 금지
- `KmaAsosNormalizer`: 선택된 ASOS 기상요소와 관측지점 좌표를 영문 공통 계약으로 변환
- `LeakageSafeFeatureEngineer`: 실제 시각 차이를 기준으로 24/168시간 lag와 24시간 이동 7일 평균 생성
- `GenerationQualityPolicy`: 음수·용량초과·주간 0·flatline·기상 범위 위반을 원인별 flag로 분리
- `PlantQualityProfiler`: 발전소별 커버리지·형태 일관성·peer 상관을 보고하되 고장을 확정하지 않음
- `FeatureAblationService`: Calibration/Test를 예약한 purged rolling-origin 평균으로 후보 컬럼 승격 여부 결정
- `TemporalSplitter`: 발전소별 관측률과 무관하게 공통 날짜 경계·purge gap·독립 Calibration을 적용
- `LazyWindowSequenceDataset`: 발전소별 원본 배열을 한 번만 보관하고 batch에서 시퀀스를 절단
- `DatasetRepository.load_training_frame`: Gold 파티션에 컬럼 선택·발전원/품질 필터를 chunk 단계에서
  적용하고 float32 변환·명시적 메모리 예산을 강제
- `DataPreparationService`: 전체 표준화와 `model_ready.csv.gz`/`model_ready_parts` 생성을 하나의
  재현 가능한 유스케이스로 묶음
- `TrainingService`: 모델 전략 선택, 전역 학습 잠금, 성공/실패 manifest 관리
- `ExplainableDynamicGate`: 발전소·지역·시간·출력 regime·모델 불일치별 Validation 근거와 행별 동적 결합
- `HybridExperiment`: Hybrid 입력/평가/결과 파일의 유스케이스 경계

## 실행 및 산출물

```bash
python app.py pipeline "합산발전량(MWh)" --data file/merge_data/val.csv
python app.py train xgboost
python app.py train cnn_bilstm
python app.py hybrid validation_predictions.csv test_predictions.csv
python app.py collect --start-date 2024-01-01
python app.py prepare-data
python app.py evaluate-features
```

장시간 기본 모델 학습은 `artifacts/.training.lock`으로 직렬화합니다. Hybrid와 보고 작업은
저장된 예측을 사용하므로 기본 모델 학습과 분리할 수 있습니다.

## 데이터 계약

Hybrid 필수 컬럼은 `timestamp, region, plant, y_true, xgb_pred, cnn_pred`입니다.
`일시, 지역, 발전구분, 합산발전량(MWh)`는 자동 변환하지만 두 예측 컬럼은 모델이 실제로
생성한 값이어야 합니다. 판단 결과에는 사용한 범위, 모델별 예상 MAE, 선택 모델, 결합비와
문장형 근거가 함께 저장됩니다.

수집 계층은 원본을 수정하지 않고 `file/raw/<회사>/normalized/`에 UTF-8 표준본을 추가합니다.
보관 데이터 준비 계층은 `file/standardized/generation/<회사>/`에 gzip 파티션을 만듭니다.
개별 설비의 공통 키는 `timestamp + company + plant_id`, 동서발전 지역 집계 자료의 공통 키는
`timestamp + region`이며 정규화기가 컬럼·단위·중복 키를 실행 중 검증합니다. 공개 포털의 누적
snapshot끼리 같은 키가 겹치면 파일명의 `YYYYMMDD`가 최신인 값을 사용하고 합산하지 않습니다.

개별 설비 계약에는 `energy_source`가 포함됩니다. 풍력·수력 행도 저장하지만 태양광 모델 설정은
`energy_source_filter=solar`를 적용합니다. Gold 호환본과 파티션에는 품질 flag와 관측 mask를 함께
보존하고, 모델 기본 피처는 Validation ablation에서 개선된 컬럼만 선택합니다.

학습은 `model_ready_parts/company=<provider>/year=<YYYY>/part.csv.gz`를 100,000행 단위로 읽습니다.
필요 컬럼과 태양광·품질 조건을 먼저 적용하고 숫자는 `float32`로 축소합니다. DataFrame 메모리가
1,536MB를 넘으면 일부 행을 몰래 버리지 않고 즉시 실패해 운영자가 기간·기관 또는 실행 자원을
명시적으로 조정하게 합니다. CNN window는 이 프레임을 다시 `(rows × lookback)`으로 복제하지 않고
batch 시점에 lazy slicing합니다.

CNN-BiLSTM 시퀀스는 `plant_id`별로 따로 만들지만 분할 경계는 발전소별 행 비율이 아니라 전체
데이터의 같은 날짜를 사용합니다. Train/Validation/Calibration/Test 경계마다 168시간 purge
gap을 두며 서로 다른 발전소의 마지막 행과 첫 행을 한 시퀀스로 연결하지 않습니다. 168시간
window는 미리 복제하지 않고 batch에서 lazy materialization합니다.

모든 발전소의 모델 입력 피처 수와 순서는 동일합니다. `plant_id`는 현재 시퀀스 경계와 품질 집계에
사용되며 모델 입력에는 직접 들어가지 않습니다. 모델은 용량·설치각·좌표·관측소 메타데이터로
발전소 차이를 설명합니다. 향후 plant embedding은 기존 발전소 미래 Test와 unseen-plant group
holdout을 둘 다 개선할 때만 공통 입력 계약의 한 구성요소로 승격합니다.

단일 노드에서 검증한 현재 처리 규모와 Parquet/Polars/분산 처리 전환 기준은
[`SCALABLE_DATA_ENGINEERING.md`](SCALABLE_DATA_ENGINEERING.md)에 정리합니다.
