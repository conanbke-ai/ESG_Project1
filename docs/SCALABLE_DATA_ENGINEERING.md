# 대량 데이터 취급 설계와 포트폴리오 설명 기준

이 프로젝트는 현재 단일 PC에서 처리 가능한 중간 규모 데이터이므로 분산 시스템을 사용했다고
과장하지 않습니다. 대신 데이터가 커져도 교체하기 쉬운 경계, 메모리 상한, lineage와 실패 격리를
코드로 보여주는 것을 목표로 합니다.

## 현재 구현된 데이터 흐름

```text
공식 원본(Bronze, 불변)
  → 출처별 adapter + 스키마 판별
  → 파일 단위 시간별 MWh 파티션(Silver, gzip)
  → 품질 flag + KMA 결합 + 공통 피처(Gold)
  → 전역 시간분할 Train/Validation/Calibration/Test
  → XGBoost native NaN / CNN Train-only 중앙값+mask+lazy window
  → 실행별 모델·예측·manifest
```

| 구현 | 기술적 의미 | 면접에서 검증 가능한 근거 |
|---|---|---|
| 원본과 표준본 분리 | 원천 재현성과 재처리 가능성 | `file/raw`, `file/standardized` 경계 |
| 스키마 registry/adapter | 회사별 컬럼 변경을 핵심 모델 코드에서 격리 | `collectors/normalization.py`, `archive.py` |
| 파일 단위 처리 | 88개 기본 파일과 4개 편입 후보 파일 전체를 한 번에 원본 DataFrame으로 만들지 않음 | source-file bounded 표준화 |
| Silver/Gold 파티션 | Silver는 원본 파일별, Gold는 회사×연도로 잘라 필요한 컬럼·행만 읽음 | generation/model-ready manifests |
| SHA-256·byte lineage | 같은 이름의 수정 파일과 입력 변경 추적 | partition별 source/output hash |
| atomic replace | 중간 실패 시 완성되지 않은 manifest/파티션 노출 방지 | `.tmp` 완성 후 replace |
| quarantine | 식별자·단위가 불명확한 파일을 억지로 병합하지 않음 | 영암 2020~2021 격리 기록 |
| 명시적 품질 flag | 보간값과 관측값, 물리 위반과 의심 패턴 분리 | `quality_*`, observation masks |
| 누수 방지 분할 | 모든 발전소에 동일 날짜 경계와 168시간 purge 적용 | `TemporalSplitter` |
| 메모리 제한 시퀀스 | `(행 × lookback)` 사전 복제를 피함 | `LazyWindowSequenceDataset` |
| bounded 학습 로더 | 10만 행 chunk, 필터 pushdown, float32, 1.5GB hard limit | `DatasetLoadPolicy`/`DatasetLoadReport` |
| bounded 하이퍼파라미터 탐색 | XGBoost 대표행·CNN lazy sequence 상한, trial/시간 예산, pruning 후 전체 Train 재학습 | 모델별 `optimizer` 설정·SQLite study |
| fingerprint 체크포인트 | 데이터/설정별 상태 격리, CNN epoch·XGBoost boosting round 재개, atomic replace | `TrainingCheckpointStore`와 실행 manifest |
| API 호출 예산 | 중부발전 요청 전 최소 호출 수를 계산하고 station/day로 원자 저장·재개 | `KomipoRenewableCollector` |
| 열 단위 CNN 통계 | Train 중앙값 계산 시 전체 훈련행×피처 행렬을 한 번 더 합치지 않음 | feature-wise median fitting |
| 프레임워크별 결측 처리 | XGBoost는 native NaN, CNN은 Train 통계+mask | split별 preprocessing manifest |

현재 검증 범위는 기본 공식 원본 88개와 농어촌공사 파일 6개입니다. 식별자 없는 2개 파일을
격리하고 92개 파일에서 4,341,756개 표준 시간 행을 만들었습니다. 발전소×발전원 registry 68개 중
행정·기상 매핑 근거를 통과한 Gold는 46개 발전소, 2,525,434행이며 22개는 quarantine했습니다.
이 수치들은 분산 빅데이터를 의미하지는 않지만,
파티셔닝·lineage·idempotent 재처리·품질 게이트를 실데이터로 설명하기에는 충분합니다.
편입 원본 28,833,903 byte가 시간별 long gzip 파티션 33,707,476 byte가 됐습니다. gzip인데도
원본보다 큰 이유는 일별 wide 행의 메타데이터가 24개 시간 행에 반복되기 때문이며, 이 측정은
Parquet columnar encoding 전환의 구체적인 근거로 사용합니다.

Gold 32개 회사×연도 gzip 파티션은 126,723,824 byte입니다. 실제 XGBoost 계약으로 전체
2,525,434행을 스캔해 태양광·품질 게이트를 통과한 2,401,306행만 유지했을 때, 26개 피처와
문맥 컬럼을 포함한 pandas DataFrame은 409.93MB였습니다. 모든 수치 피처·타깃은 `float32`이며
설정된 1,536MB 상한을 넘으면 일부 행을 임의 샘플링하지 않고 실패합니다. 이 값은 DataFrame
메모리 측정값이고 프로세스 peak RSS라고 과장하지 않습니다.

## 확장 시 교체 순서

1. CSV/gzip reader를 PyArrow Dataset 또는 Polars lazy scan으로 교체하고 Silver를 Parquet으로
   전환합니다. application service와 모델 계약은 유지합니다.
2. 로컬 분석이 수천만~수억 행으로 커지면 DuckDB/Polars predicate pushdown으로 필요한 기간과
   컬럼만 읽습니다.
3. 단일 노드 메모리·처리시간 한계를 실제 측정으로 넘을 때 Spark/Ray를 도입합니다. 단순히
   이력서 키워드를 늘리기 위해 분산 프레임워크를 먼저 넣지 않습니다.
4. 객체 저장소에는 원본 hash 기반 경로와 catalog를 두고, orchestration은 수집·표준화·품질·학습
   단계를 재시도 가능한 task로 분리합니다.
5. 데이터 계약 버전, 모델 버전, feature contract, split boundary를 MLflow/DVC 같은 registry에
   연결해 어떤 데이터로 어떤 모델을 만들었는지 역추적합니다.

## 성능 측정 항목

기업 프로젝트 설명에는 “대용량 처리 가능”이라는 문장보다 다음 수치를 남깁니다.

- 입력 파일 수·byte, 표준 행 수, 출력 byte와 압축률
- 단계별 wall time, peak RSS, rows/sec
- 변경된 파티션만 재처리했을 때의 절감률
- 스키마 실패·격리 파일 수와 이유
- 중복 키, 결측률, 최대 연속 공백, 발전소별 기간
- lazy window의 원본 배열 크기와 eager materialization 대비 예상 메모리
- Train/Validation/Calibration/Test 행 수와 경계

현재 manifest는 입력·출력 byte와 SHA-256, 행 수, 기간, 품질 요약을 기록하고 학습 결과에는
scanned/retained rows, 선택 컬럼, chunk 크기, dtype, DataFrame 메모리를 남깁니다. wall time과
프로세스 peak RSS 자동 계측, Parquet predicate pushdown, 변경 파티션 skip은 다음 데이터
엔지니어링 실험으로 분리합니다.

Optuna는 모델별 study를 SQLite에 저장하고 중단 후 남은 trial만 수행합니다. 튜닝 단계는 XGBoost
Train 750,000/Validation 250,000행, CNN Train 250,000/Validation 100,000 lazy sequence를 기본
상한으로 사용하지만, 최종 선택 모델은 전체 Train으로 다시 학습합니다. Calibration/Test를 탐색
표본으로 사용하지 않으므로 계산량 제한이 평가 누수로 이어지지 않습니다.

Optuna study 재개와 실제 모델 학습 재개는 분리했습니다. 전자는 trial 목록과 heartbeat를 SQLite에,
후자는 CNN 모델·optimizer·early stopping·RNG 또는 XGBoost Booster를 fingerprint 디렉터리에
저장합니다. 임시 파일을 완성한 뒤 atomic replace하며 Windows에서 백신·인덱서가 순간적으로 파일을
점유하는 경우에는 제한된 재시도를 적용합니다. 장애 시 손실되는 최대 작업량은 설정상 CNN 1 epoch,
XGBoost 50 boosting round입니다.

## 권장 포트폴리오 서술

> 서로 다른 공공기관의 wide CSV를 버전 있는 시간별 MWh 계약으로 표준화하고, 434만 Silver 행과
> 253만 Gold 행을 gzip 파티션으로 처리했습니다. 원본·산출물 SHA-256과 atomic write로
> lineage와 실패 안전성을 확보하고, 불확실한 22개 발전소는 quarantine했습니다. 학습 시에는
> 10만 행 chunk에서 컬럼·품질 필터를 먼저 적용하고 float32·1.5GB hard limit를 강제했으며,
> XGBoost native NaN과 CNN lazy window/missing mask로 결측과 메모리를 모델 특성에 맞게 처리했습니다.

이 문장은 현재 저장소에서 확인 가능한 구현만 포함합니다. 이후 Parquet/Polars/Spark를 도입하면
반드시 실제 benchmark와 함께 별도 수치로 갱신합니다.
