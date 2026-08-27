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
| 파일 단위 처리 | 86개 파일 전체를 한 번에 원본 DataFrame으로 만들지 않음 | source-file bounded 표준화 |
| 연도·회사 gzip 파티션 | 필요한 범위만 읽고 wide→long 반복값의 저장 팽창을 완화 | generation partition + manifest |
| SHA-256·byte lineage | 같은 이름의 수정 파일과 입력 변경 추적 | partition별 source/output hash |
| atomic replace | 중간 실패 시 완성되지 않은 manifest/파티션 노출 방지 | `.tmp` 완성 후 replace |
| quarantine | 식별자·단위가 불명확한 파일을 억지로 병합하지 않음 | 영암 2020~2021 격리 기록 |
| 명시적 품질 flag | 보간값과 관측값, 물리 위반과 의심 패턴 분리 | `quality_*`, observation masks |
| 누수 방지 분할 | 모든 발전소에 동일 날짜 경계와 168시간 purge 적용 | `TemporalSplitter` |
| 메모리 제한 시퀀스 | `(행 × lookback)` 사전 복제를 피함 | `LazyWindowSequenceDataset` |
| 프레임워크별 결측 처리 | XGBoost는 native NaN, CNN은 Train 통계+mask | split별 preprocessing manifest |

현재 검증 범위는 86개 공식 원본, 4,146,757개 표준 시간 행, 85개 표준 `plant_id`입니다. 태양광
학습용 Gold 테이블은 현재 18개 발전소 314,496행입니다. 이 수치들은 분산 빅데이터를 의미하지는
않지만, 파티셔닝·lineage·idempotent 재처리·품질 게이트를 실데이터로 설명하기에는 충분합니다.
현재 원본 27,630,962 byte가 시간별 long gzip 파티션 31,951,557 byte가 됐습니다. gzip인데도
원본보다 큰 이유는 일별 wide 행의 메타데이터가 24개 시간 행에 반복되기 때문이며, 이 측정은
Parquet columnar encoding 전환의 구체적인 근거로 사용합니다.

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

현재 manifest는 입력·출력 byte와 SHA-256, 행 수, 기간, 품질 요약을 기록합니다. wall time과 peak
RSS 자동 계측, Parquet predicate pushdown, 변경 파티션 skip은 다음 데이터 엔지니어링 실험으로
분리합니다.

## 권장 포트폴리오 서술

> 서로 다른 공공기관의 wide CSV를 버전 있는 시간별 MWh 계약으로 표준화하고, 415만 행을
> 파일 단위 gzip 파티션으로 처리했습니다. 원본·산출물 SHA-256과 atomic write로 lineage와 실패
> 안전성을 확보했으며, 식별자·단위가 불명확한 파일은 quarantine했습니다. 학습 단계에서는
> 전역 시점 분할과 purge gap으로 누수를 막고, XGBoost native NaN 및 CNN lazy window/missing
> mask로 결측과 메모리 사용을 모델 특성에 맞게 처리했습니다.

이 문장은 현재 저장소에서 확인 가능한 구현만 포함합니다. 이후 Parquet/Polars/Spark를 도입하면
반드시 실제 benchmark와 함께 별도 수치로 갱신합니다.
