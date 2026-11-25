# 태양광 발전소 데이터 전처리 최종 정책

최종 확정일: 2026-09-18 (Asia/Seoul)  
Canonical config: `config/plant_data_preprocessing_policy.json`

## 목적

이 문서는 태양광 발전량 예측에 사용되는 발전소 데이터가 **어디서 보존되고, 어떤 기준으로 표준화·품질 판정·학습 편입되는지**를 최종 확정한다.

전처리의 목표는 데이터 수를 최대화하는 것이 아니라 다음 네 가지를 동시에 만족하는 것이다.

1. 공식 원본과 서비스 전체 현황을 훼손하지 않는다.
2. 모델 타깃·기상 입력·시간축의 의미를 명확히 한다.
3. 모든 발전소에 동일한 시간 평가 기준을 적용한다.
4. 제외·보류 이유와 입력 lineage를 재현 가능하게 남긴다.

이 정책 이후 전처리 동작을 바꾸려면 정책 버전, 관련 contract, 테스트와 포트폴리오 설명을 함께 변경한다. 코드 한 곳의 임의 수정으로 기준을 바꾸지 않는다.

---

## 1. 서비스 모집단과 AI 모델 모집단

두 모집단은 끝까지 분리한다.

### 서비스 전체 발전소 현황

- EPSIS 공식 공개 전국 태양광 등록 범위를 사용한다.
- AI 학습 가능 여부 때문에 설비·등록행을 제거하지 않는다.
- 공식 고유키가 없으면 중복처럼 보이는 행도 임의 삭제하지 않는다.
- 지역명 표준화와 검토된 위치 보정은 표시·집계용이며 원본 총량을 임의 축소하지 않는다.

### AI 모델 모집단

- 발전량 예측에 필요한 데이터 적격성 게이트를 별도로 적용한다.
- 학습 제외는 원본 삭제가 아니라 `model eligibility` 상태다.
- 학습에서 제외된 발전소도 서비스 전체 현황에서는 계속 유지된다.

---

## 2. Canonical preprocessing pipeline

```text
Provider source / official archive
        |
        v
[1] Source preservation
        |
        v
[2] Provider schema + unit standardization
        |
        v
[3] Plant identity / provenance / ASOS mapping
        |
        v
[4] Gold generation-weather join
        |
        v
[5] Quality annotation
        |
        v
[6] Leakage-safe history/features
        |
        v
[7] Training eligibility
        |
        v
[8] Training admission
        |
        v
[9] Materialized ADMITTED dataset
        |
        v
[10] Model-specific forecast readiness
        |
        v
XGBoost / CNN-BiLSTM / Hybrid benchmark
```

### [1] Source preservation

Owner: `collectors/`, `file/raw/`, 기존 공식 archive.

- 공급기관의 byte, 파일명, 인코딩, source path, hash를 보존한다.
- Bronze 원본을 보기 좋게 만들기 위해 재인코딩·rename·값 보정하지 않는다.
- 이후 단계의 모든 정제는 파생 데이터에서 수행한다.

### [2] Standardization

Owner: `datasets/`.

공급기관마다 다른 wide/long schema와 단위를 공통 plant-hour 계약으로 변환한다.

```text
timestamp
company
plant_id
plant
unit
energy_source
generation_mwh
capacity_mw
tilt_deg
latitude
longitude
address
source_file
```

- Wh/kWh/MWh는 검증된 규칙으로 MWh에 맞춘다.
- 애매한 단위는 물리 상한·공식 합계 등 독립 근거 없이는 추정하지 않는다.
- 발전소 식별자가 없는 지역합계/집계 파일은 개별 발전소 학습 데이터로 승격하지 않는다.
- 수정 snapshot은 합산하지 않고 명시된 최신 snapshot 우선 규칙으로 reconcile한다.

### [3] Plant identity / weather mapping

Owner: `datasets.plant_registry`.

- `company + plant + energy_source`의 안정 식별자를 사용한다.
- 행정구역과 ASOS 관측소는 별도 개념으로 관리한다.
- 공식 좌표, 명확한 행정구역 매칭, version-controlled reviewed mapping만 모델 Gold에 허용한다.
- 과거 병합본의 지점번호는 audit-only 후보이며 독립 근거 없이는 승인하지 않는다.
- 애매한 매핑은 nearest city로 조용히 대체하지 않고 quarantine한다.

### [4] Gold join

Owner: `datasets.model_dataset_builder`.

- 발전기간과 실제 존재하는 공식 ASOS 연도만 결합한다.
- 아직 기상자료가 없는 발전량은 Silver에 보존하고 Gold에서만 보류한다.
- 관측소-hour 자체가 없거나 특정 기상변수만 결측인 경우 그 이유를 별도 provenance column으로 남긴다.
- 결측 기상을 0으로 만들어 관측값처럼 취급하지 않는다.

### [5] Quality annotation

Owner: `quality/`.

원본 타깃을 조용히 수정하지 않고 flag와 `quality_train_eligible`로 관리한다.

| 상태 | 처리 |
|---|---|
| 발전량 결측 | 학습 타깃 제외, 원본 보존 |
| non-finite 발전량 | 학습 타깃 제외 |
| 음수 발전량 | 학습 타깃 제외, 0으로 치환 금지 |
| 일 총량이 한 야간 bucket에 적재된 hourly-shaped 자료 | 발전소 전체 hourly 학습 제외, 원본 보존 |
| capacity exceeded | review flag |
| daylight zero | review flag |
| positive flatline | review flag |
| invalid weather | 결측으로 mask + reason flag |
| missing weather | 결측 유지 |
| 임의 target interpolation | 기본 금지 |

capacity exceed, daylight zero, flatline은 공개자료만으로 설비 고장·정비·출력제어와 구분할 수 없으므로 자동 삭제 근거로 사용하지 않는다.

### [6] Leakage-safe features

Owner: `features/`.

- history feature는 실제 과거 시각만 사용한다.
- lag는 최소 24시간 이상 뒤로 이동한다.
- rolling 통계도 미래/현재 target을 포함하지 않도록 shift한다.
- 결측 history 때문에 행을 좋은 구간만 골라 제거하지 않는다.
- 모델별 imputation/scaling 통계는 Train에만 fit한다.

### [7] Training eligibility

Canonical owner: `preprocessing/training_eligibility.py`.

모델 종류와 무관한 **발전소 시계열 구조 적격성**을 판정한다.

필수 규칙:

1. `energy_source=solar`
2. `quality_train_eligible=true`
3. 발전량과 ASOS가 겹치는 시간이 존재
4. Train/Validation/Calibration/Test 네 구간 모두 overlap 존재
5. 각 split 내부에 1시간 초과 시간 gap이 없음
6. Validation overlap rows <= Train rows
7. Calibration overlap rows <= Train rows
8. Test overlap rows <= Train rows
9. 발전소별 개별 split cutoff 사용 금지

통과 상태는 `STRUCTURALLY_ELIGIBLE`, 실패는 `STRUCTURAL_REJECT`이다.

과거 버전관리 코드/PPT/README에서는 별도의 고정 최소 N건 기준이 확인되지 않았으므로 근거 없는 최소 N은 만들지 않는다.

### [8] Training admission

Canonical owner: `preprocessing/training_admission.py`.

- frozen policy에 따라 `STRUCTURALLY_ELIGIBLE -> ADMITTED`.
- 구조 규칙에 실패한 발전소는 `REJECT`.
- Test 성능을 보고 모집단을 다시 고르지 않는다.
- admission은 모델 모집단에만 영향을 주며 서비스 전국 현황에는 영향을 주지 않는다.

### [9] Materialized ADMITTED dataset

`scripts/materialize_admitted_training_data.py`가 frozen population manifest를 읽어 별도 학습 파일을 만든다.

필터는 다음 세 개뿐이다.

- ADMITTED `plant_id`
- `energy_source=solar`
- `quality_train_eligible=true`

원본 Gold와 서비스용 데이터는 변경하지 않는다.

### [10] Model-specific readiness

Owner: `evaluation/forecast_readiness.py`.

여기부터는 **전처리 모집단 정책이 아니라 모델 입력 geometry 검증**이다.

- XGBoost horizon별 exact origin 존재 여부
- CNN lookback 24/168시간 연속 입력창
- 1/24/72시간 horizon 적용 후 실제 usable sample
- Train feature sanity
- 후보 간 Validation/Calibration/Test 공통 평가 key coverage

즉 CNN 168시간 history 때문에 표본이 줄어드는 것은 source preprocessing reject가 아니라 model readiness 결과다.

---

## 3. 시간 분할 최종 정책

메인 실험은 발전소별 percentage split을 사용하지 않는다.

모든 ADMITTED 발전소에 같은 calendar boundary를 적용한다.

| Split | Canonical boundary |
|---|---|
| Train | ~ 2024-06-30 23:00 |
| Validation | purge 후 ~ 2024-09-30 23:00 |
| Calibration | purge 후 ~ 2024-12-31 23:00 |
| Test | purge 후 ~ 2025-12-31 23:00 |
| Boundary purge | 168시간 |

발전소별 데이터 확보 시작시점이 다르기 때문에 실현 비율은 달라질 수 있다. 이것은 분류 규칙이 다른 것이 아니다. **같은 실제 timestamp를 같은 역할로 분류하는 것**이 일관성 기준이다.

짧은 이력 발전소 때문에 그 발전소만 Test를 앞당기거나 60/15/10/15를 따로 계산하지 않는다. 공통 경계를 만족하지 못하면 main cohort에서 제외하거나 short-history/cold-start 분석으로 별도 보고한다.

현재 로컬 Gold 재감사에서 이 경계가 전체 main cohort에 부적합하다고 확인되면, Test 성능을 보기 전에 하나의 새 global boundary 세트를 version up한다. 발전소별 cutoff는 만들지 않는다.

---

## 4. Split별 정보 사용 범위

| 구간 | 허용 목적 |
|---|---|
| Train | 모델 fit, imputation/scaling 통계 fit |
| Validation | feature/model/hyperparameter/early stopping 선택 |
| Calibration | Validation에서 확정된 base model의 Hybrid/gate 조정 및 채택 판단 |
| Test | 모든 선택 고정 후 최종 결과 보고 |

Test 결과는 발전소 admission, split boundary, feature set, model hyperparameter, Hybrid rule을 변경하는 근거로 사용하지 않는다.

---

## 5. 과거 기준과 현재 기준의 관계

과거 구현:

```text
Train       <= 2021
Validation  = 2022~2023
Test        >= 2024
purge       = 24h
```

복원된 발전소 선택 원칙:

- 세 split 모두 존재
- 1시간 초과 gap 없음
- Validation <= Train
- Test <= Train

현재는 Calibration을 추가했지만 핵심 원칙은 유지한다.

```text
Validation <= Train
Calibration <= Train
Test <= Train
```

그리고 과거처럼 모든 발전소에 하나의 공통 시간 경계를 적용한다.

---

## 6. 코드 소유 구조

```text
src/solar_forecast/
├─ collectors/       # 공식 원천 수집 + 공급자 adapter
├─ preprocessing/    # 전처리 orchestration, eligibility, admission, contracts
├─ datasets/         # 표준화 저장, registry, Gold builder, repository
├─ quality/          # 물리/문맥 품질 flag
├─ features/         # leakage-safe input features
├─ evaluation/       # 모델별 readiness와 성능 평가
├─ models/           # XGB/CNN/Hybrid
└─ reporting/        # 서비스/포트폴리오 projection
```

기존 `datasets.preparation_service`, `evaluation.training_eligibility`,
`evaluation.training_admission` 경로는 현재 호출자 보호를 위한 compatibility shim일 뿐이다.
새 코드는 반드시 `solar_forecast.preprocessing` 경계를 사용한다.

---

## 7. 산출물 lineage

```text
official source
  -> standardized generation manifest
  -> plant registry
  -> model_ready_manifest.json
  -> training_data_eligibility.json
  -> training_population.json
  -> admitted_training.csv.gz (+ manifest)
  -> admitted_experiment.json
  -> forecast_readiness.json
  -> model benchmark artifacts
```

각 단계에서 이전 단계 파일/hash/contract를 참조하며, source dataset을 덮어쓰지 않는다.

---

## 8. 포트폴리오 설명

> 공식 원본과 전국 설비 현황은 보존하면서 AI 모델용 모집단만 별도의 적격성 정책으로 선별했습니다. 공급기관별 스키마와 단위를 plant-hour 공통 계약으로 표준화하고, 발전소 식별·ASOS 매핑·시간 해상도·품질 플래그를 단계적으로 검증했습니다. 학습 모집단에는 모든 발전소에 동일한 calendar boundary를 적용해 같은 실제 기간을 Train/Validation/Calibration/Test로 분류했고, 데이터 이력이 짧은 발전소만을 위해 평가 기간을 개별 조정하지 않았습니다. 모델별 lookback·horizon으로 인한 실제 표본 감소는 source 전처리와 분리해 readiness 단계에서 검증함으로써 데이터 품질 판단과 모델 구조 영향을 구분했습니다.

이 문장을 README/면접 요약에 사용하되, 실제 발전소 수·행 수·성능 값은 최신 실행 manifest를 기준으로 갱신한다.
