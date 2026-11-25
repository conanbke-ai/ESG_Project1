# Solar Forecast Track B 데이터·모델 검증 — 2026-11-25 KST

**판정** — 기존 원본과 평가 산출물은 확인된 범위에서 재사용한다. 새 원본의 Gold 편입, 실발행 시각 기준의 운영 예측 검증, 전면 재학습은 보류한다. Windows 선행 검수와 이 작업의 Cloud 재검증은 각 실행의 파일과 SHA로 식별한다.

## P0–P1. 모집단과 원천 최신 시점

- EPSIS 발전기 등록 원본의 기준일은 **2026-08-05**, 188,594행·33,059.516180 MW다. 공개 발전기 변경 현황에도 같은 날짜가 표시됐다. 고유키 없이 모든 필드가 일치하는 초과 557행은 기존 모집단에 포함한다. 지역 충돌 51행 중 47행은 기존 근거로 표준화했고, 미확정 4행·4.394255 MW도 모집단에 포함했다. 원본 SHA-256은 `config/national_solar_inventory.json`의 등록값과 일치했다. [EPSIS 발전기 현황](https://epsis.kpx.or.kr/epsisnew/selectEkpoBcrGrid.do?menuId=020400) · [등록 원본 설정](../config/national_solar_inventory.json)
- KOSPO 부산철도 2는 공식 명칭과 주소로 1·4호와 구분된다. EPSIS에는 부산진구 `부산 철도 2호 태양광`, `호기=1`, **0.80920 MW**가 별도로 등록되어 있다. 이는 해당 EPSIS 스냅샷의 설비 등록값이다. 발전량 CSV의 시설 식별자·계량점과 연결되는 증거, 과거 적용 용량, 승인된 ASOS 매핑이 확보되면 편입 판단을 다시 한다. [KOSPO 식별 감사](KOSPO_IDENTITY_AUDIT.md) · [공식 발전량 데이터셋 15156688](https://www.data.go.kr/data/15156688/fileData.do)

Cloud에서 추적 원본 88개를 기존 정규화기로 읽었다. 시간은 KST이며, 아래 건수의 모집단은 원본별 정규화 파티션이다.

| 발전사 | 정규화 행 | Cloud 최신 발전 시각 | Windows 선행 수집 범위 |
|---|---:|---|---|
| KOEN | 737,496 | 2025-09-30 23:00 | 2026년 10월 |
| KOSPO | 3,093,580 | 2026-02-28 23:00 | 2026년 8월 |
| EWP | 182,400 | 2025-06-30 23:00 | 2026년 6월 |
| IWP | 223,089 | 2023-12-31 23:00 | 추가 날짜 확인 근거 없음 |
| 전체 | 4,236,565 | | |

KOEN 공식 조회 화면에서 2026-10-31 일자 자료를 확인했다. Windows 선행 수집에서 2024년 5·9·10월은 부분 응답이고, 2026년 11월 빈 응답의 누락 여부는 미확정이다. KOSPO 포털의 수정일 2026-10-19와 발전 데이터 종료 시점은 별개다. Cloud에서 공식 원천과 Open-Meteo 요청은 프록시 `CONNECT 403`이었다. 새 원본은 0건이며 Windows 수집본을 재요청하지 않았다. [KOEN 월별 조회](https://www.koenergy.kr/kosep/gv/nf/dt/nfdt21/main.do?menuCd=FN0912020217) · [Windows 감사가 인용된 PR #2](https://github.com/conanbke-ai/ESG_Project1/pull/2)

Cloud ASOS 2025년 원본은 **97지점·714,692행**, `2025-11-03 23:00 KST`까지다. 지점·시간 중복과 비정시 행은 0건이며, 내부 누락은 2구간·4지점시간이다. 기온 결측 1,131건, 일조 313,433건, 일사 449,767건을 관측했다. 지점 이력 유효기간 밖은 2,736행이다. 원본 SHA-256은 `734ee62139ead28320de66820e3ad3d9191dfa722a5c8c48e01bbb86fee74373`이다. Windows 2026년 준비본은 ASOS 7지점과 24/48/72h 예보 각 9개소를 기록하며, 확인된 가장 긴 범위는 6월 말까지다. Cloud에는 해당 준비본과 학습용 MSM/GSM 파일이 없어 각 발전소·horizon의 실커버리지를 재계산할 수 없었다.

## P2. 단위, 시간 및 예보 품질

- 정규화 발전량의 결측·비유한·음수와 파일 내 발전소·시간 중복은 각각 0건이었다. 용량 결측은 KOEN 164,280행과 KOSPO 416,184행이었다. `generation_mwh > capacity_mw × 1h`는 각각 48,247행과 1,557행이며 원천·용량 적용시점·장비 범위를 확인할 후보다. 익산 다송리 14,064시간은 Wh를 MWh로 환산했고, 하루 합계가 kWh에서 Wh로 바뀐 구간을 각각 대사했다. 불일치는 0건이다. [정규화 정책](PLANT_DATA_PREPROCESSING_POLICY.md)
- Windows 선행 검수에서 EWP 15131649의 912일 원행·21,888개 시간 슬롯 중 21,886개가 유효했다. `2026-04-05` 14·15시의 `-` 두 값은 결측으로 보존했고, `2,862,844,588 Wh → 2,862.844588 MWh`가 역산 일치했다. 2025년 목표 8,760시간에 대한 이력 키 검사는 통과했다. 해당 계량점·ASOS 연결은 HOLD다. 증거는 Windows `artifacts/preprocessing_read_audits/20261113_resumable_raw_collection_v2/EWP15131649_END_CLOCK_FULL_REVIEW_20261124.json`에 있다.
- 관측 기반 코드에서 목표 `timestamp`는 `valid_at`에 해당하고 이력은 `forecast_origin = valid_at − horizon`까지 사용한다. CNN 시퀀스는 연속된 시간 키를 사용한다. 원천 발전량의 `0–1` 같은 시간 구간은 기존 이력에서 시작 시각으로 표기되므로, 운영 기준 `observed_at`과 공개·수신 시각은 구간 종료 및 지연을 따로 확인해야 한다. 보존 산출물은 KST 현지 시각을 오프셋 없이 표현한다. [이력 결합](../src/solar_forecast/evaluation/forecast_samples.py)
- 24h MSM은 기상·일사 기본 피처와 DNI/DHI 추가 프로필을 사용한다. 72h GSM은 일사 없이 5개 기상 피처를 사용하며, 공급자의 6시간 격자 값이 제공 시계열에서 시간 단위로 보간된다. 1h에는 미래기상 입력이 없다. Previous Runs의 고정 리드는 여러 모델 런을 이어 붙인 결과다. 파생 `forecast_origin`으로 고정 리드 정렬을 검사할 수 있으며, 운영 판정에 필요한 실제 `issued_at` 및 `available_at ≤ forecast_origin` 증거는 별도로 필요하다. [실험 설정](../config/experiments/optimized.json) · [JMA 기상 변수](https://open-meteo.com/en/docs/jma-api) · [Previous Runs](https://open-meteo.com/en/docs/previous-runs-api) · [모델 공개 시각](https://open-meteo.com/en/docs/model-updates)

### 사전 점검기 변경

`tools/validate_future_weather.py`의 `validate`는 설정에 모델이 지정된 경우 manifest와 전체 선택 horizon의 행 모델을 대조한다. 정시 외 행과 빈 발전소 키를 검출하고, `usable_from`과 아카이브 타임스탬프의 타임존 표현이 달라 비교할 수 없는 경우 오류 원인을 표시한다. `_feature_stats`는 무한대를 범위 위반에 포함한다. 새 `archive_hourly_keys`는 아카이브에 나타나는 사용 기간과 발전소에서 누락된 시간 키를 집계한다. 기존 피처 유효 행 비율 기준은 해당 행이 있는 코호트에 적용된다. 전체 미수집 발전소와 양 끝 구간은 기대 모집단과 요청 기간이 있어야 검증할 수 있다. [예보 사전 점검기](../tools/validate_future_weather.py) · [회귀 테스트](../tests/test_future_weather_preflight.py)

이 변경은 읽기 전용 CLI와 보고서 필드에 한정된다. 공개 기반 `read_future_weather`의 소수 horizon 정수 절삭은 추가 결함으로 재현됐으며, Windows에서 달라진 해당 모듈의 검증과 함께 처리해야 한다. 개별 공급자 발행·수신 기록 없이 과거 운영 가용성은 판정할 수 없다.

## P3. Gold, 분할과 모델 입력

Windows에서 `44714db2`로 실행한 실제 Gold 대조는 **718,531행·96열·27파티션**, 중복·누락·값 불일치 0건을 기록했다. 태양광 22개소 중 21개소는 품질 적격이고 동결 모델 비교는 19개소다. 이 작업에 마운트된 Cloud에는 Gold와 해당 MSM/GSM manifest가 없다. 기존 검증 CLI에 실제 설정을 넣은 실행은 24h MSM manifest 미존재로 종료했다. [Windows Gold 및 체크포인트 증거](https://github.com/conanbke-ai/ESG_Project1/pull/2)

- 기존 전역 경계는 Train `2024-06-30 23:00`, Validation `2024-09-30 23:00`, Calibration `2024-12-31 23:00`, Test `2025-12-31 23:00 KST`이며 purge 168시간이다. 9월의 보존 실행 `20260922_132607_286733`에서 공통 검증 행은 horizon별 38,760개, 공통 Calibration은 38,664·38,497·38,593개다. 공통 Test는 각각 19개소·105,336행으로, 세 horizon의 발전소·목표 시각·발전량 실측 키가 일치했다. 발전소별 행 수는 1,248–8,592개다. [전역 분할 설정](../config/experiments/optimized.json)
- 9월 실행 manifest의 Gold fingerprint는 `9386c841d06ca1427be54e77d7277c9c15b72820d762e3e72a2fa3f380093211`이다. 보존 XGB 3개에서 부스터와 설정의 입력 피처 순서가 일치했고, 결측은 native NaN 처리를 사용한다. CNN 3개의 가중치를 엄격 로드했고 전처리 v2, Train 전용 통계, 마스크 추가 이후 52/52/36개 피처와 24/168/24시간 입력창을 확인했다. 이 가중치의 동일 Test 추론은 정확한 과거 Gold fingerprint 입력을 구할 때 재개한다. 이 보존 CNN의 직접 MWh 타깃과 현재 최적화 설정의 `capacity_factor` 타깃을 상호 교체할 수 없다. [모델 포장·리플레이 도구](../tools/verify_benchmark_model_artifacts.py)
- 현재 `audit_forecast_readiness.py`의 코호트는 관측 기상 기준이다. 예보의 `usable_from`, 아카이브 커버리지와 CNN 용량 필터를 적용한 실학습 코호트 입증에는 별도의 조인 결과가 필요하다. 기존 체크포인트 추론 도구의 경로는 관측 전용이므로 예보 포함 체크포인트 재검증은 해당 인터페이스에서 추가 조율이 필요하다. [코호트 감사](../src/solar_forecast/evaluation/forecast_readiness.py)

## P4. 저장 결과 재평가와 재사용 결정

2025년 Test 저장 예측을 다시 집계했다. 9월 실행 `20260922_132607_286733`의 모델은 관측 이력만 사용했다. Bias는 `예측 − 실측` 평균이며 MAE·RMSE·Bias는 MWh다. [9월 실행 manifest](https://github.com/conanbke-ai/ESG_Project1/blob/main/artifacts/benchmarks/20260922_132607_286733/manifest.json)

| Horizon | 모델 | MAE | RMSE | R² | Bias |
|---|---|---:|---:|---:|---:|
| 1h | XGBoost | .219169 | 1.454474 | .904765 | −.007903 |
| 1h | CNN-BiLSTM | .253775 | 1.548288 | .892083 | +.008425 |
| 1h | Hybrid | .213589 | 1.382959 | .913900 | −.000940 |
| 1h | Persistence | .278613 | 1.578124 | .887884 | ~0 |
| 24h | XGBoost | .358610 | 2.149745 | .791954 | −.042626 |
| 24h | CNN-BiLSTM | .364622 | 2.030738 | .814351 | +.022531 |
| 24h | Hybrid | .368285 | 2.129229 | .795906 | −.032059 |
| 24h | Persistence | .347421 | 2.468376 | .725712 | −.000765 |
| 72h | XGBoost | .380848 | 2.198695 | .782372 | −.071658 |
| 72h | CNN-BiLSTM | .591443 | 2.578504 | .700691 | +.215585 |
| 72h | Hybrid | .462809 | 2.192110 | .783674 | +.091827 |
| 72h | Persistence | .382677 | 2.656715 | .682258 | −.003463 |

9월의 두 완료 실행에서 각 3개 horizon의 저장 점수와 키·SHA 검증이 통과했다. 최신 실행의 발전소·월·주야간·발전소×월을 포함한 지표 2,232행을 `archived_benchmark_metrics.csv`에 저장했다. 72h CNN은 2025년 12월 전체에서 MAE 3.714797 MWh, 제주 용수리 744행에서 9.681249 MWh였다. 결과에는 저장 음수 예측이 포함된다. 9월 실행의 XGB 15회·CNN 5회씩 Validation Optuna trial과 설정을 보존했다. 두 보존 실행의 2025년 Test는 기존 개발 노출 구간이다.

Windows 동결 평가 `formal_2025_comparison_gpu_io_recovery_20261101_v1`의 저장 로그에서는 아래 결과를 확인했다. 각 모델 셀은 `MAE / RMSE / R² / Bias` 순서다. 이 환경에 해당 평가 원본과 가중치가 없어 아래 값은 Windows 검수 기록에서 인용했다. [동결 평가 기준](https://github.com/conanbke-ai/ESG_Project1/issues/5#issuecomment-6997796967)

| Horizon | 실제 사용 기상 | XGBoost | CNN-BiLSTM | Persistence MAE |
|---|---|---|---|---:|
| 1h | 관측 이력 | .2140 / 1.4775 / .9069 / −.0084 | .2068 / 1.3597 / .9211 / +.0099 | .2922 |
| 24h | JMA MSM | .2478 / 1.5775 / .8938 / −.0445 | .2442 / 1.5004 / .9040 / −.0559 | .3623 |
| 72h | JMA GSM | .2987 / 1.8050 / .8610 / −.0581 | .2856 / 1.7475 / .8697 / +.0068 | .3985 |

동결 24h Hybrid의 MAE는 .2326이었다. Windows 검수에서 당시 프로필 기준 checkpoint fingerprint 6/6, 현재 기본 설정 기준 0/6이었다. 당시 XGB Optuna 기록은 25/15/15회다. 9월·11월 간 점수 차이에는 코호트·타깃·설정 차이가 개입할 수 있다. 미래기상 순효과를 측정하려면 같은 Gold fingerprint, forecast-available 키, 모델·전처리·분할 및 튜닝 범위에서 날씨 입력만 바꾼 비교가 필요하다. 기존 2025 Test로 후보를 다시 선택하지 않는다.

새 학습을 정당화할 변경 입력은 시설·계량점과 기상 가용 시각이 입증된 2026년의 증분 데이터, 또는 동일 코호트에서 정한 예보 포함·제외 실험 설정이다. 이번 작업에서 확정된 학습 입력이나 설정 변경은 없었고 새 학습·Optuna·모델 승격을 실행하지 않았다. 필요한 경우 기존 Windows GPU 런타임에서 동결 모델 재검증을 선행한다.

## 실행 및 인계

- 작업 위치는 `/workspace/ESG_Project1-track-b`의 독립 Cloud worktree다. 원격 `main`은 `3dbc598141e1c8fbae4290093aabb67d7ccddf70`, 감사 시작 PR #2는 `f0a5c914e2f9af05457857c7728952f8e112e0ba`, 이 소스 변경의 기준은 병행 세션이 게시한 `44714db22e79ead6bdadaf8ad2522dfa7095a4c0`이다. Windows 최신 작업 트리 `140398fff88dc30d1d38d301348b14f15ad11d59`는 Cloud Git 객체에 없다. [공통 정책 §26–31 `f34e980`](https://github.com/conanbke-ai/Tori_Common_Project/blob/f34e980c269ffdb121ad49a0d22e9c27ce45ad0b/policies/DEVELOPMENT_TEAM_OPERATING_POLICY.md#26-tori-development-framework-v20--서비스화-및-증거-기반-개발-2026-11-24)
- 파일 범위는 `tools/validate_future_weather.py`, `tests/test_future_weather_preflight.py`, 자동 생성 `docs/CODE_INDEX.md` 및 이 보고서다. 로컬 파일의 차이는 `git diff 44714db2 HEAD`에서 확인할 수 있다. 게시·통합·Windows 동기화 상태는 LOCAL_ONLY다.
- Cloud Python 3.11에서 `python -m pytest -q -p no:cacheprovider`는 502개 및 하위 검사 136개, `python tools/check_structure.py`와 `git diff --check`는 통과했다. 테스트의 아카이브는 검증 오류를 재현하기 위한 작은 fixture다. 실제 설정을 사용하는 예보 사전 점검은 MSM manifest 부재를 확인했다.
- `artifacts/verification/track_b_20261125/`에 저장 점수 CSV, 모델·피처·파일 해시 상세 JSON, 점수 재계산 소스를 보존했다. 기존 예측·가중치·데이터를 Git에 다시 추가할 필요가 없다.
- P0 재사용과 P2 범위 검증은 완료, P1은 원천/환경 및 admission 기준이 남아 부분 완료, P3은 저장 코호트·입력 비교까지 완료, P4는 저장 점수 평가까지 완료했다. 다음 실행의 필요 조건은 Windows에서 보존된 원본·manifest·예보 아카이브에 대한 읽기 접근이다. 운영 판정에는 발행·수신 시각 증거가 추가로 필요하다. 비용·사용량은 미계측이다.

## E-B 후속 검증 — 2026-11-25 KST

공통 정책 `df0331220f6eae2b802bed612ec306ac20893c1a`의 §26–32를 적용했다. E-B Codex 세션 `6b06670f-12d0-83ee-a924-c151d9ed60be`의 작업 위치는 `/workspace/ESG_Project1-eb`다. PR #2 `13c4c85ffe1c0421b4bb9b4bd32453863d08d76e`와 기존 `d825f07171ecdecf83a831bf41a88ca19d5cbe1e`의 공통 조상은 `44714db2`이며, 두 변경의 겹침은 생성 색인 `docs/CODE_INDEX.md` 한 파일이다. PR #2 HEAD 위에 날씨 검사기와 기존 회귀·보고서를 반영하고 색인을 재생성했다. 검토 브랜치는 E-B 검증 경로를 담당한다. E-A의 정식 브랜치와 Phase 2 작성 경로는 각 기존 세션이 관리한다. [공통 정책](https://github.com/conanbke-ai/Tori_Common_Project/blob/df0331220f6eae2b802bed612ec306ac20893c1a/policies/DEVELOPMENT_TEAM_OPERATING_POLICY.md) · [PR #2](https://github.com/conanbke-ai/ESG_Project1/pull/2)

### 원천 완전성·계량점·기간별 용량

Phase 2의 실측 감사 `009edb07`을 재사용한다. KOEN 2024년 198,432개 키의 파일 간 동일값·충돌은 0건이다. 개별 호기 공백은 2024년 5월 여수 1호 192시간, 9월 예천 1호 24시간, 10월 여수 1호 456시간이다. KOSPO 누적 스냅샷은 신풍리·용수리에서 2024년 17,568시간과 2025년 2,832시간이 동일하게 겹쳤다. 신풍리·용수리 2026년 각 1,416시간은 중복·충돌이 없었고, 삼척소내 92시간 누락·비수치 및 신인천해수구취수구 호기 공란 8,760시간은 원천 판정에 남아 있다. [Phase 2 원천 완전성](https://github.com/conanbke-ai/ESG_Project1/blob/009edb07d2ed603cefafe893c08f71bb1c0b2763/docs/ESG_DATA_ADMISSION_AND_MODEL_REUSE_PHASE2.md#p0-원천-완전성) · [원천 증거 JSON](https://github.com/conanbke-ai/ESG_Project1/blob/009edb07d2ed603cefafe893c08f71bb1c0b2763/docs/evidence/data_admission_phase2/source_completeness.json)

추가 도구는 기존 정규화 원본 또는 Silver에서 공급자 발전소·호기와 기간별 MW 보고값을 확인한다. 실제 계량점과 설비용량 유효기간은 로컬 출처 해시가 붙은 `[시작, 종료)` 검토 증거로 판단한다. 보고값 누락, 계량점 또는 용량 효력 기간의 공백·충돌, 같은 물리 계량점에 대한 중복 바인딩을 드러낸다. 2026년 Gold 편입에는 계량점, 시간 버킷의 의미, 발전량·기상 실제 가용성까지 필요하다. [발전량 검증 도구](../tools/audit_generation_readiness.py)

Phase 2 시간축 검사를 IWP 보관 원본 두 개에 적용했다. 9개 소스 시계열·223,089시간, `2017-01-01`~`2023-12-31 KST`, 파일 간 동일값·충돌 0건, 내부 누락 255 시계열시간이다. 별도 보존 원본 10개에서 공급자 시계열 19개·581,649행의 용량과 계량 근거를 점검했다. 다송리 14,064시간은 원천에 0.893 MW가 기록됐다. 신풍리·용수리 2026년 원천 각각 1,416시간에는 용량 값이 없어, 별도 설비 사양 0.99072·0.88893 MW의 실제 적용 기간은 추가 증거가 필요하다. 물리 계량점을 확인한 시간은 0이다. Cloud 조사 출력은 `artifacts/verification/track_b_eb_20261125/`의 `iwest_retained_source_completeness.json`(SHA-256 `83f52d00fe0a4af985b375aa79046d4a43481eea3dd5e391b1baf25accd30258`), `generation_asset_evidence.json`(SHA-256 `f5e4b3662adad5b3d96642a4393151092258db6e6b6e61f72b18ae71435bf180`)에 보존했다.

`python tools/audit_generation_readiness.py --partition <보존 Silver> --review <검토 JSON> --output <새 결과 JSON>`으로 기간 검토를 재현한다. `--archive-file "kospo=<보존 원본>"`도 반복해서 지정할 수 있다. `solar-generation-asset-evidence.v1` 검토 JSON은 출처 ID별 `{path, sha256}`과 `meter_bindings`·`capacity_periods`를 받는다. 기간 항목은 `source_series_id`, `start_kst`, `end_exclusive_kst`, `evidence_id` 및 각각 `meter_point_id` 또는 `capacity_mw`를 포함한다.

### 예보 시각 증거

| 필드 | 검증 의미 |
|---|---|
| `forecast_origin` | 예측 기준 시각이다. 고정 리드 아카이브에서는 `valid_at − horizon_hours`이며 행 키로 사용한다. |
| `issued_at` | 해당 예보의 기록된 발행 시각이다. Phase 2의 `forecast_issued_at`에 대응한다. |
| `available_at` | 해당 예보가 소비자에게 실제 가용해진 기록 시각이다. Phase 2의 `forecast_received_at`은 같은 소비자의 가용을 증명하는 수신 기록일 때 변환할 수 있다. |
| `valid_at` | 예보가 적용되는 목표 시각이다. 기존 아카이브의 `timestamp`에 대응한다. |

`--timing-evidence` CSV에는 `plant_id,horizon_hours,forecast_origin,issued_at,available_at,valid_at,archive_sha256,evidence_ref`를 사용한다. 네 시각은 ISO 8601에 UTC 오프셋을 포함한다. 검사기는 사용 구간의 발전소·리드·시각 키와 파일 SHA-256을 대조하고 `issued_at ≤ available_at ≤ forecast_origin < valid_at`을 검증한다. 기록이 없으면 `UNVERIFIED`, 일치한 일부 행은 `PARTIAL`, 전체 사용 행의 기록 정합성은 `VERIFIED`다. 오프셋·고유 키·해시·시각 순서를 검사하지만, `evidence_ref` 원문이나 서명의 인증은 검토자가 확보한 자료에 의존한다. `fixed_lead_verified`는 기존 로더의 시각 산술 검사다. 가용 기록과 보간한 고정 리드 산술은 별개의 증거다. [예보 사전 점검기](../tools/validate_future_weather.py) · [공급자 공개 시각](https://open-meteo.com/en/docs/model-updates)

### Windows 읽기 전용 인계

기존 담당 세션은 `[E-A] 태양광 | 아키텍처·통합` `01a03de5-b1ea-78b2-bd9c-7c2c684a4de2`다. E-A의 Windows 선행 검증에서 24h MSM 신풍리 CPU 결과 `0.365685455851 MWh`는 저장 CPU 결과와 정확히 일치했다. 저장 CUDA 결과와 차이는 약 `0.000016553 MWh`였고 당시 허용 범위를 벗어났다. 24h CPU, Gold 전체 대조, 격리 체크포인트 로드는 이 결과를 재사용한다. 직접 Codex 세션 전송 capability는 `DISPATCH_PENDING`이며, 게시된 검토 브랜치가 인계 자료다.

첫 명령은 E-B 코드 checkout에서 실행한다. 기존 Windows 원본의 **MSM/GSM 압축본 두 개와 각각의 manifest**, 기존 실험 설정, Python 환경이 최소 입력이다. 조사 출력은 새 임시 경로에 생성한다.

```powershell
$Root = 'C:\ESG_Project1'
$Py = Join-Path $Root '.venv-solar-gpu\Scripts\python.exe'
$Code = (Get-Location).Path # E-B checkout
$Out = Join-Path ([System.IO.Path]::GetTempPath()) ('esg-eb-' + [Guid]::NewGuid().ToString())
New-Item -ItemType Directory $Out | Out-Null
$Config = Get-Content (Join-Path $Code 'config\experiments\optimized.json') -Raw | ConvertFrom-Json -AsHashtable
foreach ($Lead in @('24', '72')) {
    $Relative = $Config.future_weather.by_horizon[$Lead].source
    $Source = if ([System.IO.Path]::IsPathRooted($Relative)) { $Relative } else { Join-Path $Root $Relative }
    if (!(Test-Path -LiteralPath $Source) -or !(Test-Path -LiteralPath ($Source -replace '\.csv\.gz$', '.manifest.json'))) {
        throw "Missing $Lead-hour archive or manifest at $Source"
    }
    $Config.future_weather.by_horizon[$Lead].source = $Source
}
$Config | ConvertTo-Json -Depth 50 | Set-Content -Encoding utf8NoBOM (Join-Path $Out 'optimized.json')
& $Py -X utf8 -B (Join-Path $Code 'tools\validate_future_weather.py') `
    --config (Join-Path $Out 'optimized.json') --minimum-coverage 0.98 `
    --output (Join-Path $Out 'msm_gsm_preflight.json')
if ($LASTEXITCODE -ne 0) { throw 'MSM/GSM preflight failed' }
```

발행·가용 로그가 검토된 경우 `--timing-evidence <CSV>`를 추가한다. E-A가 확인한 Windows 런타임의 72h witness는 기존 `acceptance_report.json`, `targets.csv`, `72h/source_manifest.json`과 `72h/prediction.json`을 사용한다. 비교에 필요한 최소 입력은 소스 manifest가 해시로 고정한 72h GSM·관측 이력 표본, 동결 레지스트리가 참조한 가중치·설정·전처리, 같은 키와 장치의 예상 예측값이다. 보존 PASS가 동일 해시·장치의 요청을 충족하면 기존 결과를 회수한다. 새 증거가 필요한 경우 Windows 정식 런타임에서 해당 horizon만 호출한다.

```powershell
Set-Location -LiteralPath $Root
$Witness = Join-Path $Root 'artifacts\runtime\verification\20261023_runtime_v2'
$Target = Join-Path $Witness 'targets.csv'
$Manifest = Join-Path $Witness '72h\source_manifest.json'
$SavedPath = Join-Path $Witness '72h\prediction.json'
foreach ($Path in @($Target, $Manifest, $SavedPath)) {
    if (!(Test-Path -LiteralPath $Path)) { throw "Missing frozen witness: $Path" }
}
$Saved = Get-Content -LiteralPath $SavedPath -Raw | ConvertFrom-Json
$Devices = @($Saved.rows.execution_device | Sort-Object -Unique)
if ($Devices.Count -ne 1 -or $Devices[0] -notin @('cpu', 'cuda')) { throw 'Missing original execution device' }
$Prediction = Join-Path $Out 'prediction_72h.json'
& $Py -X utf8 -B app.py production-predict --horizon 72 --request-id 'eb-frozen-72h' `
    --targets $Target --source-manifest $Manifest --ledger (Join-Path $Out 'replay.sqlite3') `
    --device "$($Devices[0])" --output $Prediction
if ($LASTEXITCODE -ne 0) { throw '72h frozen inference failed' }
$Actual = Get-Content -LiteralPath $Prediction -Raw | ConvertFrom-Json
if ($Actual.rows.Count -ne $Saved.rows.Count) { throw 'Frozen witness row counts differ' }
foreach ($Row in $Actual.rows) {
    $Previous = @($Saved.rows | Where-Object {
        $_.plant_id -eq $Row.plant_id -and $_.timestamp -eq $Row.timestamp -and
        $_.forecast_origin -eq $Row.forecast_origin -and $_.horizon_hours -eq $Row.horizon_hours
    })
    if ($Previous.Count -ne 1 -or $Previous[0].policy_sha256 -ne $Row.policy_sha256 -or
        $Previous[0].feature_source_sha256 -ne $Row.feature_source_sha256) {
        throw 'Frozen witness keys, policy or features differ'
    }
    $ErrorMwh = [Math]::Abs($Row.prediction_mwh - $Previous[0].prediction_mwh)
    if ($ErrorMwh -gt (1e-6 + 1e-5 * [Math]::Abs($Previous[0].prediction_mwh))) {
        throw "Frozen 72h parity error: $ErrorMwh MWh"
    }
}
```

기본 동결 레지스트리의 해시가 보존 기록과 다르면 `--registry <일치하는 기존 동결 registry>`를 지정한다. 회수할 최소 결과는 preflight JSON, 72h 실행 여부와 누락 경로 또는 행 수·장치·최대 오차·예보 프로필/입력/가중치 해시다. 72h 신규 재실행은 Cloud에서 미실행이다.

### 동일 조건 예보 비교

`tools/compare_future_weather_predictions.py --weather-on <후보 디렉터리> --weather-off <후보 디렉터리> --output <새 결과 JSON>`은 완료된 두 실행에서 동일 모델·데이터 fingerprint·전역 분할과 purge·설정/튜닝 조건을 확인한다. Validation/Calibration/Test의 발전소·목표·기준 시각·horizon 키, 실측 MWh와 Persistence를 대조한 뒤 MAE·RMSE·R²·Bias를 전체·발전소·월별로 계산한다. 모델별 후보 manifest와 해석된 설정, 세 분할의 기존 예측 CSV, 상위 benchmark manifest가 최소 입력이다. 두 실험의 실제 공통 Train 키와 예보 발행·수신 가용성은 별도로 검토한다. 현재 확인 가능한 9월·11월 저장 결과의 설정 및 입력이 달라 실제 weather-on/off 효과 검증은 `NOT_RUN`이다. [동일 조건 예보 비교 도구](../tools/compare_future_weather_predictions.py)
