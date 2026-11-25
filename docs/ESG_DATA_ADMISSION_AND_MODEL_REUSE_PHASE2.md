# ESG 데이터 편입과 모델 재사용 Phase 2

2026-11-25 KST 기준. Cloud에서 확인한 입력은 Git에 보존된 발전량 원본, 기상 관측자료와 9월 벤치마크다. Windows 정식 프로젝트 `C:\ESG_Project1`은 이 실행 환경에 연결되어 있지 않다. 정식 Gold의 기준은 718,531행·96열·27개 파티션이다.

| 작업 | 완료 상태 | 근거 |
| --- | --- | --- |
| 원천 수집 | 보존 파일 감사 완료, 증분 수집 0건 | 86개 파일의 시간축·중복·정정 후보를 검증했다. KOEN의 최신 조회와 Windows 수집 인덱스가 필요하다. |
| 전처리·편입 | 발전소별 감사 완료, 2026 편입 0행 | KOSPO 2,832행을 임시 시간축에 정규화하고 검토했다. 버전 `gold-2026-phase2-v1`의 편입 manifest를 남겼다. |
| 모델 재개 | CPU 검증 통과, 정식 Windows 재개 미검증 | 9월 모델 6개와 과거 식별자, 별도 CPU fixture의 checkpoint·Optuna 재개를 검증했다. |
| 새 학습 | 미실행 | 정식 Gold와 모델 입력의 적격성 확인이 남아 있다. |
| 독립 평가 | 비교 계약 준비, 평가 미실행 | 2025년 개발 비교를 재검증했다. 2026년 적격·미노출 후보 0행이다. |

## P0 원천 완전성

| 기관 | 보존 원본 | 마지막 관측 KST | 추가 확인 |
| --- | ---: | --- | --- |
| KOEN | 45개 | 2025-09-30 23:00 | 공식 태양광 조회 화면에서 2026-10-31 표본을 확인했다. 공개 파일 `15167614`의 명시 범위는 2022년~2026년 8월이다. 조회 가능한 연속 기간과 11월 결과는 확인이 필요하다. |
| EWP | 3개 | 2025-06-30 23:00 | 전국 학습자료 게시물 `43582`와 동해 발전소의 2026년 6월 공개파일은 별개 자료다. 최신 첨부 원본의 관측 시작·종료를 확인해야 한다. |
| KOSPO | 38개 | 2026-02-28 23:00 | 신풍리·용수리 각 1,416시간이 2026-01-01~02-28에 연속 관측됐다. 별도 부산철도 태양광 2호기 파일 `15156688`의 최신 게시일은 발전량 관측 범위를 뜻하지 않는다. |

KOEN 2024년 5·9·10월에는 기관 전체 기준으로 모든 달력 날짜가 있었다. 포함된 원본 행에는 24시간 값이 있었다. 아래 시간은 해당 월에서 관측된 개별 호기의 달력 시간 대비 공백이다. 여수 1호기는 6~9월 원본에서 관측되지 않아 그 기간의 운전·공개 여부가 추가로 필요하다. [원천 검증 JSON](evidence/data_admission_phase2/source_completeness.json)

| 월 | 발전소·호기 | 보존 원본에 없는 날짜 | 시간 |
| --- | --- | --- | ---: |
| 5월 | 여수태양광 1호기 | 3, 10, 13~16, 18, 20일 | 192 |
| 9월 | 예천태양광 1호기 | 5일 | 24 |
| 10월 | 여수태양광 1호기 | 1~9, 11~14, 16~21일 | 456 |

KOEN 2024년 관측 키는 198,432개였고 동일 값 반복·충돌은 0건이었다. EWP 보존 발전소 파일에서 값 충돌은 0건이었다. KOSPO 누적 스냅샷은 신풍리·용수리에서 2024년 17,568시간과 2025년 2,832시간이 동일 값으로 겹쳤다. 두 발전소의 2026년 보존 시간은 중복·충돌 0건이다. KOSPO 레거시 파일에는 삼척소내의 누락·비수치 92시간과 신인천해수구취수구의 호기 식별자 공란 8,760시간이 별도로 있다. 명시적으로 확인된 수정 공표가 없으면 충돌값을 자동 선택하지 않는다. [원천 검증 JSON](evidence/data_admission_phase2/source_completeness.json)

기존 KOEN 수집 코드는 요청 종료일을 월말로 확대했다. 11월 25일 요청은 30일까지 전송됐으며, 브라우저 예외·다운로드 시간초과·파일 부재에 따라 같은 빈 파일 목록을 반환할 수 있었다. 이제 요청 일자를 달력 월 단위로 정확히 나누고, 수집 manifest에 범위별 캐시·수신·빈 CSV·브라우저 실패 등의 상태를 기록한다. 일부 창만 완료되면 `partial`로 표시한다. 11월 실제 원인은 Cloud에서 공식 포털 날짜 대조가 터널 오류로 막혀 미확정이다. 공급자에 원본이 존재하는 것으로 확인된 누락 창만 증분 수집 대상으로 삼는다. [KOEN 공식 조회](https://www.koenergy.kr/kosep/gv/nf/dt/nfdt21/main.do?menuCd=FN0912020217) · [KOEN 공개 파일](https://www.data.go.kr/data/15167614/fileData.do) · [EWP 학습자료 게시물](https://www.ewp.co.kr/kor/subpage/content.html?pc=12V5W1N4BPZD35P8P1W818Z0TS07PXA&state=view&idx=43582) · [KOSPO 부산철도 2호기](https://www.data.go.kr/data/15156688/fileData.do)

## P1 2026년 Gold 편입

| 발전소 | 설비용량 MW | 최대 시간 발전량 MWh | 검토 ASOS | 임시 관측 시간 |
| --- | ---: | ---: | --- | ---: |
| KOSPO 신풍리 1호기 | 0.99072 | 0.8676 | 성산 188 | 1,416/1,416 |
| KOSPO 용수리 1호기 | 0.88893 | 0.711587 | 고산 185 | 1,416/1,416 |

두 발전소는 발전량 단위·설비용량·관측 시간축·물리 품질 검사를 통과했다. 발전기 식별자와 원본 계량 분류를 기록했다. 실제 계량점 ID, 시간 버킷의 공급자 시작·종료 정의, 발전량과 ASOS의 실제 가용 시각, 2026 ASOS 연결, 1·24·72시간 예보 커버리지와 실제 발행·수신 시각은 보류 근거다. 시간 표기는 KST 임시 해석이며 명목 예보 시작시각만으로 발행·수신을 증명할 수 없다. [2026 Gold 후보 manifest](evidence/data_admission_phase2/2026_gold_candidate.json)

후보 manifest는 원천과 정책 설정의 hash, 적격 판단, 보류 사유 및 기준 Gold의 물리 접근 상태를 포함한다. 2026 Gold 데이터 파일은 생성되지 않았다. 정식 Gold 실물이 연결되면 기존 검증 기록을 현재 해시와 대조하거나 `audit-gold`로 압축본과 전체 파티션을 읽기 전용 비교한다. 별도로 물리 구현된 후보가 이미 있을 경우 `--candidate-gold-manifest`로 기준 행·스키마 보존과 2026년 추가 행을 검증한다.

## P2 fingerprint와 재개

9월 보존 벤치마크 `20260922_132607_286733`의 유효 설정과 기록된 데이터 digest로 fingerprint와 최종 stage signature를 각각 6/6 재계산했다. 같은 데이터 digest와 기록 경로를 고정해도 현 구조 브랜치의 `optimized.json`은 0/6이다. 실제 현재 Gold의 digest와 Windows 정식 11월 실험은 접근이 필요하다. [모델 재사용 상세](evidence/data_admission_phase2/model_reuse.json) · [Windows 기록 접근 상태](evidence/data_admission_phase2/frozen_model_inventory.json)

| 비교 항목 | 확인한 차이 |
| --- | --- |
| 데이터와 입력 경로 | 데이터의 기록 digest를 고정해 설정 차이를 계산했다. Linux 경로 해석 차이와 실제 데이터 교체는 별도 검증 대상이다. |
| 피처와 scaler | 1h XGBoost에 비활성 `future_weather` 설정이 명시되고 24h·72h에는 미래 예보 후보가 추가됐다. 기존 관측 피처 순서·lookback은 해당 조합에서 유지된다. CNN의 저장된 Train 전처리를 읽었고 새 데이터에서 재적합하지 않았다. |
| 분할·purge | 전역 Train/Validation/Calibration/Test 종료점과 168시간 purge가 일치했다. |
| 구조·하이퍼파라미터 | CNN의 현재 설정은 capacity-factor 타깃 정책과 24h 후보 정체성이 달라졌다. 기존 선택 파라미터와 현재 설정·탐색 공간의 항목별 차이는 비교 JSON에 있다. 과거 파라미터는 재사용 검토 자료다. |
| checkpoint stage | 과거 최종 stage signature는 6/6 재현됐다. 저장 XGBoost 3개는 격리 복사에서 기존 로더 검사를 통과했다. 저장 CNN 3개는 가중치와 Train 전처리를 포함한다. 현재 데이터/설정의 fingerprint·stage signature를 검사한다. |
| Optuna study | 60개 trial 기록(완료 39, 가지치기 21)을 확인했다. 24h CNN 후보·study 이름 변경을 식별했다. Cloud에 기존 SQLite가 없어 기존 study 이어쓰기는 미검증이다. |

CPU fixture에서 CNN 1→3 epoch와 XGBoost 3→5 round 복구, 잘못된 stage signature 거절, Optuna 기존 2 trial의 중복 실행 방지와 실패 trial 파라미터 보존을 확인했다. Windows 정식 checkpoint와 DB는 실제 이름·상태를 확인한 뒤 격리 복사에서 동일한 실행 계약으로 검증해야 한다. [격리 재개 검증](evidence/data_admission_phase2/model_resume_smoke.json)

## P3 독립 평가

9월 개발 비교는 1·24·72시간마다 19개 발전소의 공통 Test 105,336행을 `2025-01-08 00:00`~`2025-12-31 23:00`에서 사용했다. 4개 모델의 저장 예측에서 MAE·RMSE·R²를 재계산했다. 저장 Hybrid 프로필의 세 파일에서는 Git의 LF를 메모리에서 CRLF로 변환한 후 원래 기록 해시가 일치했다. 이 작업에서 독립 평가용 2026년 정답은 열지 않았다. [개발 비교 지표와 독립 평가 상태](evidence/data_admission_phase2/independent_evaluation.json)

독립 평가는 같은 발전소·예측 시작·목표시각의 4개 모델 공통 키를 고정한다. Train에서만 전처리를 맞추고 Validation에서 특징·하이퍼파라미터를 선택한다. Calibration은 Hybrid 적합과 채택 판단에 쓰고 Test에서는 적합·선택 없이 평가한다. 경계 purge는 168시간이다. 예측 시작까지 실제 확보된 관측과 실제 발행·수신이 검증된 예보만 사용한다. 기존 Persistence는 정확한 시작시각의 관측을 요구하고 CNN은 그 시각을 입력창에 포함한다. 해당 값이 아직 가용하지 않은 키는 네 모델 모두에서 제외한다. 평가 전 공통 키 커버리지와 제외 사유를 발전소·horizon별로 기록하고, 동일한 MWh 정의로 지표를 산출한다.

2026년 후보를 확정하기 전 `gold-2026-phase2-v1`의 candidate hash와 학습·Validation·Calibration·피처 선택·기존 평가에 노출된 기간의 증거를 검토한다. `--exposure-review`는 후보 보고서 SHA-256, 검토 범위와 참조 자료 해시, 이미 노출된 구간을 요구한다. 검토 범위 안에서 마지막 노출 이후 168시간 purge를 지킨 기간만 후보가 된다. [평가 도구](../tools/audit_independent_evaluation.py)

## Windows 읽기 전용 검증

Phase 2 코드 checkout에서 PowerShell로 실행한다. 조사 보고서만 새 임시 경로에 쓴다. 아래 기본 입력은 `C:\ESG_Project1`을 가리킨다.

```powershell
$Root = 'C:\ESG_Project1'
$Out = Join-Path ([System.IO.Path]::GetTempPath()) ('esg-phase2-' + [Guid]::NewGuid().ToString())
New-Item -ItemType Directory $Out | Out-Null
$Index = Join-Path $Root 'artifacts\preprocessing_read_audits\20261113_resumable_raw_collection_v2\PV_ACQUISITION_REUSE_20261118_INDEX_V2.json'
$Run = Join-Path $Root 'artifacts\benchmarks\frozen_policy_regression_2025\formal_2025_comparison_gpu_io_recovery_20261101_v1'
$Gold = Join-Path $Out '2026_gold_candidate.json'
$Source = Join-Path $Out 'source_completeness.json'
python tools/audit_source_completeness.py --archive-root (Join-Path $Root 'file\solar_data_file') --raw-root (Join-Path $Root 'file\raw') --collection-index $Index --as-of 2026-11-25 --output $Source
python tools/audit_2026_gold_candidate.py --root $Root --source-audit $Source --output $Gold
python tools/audit_benchmark_model_reuse.py --inventory --project-root $Root --benchmark $Run --output (Join-Path $Out 'frozen_inventory.json')
python tools/audit_independent_evaluation.py --config config/experiments/optimized.json --development-run $Run --admission $Gold --output (Join-Path $Out 'independent_evaluation.json')

$env:CUDA_VISIBLE_DEVICES = '-1'
$env:OMP_NUM_THREADS = '1'
python -m pytest -q tests/test_benchmark_model_reuse_audit.py tests/test_model_checkpoints.py tests/test_model_optimization.py tests/test_benchmark_checkpoint_identity.py --basetemp (Join-Path $Out 'resume-tests')
```

원본 취득 증빙에서 11월 응답이 확인되지 않으면 아래 명령으로 공식 조회 화면의 10월 31일 대조군, 11월 1~25일, 11월 1~30일을 조회한다. 화면의 날짜·명시적 빈 응답·경고·접속 오류를 기록하고 CSV를 내려받지 않는다. 지금 실행한 응답은 조회 시점의 증거다.

```powershell
python tools/probe_koen_availability.py --output (Join-Path $Out 'koen_availability_20261125.json')
```

원천 보고서에서 별도 취득 인덱스의 스키마를 인식하지 못하면 원본 스캔과 인덱스 해시를 기록한다. 스키마를 확인한 뒤 인덱스의 공급자별 실제 파일 경로·해시를 대조한다. P1에는 인덱스에서 확인한 2026 입력만 `--generation 'koen=경로'`처럼 추가할 수 있다. `--asos`, `--forecast '1=경로'`, `--forecast '24=경로'`, `--forecast '72=경로'`, `--observation-availability`, `--plant-review`에는 확인된 출처 파일만 지정한다. P3는 정식 실행의 `manifest.json`, `experiment.json`, 각 `horizon_*h`의 선택·예측·Hybrid 파일을 요구한다. 정식 실행의 저장 구조나 계약이 다르면 해당 도구는 검증 오류를 반환한다.

### 전달이 필요한 최소 입력

- P0는 취득 인덱스와 연결된 KOEN·EWP·KOSPO 2026 Bronze 파일 및 수집 manifest, KOEN의 해당 월과 쿼리 응답 기록을 사용한다. 별도 EWP·KOSPO 제품은 파일에 있는 관측 범위를 확인한다.
- P1은 Gold manifest·압축본·27개 파티션, `plant_registry.csv`, `plant_quality_report.csv`, 발전량 설비·계량점 자료, `file\KMA_data_file\`, `file\forecast_weather\` 및 실제 관측 가용·예보 발행·수신 로그를 사용한다. 저장 Gold 검증은 `artifacts\verification\architecture_convergence\20261125\actual_gold_candidate.json`, 기상 준비 기록은 `artifacts\weather_full_versions\20261104_weather_full_staging_v1\manifest.json`에 있다.
- P2·P3는 `$Run`의 manifest·해석된 후보별 설정·저장 모델과 전처리·예측·선택 보고서, 일치 fingerprint의 checkpoint sidecar·payload, 일치 study SQLite 및 실행 설정을 사용한다. `artifacts\optimization\`, `artifacts\checkpoints\` 및 실행 manifest에 참조된 레거시 `optimization.db`·`.checkpoints`를 조사한다. 격리 테스트에 복사본을 쓰기 전 SQLite 작성 프로세스가 종료됐는지 확인한다. 미완료 WAL이 있으면 원본을 열어 쓰지 않고 무결한 SQLite snapshot을 확보한다.

## 코드와 검증 범위

원격 구조 개선 브랜치의 `44714db22e79ead6bdadaf8ad2522dfa7095a4c0`에서 구현을 시작했다. 작업 중 게시된 `13c4c85ffe1c0421b4bb9b4bd32453863d08d76e`의 읽기 전용 checkpoint 감사와 구조 기준 검증을 반영했다. 원격 main `3dbc598141e1c8fbae4290093aabb67d7ccddf70`에 저장된 발전량 원본과 벤치마크 결과를 재사용했다. 핵심 검증 항목은 조회 범위, 원천 공백 판정, 기존 Gold 보존, 시점 누수 방지와 checkpoint 식별자다. 최초 기준에서 전체 pytest 521개와 하위 검사 136개, 구조·색인·구문 검사 및 P0~P3 보고서 바이트 재현을 확인했다. [작업 상태](../ACTIVE_WORK.md)
