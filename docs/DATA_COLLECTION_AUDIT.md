# 공식 발전량 파일 수집 검증

검증일: 2026-08-27 (Asia/Seoul)

현재 자동화된 4개 발전사 공식 홈페이지와 공공데이터포털에서 실제 파일을 내려받고,
원본 컬럼 검증과 시간별 MWh 표준화를 수행한 결과입니다. 원본은 그대로 보존하며 표준본은
UTF-8 BOM CSV로 별도 저장합니다.

| 출처 | 검증 원본 범위 | 원본 행 | 표준 행 | 개체 수 | 원본 단위 | 키 중복 | 표준 결측 발전량 |
|---|---:|---:|---:|---:|---:|---:|---:|
| 한국남동발전 | 2025-10 | 713 일별 개체행 | 17,112 | 23 발전기 | kWh | 0 | 0 |
| 한국남부발전 | 2025-01-01~2025-11-11 | 315 일별 행 | 7,560 | 1 발전기 | kWh | 0 | 0 |
| 한국동서발전 | 2020-01-01~2024-12-31 | 806,408 | 738,905 | 17 지역 | MWh | 0 | 0 |
| 한국서부발전 | 2017-01-01~2023-06-30 | 7,109 일별 행 | 170,553 | 3 발전기 | Wh | 0 | 0 |

동서발전 원본에는 2024년 `시도명 + 발전일자` 중복 키 66,976건과 발전량 결측 527건이
있습니다. 동일 키의 뒤쪽 레코드는 기존 값보다 정밀도가 높은 수정값이므로 표준본은 뒤쪽
레코드를 사용하고, 타깃 결측 행은 학습 대상에서 제외했습니다. 서부발전 원본의 시간 발전량
결측 셀 63개도 표준 시간 행에서 제외했습니다.

## 표준 데이터 계약

네 발전사 개별 설비 발전량:

```text
timestamp,company,plant_id,plant,unit,energy_source,generation_mwh,capacity_mw,
tilt_deg,latitude,longitude,address,source_file
```

동서발전 전국 학습자료:

```text
timestamp,company,region,capacity_mw,temperature_c,rainfall_mm,
humidity_pct,snowfall_mm,wind_speed_mps,cloud_amount_tenths,
cloud_amount_thirds,sunshine_hours,extraterrestrial_irradiance,
solar_irradiance,is_leap_year,hour,dayofweek,month,generation_mwh
```

검증 규칙은 다음과 같습니다.

- 원본 필수 컬럼과 24개 시간 컬럼 존재 여부 검사
- `Wh`/`kWh`를 `MWh`로 명시적 변환
- 남동발전 기존 CSV의 `MWh` 오표기는 최신 `KWh` 헤더 및 물리적 값 범위와 교차검증해 보정
- `1시`→`00:00`, `24시`→`23:00` 변환
- 표준 키 중복 거부
- 잘못된 일시와 발전량 결측 제거
- 음수 발전량, 최종 결측, 키 중복 집계

## 보관 원본 전체 재표준화 결과

`python app.py prepare-data`로 `file/solar_data_file/`의 88개 CSV를 모두 다시 읽어 같은 계약으로
변환했습니다. 출력은 원본 파일 단위 gzip 파티션이므로 대용량 전체를 메모리에 올리지 않고 사용할 수
있습니다. 각 파티션은 deterministic gzip 설정과 atomic replace로 쓰고, manifest에는 원본·산출물
byte와 SHA-256을 남겨 동일 파일명의 수정본도 식별합니다.

| 회사 | 원본 파일 | 표준 시간 행 | registry 발전소×발전원 | 기간 | 용량 커버리지 | 설치각 커버리지 | 좌표 커버리지 |
|---|---:|---:|---:|---|---:|---:|---:|
| 남동발전 | 45 | 737,496 | 15 | 2022-01-01~2025-09-30 | 66% | 0% | 0% |
| 남부발전 | 38 | 3,093,580 | 36 | 2013-01-01~2026-02-28 | 85% | 80% | 0% |
| 동서발전 | 3 | 182,400 | 7 | 2022-04-01~2025-06-30 | 100% | 0% | 77.99% |
| 서부발전 | 2 | 223,089 | 8 | 2017-01-01~2023-12-31 | 100% | 0% | 0% |
| 합계 | 88 | 4,236,565 | 66 | 2013-01-01~2026-02-28 | - | - | - |

개별 파티션의 키와 단위를 검증한 결과 음수 발전량과 변환 실패 파일은 0건입니다. 공식 파일의
완전 동일 중복 행 2개는 제거했습니다. 누적 공개 snapshot끼리 겹치는 키는 Silver에서 보존하고,
Gold에서 파일명의 최신 `YYYYMMDD` 개정값을 선택해 snapshot을 합산하지 않습니다.
발전원별로는 태양광 4,112,437행, 소수력 106,608행, 풍력 17,520행입니다. 서부발전 풍력과
남부발전 소수력을 원본 표준화에서 삭제하지 않고 `energy_source`로 분리했으며, 태양광 모델
학습 설정에서만 `solar`를 선택합니다.

남부발전 사양 파일은 발전기 단위 용량·설치각을 우선 매칭합니다. `하동화력 #5`의 `997.56W`는
같은 행의 `340W × 2,934 모듈`, `500kW × 2 인버터`와 모순되므로 공식 메타데이터의 단위 오기로
판단해 997.56kW로 보정했습니다. 여러 호기가 있는데 총용량만 공개된 경우에는 각 호기에 총용량을
복제하지 않고 결측으로 남깁니다.

## 추가 컬럼 선택 검증

아래 첫 표는 기존 `val.csv`에 기록돼 있던 단일 2022 Train/2023 Validation 실험이며, 현재
선정 근거가 아니라 과거 비교 기준으로만 보존합니다.

| 피처 계약 | 개수 | MAE | RMSE | R² | 주간 MAE |
|---|---:|---:|---:|---:|---:|
| 기존 기상·시간 | 7 | 0.17824 | 0.48733 | 0.25233 | 0.25314 |
| selected_v2_day_ahead | 23 | 0.03773 | 0.16005 | 0.91935 | 0.04562 |

선정 계약은 기온·강수·일조·일사에 풍속·습도·전운량·중하층운량을 추가하고, 시간 주기,
24/168시간 lag, 24시간 이동한 7일 평균, 용량·설치각·관측지점 좌표/고도를 사용합니다.
당시 생성된 `model_ready.csv`는 328,776행, 19개 발전소, 2022-01-09~2023-12-30이며 용량 94.74%,
설치각 89.47%, 관측지점 좌표 99.998%의 커버리지를 가집니다.

증기압·적설·최저운고를 더한 실험은 compact 기상 조합보다 MAE가 소폭 나빠 기본 계약에서
제외했습니다. 동서발전 하루 전 예측 API의 발전량 예측값은 목표값을 우회할 수 있으므로 외부
기준모델로 분리하며 기본 피처에는 포함하지 않습니다.

위 표는 과거 19개 발전소 병합본에 소수력이 태양광으로 포함된 상태의 기존 기록입니다. 발전원을
다시 분류하고 공식 원본에서 타깃을 재구성한 현재 실험은 마지막 10% Calibration과 15% Test를
완전히 예약하고, 그 이전 구간에서 168시간 gap과 2,160시간 Validation 창을 둔 3개 expanding
rolling-origin fold를 사용했습니다. XGBoost에는 중앙값을 넣지 않고 native NaN 분기를 사용했습니다.

| 피처 계약 | 개수 | 평균 MAE | 최악 fold MAE | 평균 RMSE | 평균 R² | 평균 주간 MAE |
|---|---:|---:|---:|---:|---:|---:|
| selected_v3_physics_aware | 26 | **0.03904** | **0.04741** | **0.16463** | **0.92257** | **0.07357** |
| + 이력/기상 결측 문맥 | 33 | 0.03914 | 0.04742 | 0.16545 | 0.92190 | 0.07357 |
| + 기상 관측 mask | 30 | 0.03937 | 0.04750 | 0.16546 | 0.92183 | 0.07416 |
| selected_v2_day_ahead | 23 | 0.04028 | 0.04774 | 0.16888 | 0.91884 | 0.07643 |

태양고도·clear-sky proxy·주야만 더한 26개 조합이 23개 대비 평균 MAE를 3.09% 개선했습니다.
결측 문맥 컬럼은 품질 분석 테이블에 보존하되 기본 모델 입력에서는 제외합니다. 실험 원본은
`output/evaluation/features/feature_ablation.csv`와 `feature_ablation_folds.csv`이며, 이 선택에는
예약한 Calibration/Test를 사용하지 않았습니다.

history 결측 때문에 행을 삭제하던 정책도 제거했습니다. 현재 전국 Gold는 2,230,839행,
37개 발전소(태양광 35, 풍력 1, 소수력 1), 2013-01-01~2025-09-30입니다. 모델 설정은
`energy_source=solar`와 `quality_train_eligible=true`를 chunk 읽기 단계에서 먼저 적용하며
2,115,471행을 유지합니다.

## 발전소 품질 진단

`legacy_pipeline_quality_report.csv`로 19개 병합 발전소를 공식 표준 원본과 대조한 결과는
`pipeline_artifact` 9곳, `review` 1곳, `low` 9곳, 원본 센서 `high` 0곳입니다.
하동변전소·하동보건소·하동정수장·하동하수처리장은 병합 학습본에서 주간 양의 발전량이 4시간
이상 정확히 같은 flatline 비율이 약 9.2~10.3%였지만, 같은 2022~2023 공식 원본에서는 0%였습니다.
남제주소내·부산복합자재창고·부산수처리장·부산신항·행원소수력도 병합본의 1% 이상 반복값이
원본에는 없거나 0.03% 미만이었습니다. 따라서 이 패턴은 센서 고장보다 기존 선형보간/중앙값
전처리가 만든 인공 패턴일 가능성이 높으며, 향후 학습본은 표준 원본에서 다시 구성해야 합니다.

현재 `model_ready.csv.gz`와 회사×연도 Gold 파티션은 이 권고대로 공식 원본 발전량을 다시 집계해
생성합니다. 기존 병합본은
발전소별 지역·ASOS 지점 매핑에만 사용합니다. 새 `plant_quality_report.csv`에서는 legacy
flatline이 사라졌고, peer 비교도 같은 지역이 아니라 `같은 지역 + 같은 발전원` 안에서만 수행합니다.
재구성 후 결과는 `low` 32곳, `review` 5곳, `high` 0곳입니다. 검토 표시는 공개 자료만으로 원인을
확정할 수 없는 peer/시간 패턴을 사람이 확인해야 한다는 뜻이며 자동 고장 판정이 아닙니다.

음수 발전량은 0으로 바꾸지 않고 `quality_negative_generation`으로 표시해 학습에서 제외합니다.
주간 0, 용량초과, flatline은 출력제어·메타데이터 오류·정비와 구별할 수 없으므로 자동 삭제하지
않고 `quality_review_required`로 보존합니다. 강수 공란은 다른 핵심 기상센서가 정상일 때만 0으로
해석하고 관측 여부 mask를 남기며, 일조·일사 공란은 계산 태양고도가 야간일 때만 0으로 바꿉니다.

## 추가 발전소 후보 staging

남부발전 누적 공개파일 `[용수리]태양광발전실적_20260228.csv`(2020-10-01~2026-02-28,
1,977일)와 `[신풍리]태양광발전실적_20260228.csv`(2021-05-01~2026-02-28, 1,765일)를
공공데이터포털에서 내려받아 89,808개 시간행을 추가 표준화했습니다. 두 파일은 발전량 기간은
충분하지만 주소·좌표가 없어 근거 없는 KMA 지점을 붙이지 않고 registry에서 격리했습니다.

한국농어촌공사 진도 2019 파일도 내려받아 365개 유효 일행과 24시간 wide 스키마를 확인했습니다.
다만 1년 자료는 공통 4구간 평가에 부족하므로 어댑터 후보로 staging하고 학습에는 편입하지 않습니다.

한국농어촌공사 영암 태양광 주기성 과거파일 2020~2025를 별도 후보 원본으로 확보했습니다.
2022~2025의 영암1차·영암2차·율치 3개소는 총 105,191개 유효 시간 행, 중복 0건, 음수 0건이며
24개 시간 셀 결측은 1건입니다. 2020~2021은 발전소 식별 컬럼이 없어 격리했습니다. 발전량 품질은
충분하지만 `kW`로 표시된 시간 버킷의 에너지 단위 해석과 ASOS 매핑이 검토되지 않아 현재 Gold에는
아직 합치지 않았습니다. 상세 후보와 편입 게이트는
[`ADDITIONAL_DATA_SOURCE_AUDIT.md`](ADDITIONAL_DATA_SOURCE_AUDIT.md)를 따릅니다.

## 생성 파일

```text
file/raw/koen/한국남동발전/2025/남동발전량_2025_10.csv
file/raw/koen/normalized/koen_generation_2025_10.csv
file/raw/kospo/kospo_busanjrail2_solar_20251111.csv
file/raw/kospo/normalized/kospo_busanjrail2_solar_20251111_normalized.csv
file/raw/ewp/ewp_solar_training_2020_2024.csv
file/raw/ewp/normalized/ewp_solar_training_2020_2024_utf8.csv
file/raw/iwest/iwest_solar_status_20230630.csv
file/raw/iwest/normalized/iwest_solar_status_20230630_normalized.csv
file/standardized/generation/<company>/*_standardized.csv.gz
file/standardized/generation_manifest.json
file/standardized/model_ready.csv.gz
file/standardized/model_ready_parts/company=<provider>/year=<YYYY>/part.csv.gz
file/standardized/plant_registry.csv
file/standardized/model_ready_manifest.json
file/standardized/plant_quality_report.csv
file/standardized/quality_manifest.json
file/standardized/candidates/krc_yeongam/candidate_generation.csv.gz
file/standardized/candidates/krc_yeongam/candidate_manifest.json
```
