# 품질 인지형 다발전소 예측 개선안

검토일: 2026-08-27 (Asia/Seoul)

이 문서는 기존 발표자료의 전처리·모델 비교를 현재 코드와 대조하고, 여러 공개기관 발전소의 서로 다른
설비·센서 패턴을 한 모델에서 다루기 위한 개발 기준을 정리합니다. 논문의 개선 폭은 서로 다른
데이터셋에서 나온 결과이므로 이 프로젝트의 성능 향상을 보장하지 않습니다. 후보는 반드시 같은
rolling-origin 분할과 같은 입력정보로 비교합니다.

## 먼저 바로잡은 데이터 계약

서부발전 신재생에너지 파일의 풍력 행을 원본 표준화에서 제외하면 안 됩니다. 풍력은 태양광
예측값을 늘리는 보조 피처가 아니라 물리와 목표 분포가 다른 별도 발전원입니다. 따라서 다음처럼
두 경계를 분리합니다.

- 보존·표준화 계층: 태양광과 풍력을 모두 보존하고 `energy_source`를 `solar`, `wind`,
  `hydro`, `unknown` 중 하나로 기록합니다.
- 태양광 모델 계층: `energy_source == "solar"`만 태양광 목표로 학습합니다.
- 신재생 포트폴리오 계층: 풍력을 태양광 출력에 합쳐 학습하지 않고, 발전원 임베딩과 발전원별
  head를 둔 multi-task 모델로 별도 실험합니다.

서부발전 풍력은 태양광 단일모델의 타깃에는 직접 영향을 주지 않지만, 신재생 총발전량 예측과
지역 간 공동 변동 학습에는 영향을 줍니다. 실제로 풍력·PV 이질성을 명시적으로 보존한 동적
이종 그래프 모델이 다지역 공동예측에 사용됐습니다
([IEEE Transactions on Power Systems, 2024](https://ieeexplore.ieee.org/document/10356761)).

같은 점검에서 기존 병합본의 `행원소수력`이 태양광으로 학습되고 있던 것도 확인했습니다.
발표자료에서 이 개체의 성능이 유독 낮았던 이유를 모델 구조만으로 해석하면 안 됩니다. 현재는
소수력도 `hydro`로 보존하고 태양광 학습에서 제외합니다.

공통 발전량 계약은 다음과 같습니다.

```text
timestamp,company,plant_id,plant,unit,energy_source,generation_mwh,
capacity_mw,tilt_deg,latitude,longitude,address,source_file
```

## 기존 발표자료 방식에서 바꿀 점

발표자료에는 강수 결측을 월평균으로 채우고, 발전량 결측을 선형보간·중앙값으로 보완하며,
전역 Z-score 3σ 밖을 제거한 뒤 중앙값으로 치환하고, 기온 외 음수를 0으로 바꾸는 방식이
기록돼 있습니다(슬라이드 8, 12). 이 방식은 다음 문제를 만듭니다.

| 기존 처리 | 위험 | 변경 기준 |
|---|---|---|
| 강수 결측 → 월평균 | 비가 오지 않은 시간과 관측 실패를 혼동하고 가짜 강수를 만듦 | 원천의 무강수 표기일 때만 0, 관측 실패는 결측 마스크 유지 |
| 발전량 결측 → 선형보간/중앙값 | 일출·구름 ramp와 정지 구간을 인공적으로 매끈하게 만듦 | 짧은 결측도 물리·인접 발전소 근거가 있을 때만 별도 복원; 학습 loss mask 우선 |
| 전역 3σ | 대용량 발전소·계절 피크를 이상치로 오인 | 발전소별 용량비, 태양 위치, 일사량, 변화율, peer residual을 함께 사용 |
| 기온 외 음수 → 0 | 센서 오류를 실제 무발전·무강수와 합침 | 물리적으로 불가능한 음수는 결측으로 전환하고 원인 flag 보존 |
| 잔차 상위 1% 고정 | 데이터가 정상이어도 항상 정확히 1%를 이상치로 만듦 | Calibration 구간의 조건부 임계값 또는 conformal interval 이탈로 판정 |

현재 품질 정책은 원값을 조용히 덮어쓰지 않고 `quality_code`,
`quality_train_eligible`, `quality_review_required`, 음수·용량초과·주간 0·flatline·기상범위
위반 flag와 `capacity_factor`를 생성합니다. 음수 발전량과 일 총량이 최소 30일 동안 95% 이상
동일한 단일 야간 버킷에 적재된 태양광 자료는 시간별 정답으로 쓸 수 없어 자동 학습 제외합니다.
용량초과·주간 0·flatline은 메타데이터 오류·출력제어·정비와 구별할 수 없으므로 검토 대상으로만
표시합니다.

발표자료의 모델 비교표(슬라이드 37~38)는 XGBoost의 MAE와 CNN의 MSE가 같은 열에 배치되어
동일 지표 비교가 아닙니다. 기존 수치를 기준선으로 보존하되, 새 실험에서는 동일 타깃·동일
기간·동일 단위의 MAE/RMSE/R²/nMAE를 다시 계산해야 합니다. 또한 전체 잔차 상위 1% 방식
(슬라이드 41, 44~46)은 이상치 개수 자체를 성능 근거로 사용하지 않습니다.

## 발전소마다 패턴이 다를 때의 표준화

표준화는 하나가 아니라 세 층으로 분리합니다.

1. 물리 표준화
   - 단위를 시간별 MWh로 통일합니다.
   - 설비용량이 검증된 발전소는 `capacity_factor = generation_mwh / capacity_mw`를 보조
     타깃으로 사용합니다. 최종 출력은 다시 MWh로 환산합니다.
   - 좌표가 확보되면 태양고도·방위각·clear-sky irradiance와
     `clear_sky_index = actual / clear_sky_expected`를 추가합니다.
   - 설비용량이 불확실한 행은 억지로 전체용량을 호기에 복제하지 않고 결측과 flag를 유지합니다.

2. 통계 표준화
   - 전체 발전소를 한 지역 평균 scaler로 맞추지 않습니다.
   - XGBoost는 중앙값으로 덮지 않고 native `NaN` 분기 방향을 Train에서 학습합니다.
   - CNN은 Train 행으로만 중앙값을 적합하고 모든 원본 피처와 같은 순서의 missing mask 채널을
     붙입니다. Train에서 전부 비어 있는 피처도 미래 통계를 빌리지 않고 `0 + mask`로 표현합니다.
   - 딥러닝은 발전소별 분포 이동을 줄이는 RevIN 후보를 우선 비교합니다. RevIN은 입력 window의
     통계를 제거한 뒤 출력에서 복원하도록 설계됐습니다
     ([ICLR 2022](https://openreview.net/pdf?id=cGDAkQo1C0p)).
   - 전역 MinMax는 극단 센서값에 민감하므로 Train-only Robust scaling과 RevIN을 각각 ablation합니다.

3. 개체·관계 표준화
   - 모든 발전소가 동일한 피처 컬럼을 사용하되, 공유 backbone에 발전원·용량·설치각·좌표 같은
     공통 static covariate를 제공합니다. 현재 `plant_id`는 시퀀스 경계에만 사용하고 입력에는
     직접 넣지 않습니다.
   - plant embedding은 알려진 발전소 정확도에는 유리할 수 있지만 새 발전소에서 ID 암기로
     실패할 수 있으므로, `unknown plant` 처리와 leave-one-plant-out 성능을 함께 검증합니다.
   - 데이터가 적은 발전소는 공유 backbone에서 통계력을 빌리고, 충분한 발전소는 작은 adapter/head로
     고유 패턴을 학습하는 partial-pooling 구조를 사용합니다.
   - 거리만으로 같은 지역 발전소를 묶지 않습니다. 거리·동일 기상계·Train 기간 residual 상관으로
     edge를 만들고, 필요하면 관계 자체를 학습합니다. 계층 관계를 학습하고 합계 일관성을 보장하는
     그래프 기반 방법도 보고돼 있습니다
     ([ICML 2024](https://proceedings.mlr.press/v235/cini24a.html)).

`plant_quality_report.csv`는 시간 커버리지, 주간 0, 양의 값 flatline, 용량초과, 기상 결측,
일별 형태 일관성, 같은 지역 peer 상관을 발전소별로 기록합니다. `high`는 고장 확정이 아니라
학습 전 원천·계량상태 확인이 필요한 위험도입니다.

병합 학습본과 공식 원본을 같은 기간으로 대조하자 9개 발전소의 1% 이상 양의 flatline이 원본에는
없거나 0.03% 미만이었습니다. 특히 하동 4개 발전소는 병합본 약 9.2~10.3%, 공식 원본 0%였습니다.
이 항목들은 센서 위험이 아니라 `pipeline_artifact`로 재분류했습니다. 따라서 기존 병합본의
flatline을 결측복원 모델로 다시 채우기 전에, 표준 원본에서 학습 테이블을 재구성해야 합니다.

공식 원본 타깃으로 다시 만든 현재 ablation은 마지막 10% Calibration과 15% Test를 예약하고,
그 이전 구간에서 168시간 gap을 둔 3개 expanding rolling-origin fold를 사용합니다. 기존 23개
피처의 평균 MAE 0.04028 대비 태양고도·clear-sky proxy·주야를 더한 26개 계약이 0.03904로
3.09% 개선되어 기본 계약으로 승격됐습니다. 기상 mask를 더한 30개는 0.03937, 이력 가용성까지
더한 33개는 0.03914로 더 나빠 기본 입력에서 제외했습니다. 이 선택에는 Calibration/Test를
사용하지 않았습니다.

2026-08-31에 무증빙 legacy ASOS 매핑을 차단하면서 태양광 Gold는 22개·701,011행으로
변경됐습니다. 시간 해상도 품질 게이트가 여수태양광 29,280행을 제외한 실제 학습 적격 모집단은
21개·671,731행입니다. 위 18개 표본 ablation은 재현 가능한 이전 기준선으로만 유지하며, 새 모집단에서는
같은 rolling-origin 계약으로 `evaluate-features`와 Optuna를 다시 실행합니다. 이전 최적 피처나
하이퍼파라미터를 새 데이터의 정답으로 하드코딩하지 않습니다.

## 결측 처리 개선 실험

결측은 `센서 통신 단절`, `원천에서 비관측`, `발전소 신규 편입에 따른 cold-start`, `물리적으로
무효인 값`으로 원인을 나눕니다. 무작위로 셀을 지운 MCAR 실험 하나만으로 방법을 고르지 않습니다.
실세계 결측 메커니즘과 연속 block 결측에서 방법 순위가 달라질 수 있다는 2025 benchmark가 이를
지적했습니다
([PMLR 2025](https://proceedings.mlr.press/v287/toye25a.html)).

권장 stress test는 정상 관측 구간에서 1·3·6·24·72시간 block을 가리고, 발전소 전체 단절과
기상요소 하나만의 단절을 따로 생성한 뒤 다음을 비교하는 것입니다.

1. XGBoost native NaN과 CNN median+mask를 현재 기준선으로 둡니다.
2. 짧은 결측 선형보간은 예측 피처에만 제한적으로 비교하고, 발전량 정답을 만들어 학습시키지
   않습니다. 실제로 보간이 항상 복잡한 모델보다 나쁘지는 않으므로 방법명보다 결측 메커니즘별
   bias와 downstream forecast MAE를 봅니다.
3. 연속 block 결측이 많아지면 결측 mask를 latent representation과 함께 처리하는 S4M을 후보로
   비교합니다. S4M은 별도 imputation 후 예측하는 2단계를 피하는 missing-aware S4 구조입니다
   ([ICLR 2025](https://proceedings.iclr.cc/paper_files/paper/2025/hash/7b2f0758334389b8ad0665a9bd165463-Abstract-Conference.html)).
4. 동시 시각의 다발전소와 관계 graph가 충분할 때는 결측 패턴에 조건화된 시공간 downsampling
   모델을 비교합니다. 이 방법은 특히 contiguous block missing을 대상으로 평가됐습니다
   ([ICML 2024](https://proceedings.mlr.press/v235/marisca24a.html)).
5. 확률구간이 필요한 불규칙 시계열에는 결측 시점·채널을 조건으로 공동 미래분포를 학습하는
   ProFITi를 장기 후보로 둡니다
   ([AAAI 2025](https://ojs.aaai.org/index.php/AAAI/article/view/35494)).

승격 기준은 imputation RMSE 하나가 아니라 최종 24시간 발전량 MAE/nMAE, ramp MAE, 편향,
예측구간 coverage, 결측률별 열화를 함께 개선하는지입니다.

## 모델 후보와 적용 순서

### 0단계: 누수 없는 기준선

- XGBoost를 필수 기준선으로 유지합니다.
- Train/Validation/Calibration/Test를 모든 발전소에 같은 날짜로 분리하고 각 경계 앞에 최대
  lookback만큼 purge gap을 둡니다.
- XGBoost는 native NaN을 사용하고, 결측 통계나 scaler가 필요한 모델은 Train에서만 적합합니다.
- Test는 모델·피처·threshold 선택에 사용하지 않습니다.
- 기존 발전소의 미래 예측은 전역 시간 Test로, 신규 발전소 cold-start는 plant group holdout으로
  따로 보고합니다.

### 1단계: TimeXer + RevIN

가장 먼저 비교할 딥러닝 후보입니다. 발전량은 endogenous series, 예보 발행시각이 보존된
기온·일사·구름·풍속 등은 exogenous series로 분리합니다. TimeXer는 외생 시계열을 단순히 모든
채널과 동등 취급하지 않고 patch-wise endogenous attention과 variate-wise cross-attention으로
결합하며 12개 실데이터 benchmark에서 평가됐습니다
([NeurIPS 2024](https://proceedings.neurips.cc/paper_files/paper/2024/hash/0113ef4642264adc2e6924a3cbbdf532-Abstract-Conference.html)).

권장 구성은 168~672시간 lookback, 24개 직접 출력 head, RevIN, plant embedding, quality mask,
known-future weather, quantile loss입니다. 현재 ASOS 관측값을 미래 기상처럼 넣으면 누수이므로,
운영 실험에는 반드시 예보 발행시각 기준 weather forecast가 필요합니다.

### 2단계: PatchTST/TimeMixer 대조군

PatchTST는 시계열을 patch token으로 만들어 긴 문맥의 계산량을 줄이고 channel-independent
weight sharing을 사용합니다
([ICLR 2023](https://openreview.net/pdf?id=Jbdc0vTOcol)). TimeMixer는 미세·거시 시간축의
계절/추세 패턴을 분해해 섞는 대조 후보입니다
([ICLR 2024](https://openreview.net/pdf?id=7oLshfEIC2)). 복잡한 모델이 항상 우수하지 않으므로
빠른 TiDE 기준선도 포함합니다
([TMLR 2023](https://research.google/pubs/long-horizon-forecasting-with-tide-time-series-dense-encoder/)).

### 3단계: 다발전소 그래프 모델

동시 시각 데이터와 좌표 커버리지가 충분해진 뒤 적용합니다. 센서 결측 복원은 시간과 공간을
함께 쓰는 GRIN을 비교할 수 있습니다
([ICLR 2022](https://openreview.net/pdf?id=kOu3-S3wJ7)). 풍력·태양광 포트폴리오는 발전원별
node type/head를 둔 heterogeneous graph로 분리합니다. 단, 그래프 모델이 단일 발전소 기준선보다
실제 rolling-origin Test에서 나을 때만 운영 후보로 승격합니다.

## Hybrid는 고정 가중치가 아님

현재 Hybrid v2는 한 개의 전국 고정 비율을 사용하지 않습니다. Validation에서 발전소, 지역,
시간, 일출/ramp/주간/야간 regime, 두 모델의 예측 불일치 정도별로 최적 convex 결합비와 모델별
MAE를 학습합니다. Test 행은 가장 구체적이면서 표본이 충분한 근거를 선택하고 부족하면
`plant_hour → region_regime_disagreement → region_hour → plant → region → global` 순으로
후퇴합니다. 결과에는 scope, 표본 수, 예상 MAE, 불일치 구간, 결합비와 문장형 근거를 저장합니다.

이 구현은 설명 가능한 contextual router이며 완전한 end-to-end MoE는 아닙니다. 다음 단계에서는
각 expert가 시간 패턴에 특화되고 router가 입력별로 조합하는 MoE를 controlled ablation합니다.
MoLE은 이런 적응형 router가 단일 예측 규칙의 주기 변화 한계를 완화할 수 있음을 보였습니다
([AISTATS 2024](https://proceedings.mlr.press/v238/ni24a.html)). 고정 50:50이나 Test 오차로 정한
가중치는 사용하지 않습니다.

## 이상징후와 불확실성 평가

- 점 예측 오차만 보고 고장을 판정하지 않습니다.
- 물리 위반, 데이터 품질, 기상 공통충격, 한 발전소만의 residual 이탈을 분리합니다.
- 잔차 임계값은 독립 Calibration에서 발전소·시간대·일사 regime별로 고정합니다. Validation은
  모델·피처·Hybrid gate 선택에만 사용합니다.
- 24시간 전체 구간에는 quantile forecast와 multi-horizon conformal calibration을 비교합니다.
  ConForME는 다중 horizon의 공동 coverage를 대상으로 설계됐습니다
  ([PMLR 2024](https://proceedings.mlr.press/v230/galvao-lopes24a.html)).
- 최종 보고는 MAE/RMSE/R² 외에 capacity-normalized MAE, ramp MAE, interval coverage/width,
  센서 품질 등급별 성능을 함께 냅니다.

## 구현 상태와 다음 실험

- 완료: 풍력·수력 원본 보존, `energy_source` 계약, 공식 원본 기반 학습 타깃 재구성,
  물리·문맥 품질 flag, legacy-vs-raw 품질 보고서, XGBoost native NaN, CNN Train-only 중앙값과
  missing mask, 전역 4구간 분할, lazy sequence, 26개 물리 피처, Hybrid v2 근거 계층,
  회사×연도 Gold 파티션과 10만행/float32/1.5GB 학습 로더, XGBoost/CNN-BiLSTM의 재개 가능한
  Validation-only Optuna와 trial/시간 상한·pruning.
- 추가 데이터 완료: 한국농어촌공사 영암 2020~2025·진도 2019, 남부발전 용수리·신풍리 최신
  원문 확보. 영암 2022~2025는 개체 안전 정규화·단위 물리검증·검토 ASOS 매핑 후 Gold 편입했고,
  식별자 없는 2020~2021과 기상이 아직 없는 2026 구간만 quarantine/withhold.
- 다음: 좌표/용량 미확보 발전소 보강, 예보 발행시각이 있는 기상 예보 수집,
  plant embedding group-holdout ablation, capacity-factor 보조 타깃.
- 그 다음: XGBoost → TimeXer+RevIN → PatchTST/TimeMixer → graph model 순서의 동일 조건 실험.
- 운영 승격 조건: 최소 3개 rolling-origin fold에서 기준선 개선, 발전소별 치명적 열화 없음,
  calibration coverage 충족, 센서 결손 stress test 통과.
