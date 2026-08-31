# 파일 포맷·발전소 매핑·대시보드 구조 감사

검증일: 2026-08-31 (Asia/Seoul)

## 결론

- 한글이 실제로 대체문자(`�`)로 손상된 원본은 발견되지 않았다.
- 현재 감사 대상 CSV 107개는 CP949 55개, UTF-8-SIG 52개다. 공급기관 원본 형식이 섞인 것이며
  일괄 재인코딩할 이유가 없다.
- Bronze는 원본 byte와 SHA-256을 보존하고, Silver/Gold만 영문 공통 컬럼과 UTF-8-SIG로 쓴다.
- 새 발전사 다운로드는
  `발전소명_[세부발전소명] 태양광발전실적_다운로드날짜.csv`로 생성한다.
- 과거 파일 69개는 공급기관 원명 또는 이전 규칙이다. rename하면 manifest의 source path와
  원본 hash lineage를 불필요하게 바꾸므로 그대로 두고 새 수집분부터 규칙을 강제한다.

Windows 터미널이나 Excel에서 한글이 깨져 보이는 현상은 파일 손상과 구분해야 한다. CP949 파일을
UTF-8로 강제해서 열거나, BOM 없는 UTF-8을 Excel이 로컬 코드페이지로 추정하면 화면만 깨진다.
수집 manifest는 각 CSV의 감지 인코딩, BOM, byte, SHA-256, `provider_original_bronze` 또는
`standardized_silver` 역할을 남긴다.

## 발전소·지역·기상 매핑

현재 registry는 68개 자산이며 태양광 65, 풍력 2, 소수력 1이다. 태양광 중 공식 주소·좌표 또는
근거가 있는 reviewed ASOS 매핑을 통과한 22개만 태양광 Gold에 결합하고 43개는 격리한다.
Gold 22개 중 시간 해상도 품질 게이트까지 통과한 학습 적격은 21개·671,731행이다. 풍력 2개는
원본 보존용 Gold에 남고 소수력 1개는 무근거 legacy 지점번호만 있어 격리한다. 두 태양광 모델은
`energy_source=solar`만 읽는다.

매핑은 다음 세 필드를 서로 대체하지 않는다.

1. `admin_province/admin_city`: 공식 주소에서 얻은 행정구역
2. `latitude/longitude`: 공식 자료에 실제 발전소 좌표가 있을 때만 기록
3. `weather_station_id/name`: 기상 결합에 사용한 ASOS와 근거 방식

registry의 실좌표는 5개, ASOS 결합은 공식/검토 근거를 통과한 24개, 미매핑은 44개다. 대리좌표에는
`location_basis=weather_station_proxy`를 붙여 발전소 위치처럼 보이지 않게 한다. Registry 자동
검사는 plant ID 유일성, Gold 결합 태양광의 ASOS 존재, 좌표 쌍 완전성, 대한민국 범위,
미해결 자산 격리를 모두 통과했다.

Gold 27개 회사×연도 파티션의 718,531행을 필요한 식별 컬럼만 파티션 단위로 읽어 registry와
대조했다. 24개 Gold 자산의 `plant_id`, 회사, 표준 발전소명, 발전원, 행정구역, ASOS 지점,
매핑 방법 불일치는 0건이다. 이 검사는 학습 데이터 생성·검증 테스트에서 유지한다. 사용자용
대시보드 JSON에는 내부 검사명이나 매핑 감사 결과를 포함하지 않는다.

과거 `val.csv`의 지점번호는 19개 직접 자산과 동일주소 전파 3개를 근거 없이 `high`로 승인해
기존 Gold 1,810,752행에 영향을 줬다. 새 registry는 이 값을 audit-only 후보 열에만 보존하고
동일주소 전파를 제거한다. stale registry에 `reviewed_legacy_mapping` 또는
`reviewed_colocated_address`가 남아 있어도 모델 로더가 다시 차단한다. 전체 재생성 뒤 unsafe
method 학습행, `weather_mapping_review_required=true` 학습행, reviewed config 근거 누락은 모두
0건이다. 이는 ASOS 매핑 승인 상태이며 발전량의 `quality_review_required`와는 별도 게이트다.

행정구역과 기상관측소는 의도적으로 별개다. 예를 들어 전남 영암은 목포 ASOS를 사용할 수 있지만
행정구역을 목포로 바꾸지 않는다. 공식 주소가 없는 자산의 Gold `region`은 `unknown`으로 유지하고
ASOS 이름은 별도 컬럼에 둔다.

## 디렉터리와 모델 산출물

```text
src/solar_forecast/models/       모델 구현과 공통 체크포인트 코드
artifacts/models/                새 모델 결과
artifacts/checkpoints/           fingerprint 기반 재개 상태
artifacts/optimization/          Optuna SQLite와 trial 결과
artifacts/legacy/cnn_bilstm/     보존한 2025년 구형 발전소별 체크포인트
dashboard/                       현재 웹 화면과 공유 assets/data
```

루트 `xgboost/`는 빈 껍데기였고, 루트 `cnn_bilstm/`은 구형 출력만 포함해 제품 구조에서 제거했다.
구형 `.pt` 45개(343,373,095 byte)는 삭제하지 않고 ignored legacy 경로로 옮겼다. 낡은 정적
`plant_region_report_perm.html`은 사용자용 `model_analysis.html`로 이동하는 호환 진입점으로
축소했다. registry 검사명·인코딩·파일 경로는 사용자 화면에서 제거했다.

## HTML 구성 판단

`solar_dashboard.html`은 학습 데이터 보유 여부와 무관한 전국 태양광 발전설비 현황으로 복원했다.
2026-08-28에 EPSIS 발전기 현황 화면에서 `발전원=태양에너지`로 직접 재수집한 CP949 CSV를
Bronze 원본으로 보존한다. 공식 기준일은 2026-08-05이며 합계행 제외 후 188,594행,
33,059.516180 MW이다. 세부지역 320개 원본 표기는 2026-07-01 행정구역 reference와 검토된
raw-row override를 적용해 258개 집계로 통합한다. [행정안전부 행정표준코드관리시스템](https://www.code.go.kr/stdcode/regCodeL.do)의
2026-07-01 변경을 기준으로 과거 광주·전남과 새 표기를 `전남광주통합특별시`로 합쳐 16개 최신
광역 기준으로 표시하되 원본 fact와 총량은 변경하지 않는다. 이 공식 명칭 통합은 원본 충돌 행의
실제 소속을 추정하는 규칙과 분리한다. 메인 지도에는 세부지역 점을 겹쳐 그리지 않고 광역
단계구분도만 기본 표시하며, 시도명·약칭·세부지역명을 전국에서 검색한다.

EPSIS 원본에서 광역지역과 세부지역 표기가 충돌하는 범위는 51행·53.289725 MW다. 공식 근거를
검토해 확정한 47행·48.895470 MW만 raw-row override를 적용한다. 미해결 4행·4.394255 MW는 원본
광역지역 합계를 유지하되 표준 세부지역 집계에서 `미확정 지역`으로 격리한다. Bronze 원문과 전국
합계 188,594행·33,059.516180 MW는 이 처리 전후에 동일하다. 기준정보는
[`config/national_solar_location_overrides_20260828.json`](../config/national_solar_location_overrides_20260828.json),
상세 근거와 승인 조건은 `NATIONAL_REGION_REFERENCE_AUDIT.md`에 기록한다.

과거 코드는 `발전기명 count`를 `발전소수`로 표기했지만, 공식 물리 발전소 ID가 없어 의미가
확정되지 않는다. 새 화면은 해당 수치를 `태양광 설비 등록 레코드`로 표기하고, 정확히
같은 557개 레코드도 공식 고유키가 없으므로 임의 삭제하지 않는다. 원본 합계행은 행 수·용량
교차검증에만 사용하고 지역 집계에서는 제외해 용량 2배 집계를 방지한다.

세 HTML은 동일 CSS/JavaScript와 생성 JSON을 공유하지만 역할은 섞지 않는다. 메인 화면의
`national_inventory`는 EPSIS 전국 원본의 사용자용 집계, `forecast`는 정식 Test 구간의 실제값과
선택 모델 예측 시계열, `model_analysis`는 동일 Test 표본의 모델 성능과 태양광 데이터 이상 신호만
담는다. `forecast`는 예보 발행시각 기상 입력이 연결되기 전까지 미래 운영예측으로 표현하지 않는다.
시도 경계는 로컬에서 제공하고 저채도 OSM
배경은 위치 문맥에만 사용한다. 발전소 점 마커는 사용하지 않는다.

사용자 JSON과 화면에는 운영자용 원천 해시·로컬 경로·인코딩·스키마 검사명을 노출하지 않는다.
해당 정보는 내부 원천 config와 이 감사 문서에 유지한다. 전국 현황 지도는 OSM 라벨을 행정구역
식별 근거로 사용하지 않으며, 로컬 경계와 표준화된 공식 한글 지역명이 통계·선택의 기준이다.

전국 원본은 SHA-256과 필수 11개 컬럼을 확인한 뒤 각 데이터 행의 `발전원=태양에너지`도
검증한다. 이후 다운로드에서 조회 필터가 풀려 다른 발전원이 섞이면 부분 집계를 게시하지 않고
즉시 실패한다. 검토된 raw-row override와 미해결 격리 규칙도 이 원본 SHA-256에 결합하며, 그룹별
기대 행 수·설비용량 합계와 근거 URL·판정 방법·신뢰도·사유를 함께 검증한다. SHA 또는 기대 집계가
다르면 일부 override나 이전 산출물을 적용하지 않고 전체 생성을 실패시킨다. 시도 경계 파일도
config의 `boundary_path`를 실제 게시 경로에 반영한다.

```powershell
python app.py build-dashboard
python app.py serve-dashboard --port 5500
```

이 전용 명령은 `dashboard/`를 서버 문서 루트로 고정한다. 프로젝트 루트에서 실행되는 Live Server를
위해서는 `/solar_dashboard.html`, 과거 북마크를 위해서는 `/map/html/solar_dashboard.html` 호환
진입점을 제공하며 둘 다 현재 `/dashboard/solar_dashboard.html`로 이동한다.
사용자 지정 `--output-dir`에는 정적 HTML/CSS/JavaScript와 생성 JSON/GeoJSON을 모두 복사한다.

대량 데이터 측면에서는 전국 CSV를 행 단위로 스트리밍하여 시도·세부지역 집계만 메모리에
유지한다. 정확 중복 비교도 임시 SQLite에 내려 원본 행 전체를 RAM에 적재하지 않는다.
모델 분석도 두 예측 CSV를 100,000행 chunk로 임시 SQLite에 적재해 공통 Test 키를 디스크에서
정렬한다. 지표 accumulator와 전체 이상 신호 카운터만 누적하고, 상세 이벤트는 모델별 초과비율
상위 250건만 heap에 유지한다. 선 그래프는 전역 무작위 표본이 아니라 발전소별 최근 연속
168시간을 bounded deque로 유지한다. NMAE와 이상 기준은 설비용량 적용 가능 표본을 명시하고,
Calibration 발전소별 표본이 168개 미만이면 용량 정규화 전역 기준으로 안전하게 후퇴한다.
`원본 해시 검증 → footer 교차검증 → SHA-bound override·격리 계약 검증 → 유효일 지역명 표준화 →
지역 집계 → 대시보드 JSON 원자적 교체`로 lineage를 남긴다. `config/national_solar_inventory.json`은
원천 경로, 제공기관, 기준일, 다운로드 시각, 인코딩, SHA-256, 범위 제약을 보존한다.
