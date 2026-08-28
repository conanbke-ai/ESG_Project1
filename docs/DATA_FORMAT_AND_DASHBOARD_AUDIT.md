# 파일 포맷·발전소 매핑·대시보드 구조 감사

검증일: 2026-08-28 (Asia/Seoul)

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

현재 registry는 68개 자산이며 태양광 65, 풍력 2, 소수력 1이다. 태양광 중 43개는 검토 가능한
ASOS 매핑이 있어 학습 가능하고, 22개는 주소·좌표·검토 근거가 부족해 격리한다. 풍력과 소수력은
원본 계약에 보존하지만 두 태양광 모델은 `energy_source=solar`만 읽는다.

매핑은 다음 세 필드를 서로 대체하지 않는다.

1. `admin_province/admin_city`: 공식 주소에서 얻은 행정구역
2. `latitude/longitude`: 공식 자료에 실제 발전소 좌표가 있을 때만 기록
3. `weather_station_id/name`: 기상 결합에 사용한 ASOS와 근거 방식

지도 표시는 실좌표 5개, ASOS 대리좌표 38개, 미매핑 22개다. 대리좌표에는
`location_basis=weather_station_proxy`를 붙여 발전소 위치처럼 보이지 않게 한다. Registry 자동
검사는 plant ID 유일성, 학습 가능 태양광의 ASOS 존재, 좌표 쌍 완전성, 대한민국 범위,
미해결 자산 격리를 모두 통과했다.

Gold 32개 회사×연도 파티션의 2,525,434행을 필요한 식별 컬럼만 파티션 단위로 읽어 registry와
대조했다. 46개 Gold 자산의 `plant_id`, 회사, 표준 발전소명, 발전원, 행정구역, ASOS 지점,
매핑 방법 불일치는 0건이다. 이 검사는 `python app.py build-dashboard` 때 다시 실행되고 결과가
`mapping.gold_consistency`에 저장된다.

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
`plant_region_report_perm.html`은 현재 registry를 읽는 화면으로 교체했다.

## HTML 구성 판단

`solar_dashboard.html`은 다시 기본 진입점으로 제공한다. 다만 과거 화면의 전국 인허가 설비
3만여 개 수치는 현재 저장소에 원천 파일과 갱신 파이프라인이 없어 최신 시간별 발전실적 통계와
결합하지 않았다. 화면은 정확하게 “공개 시간별 발전실적을 확보한 학습 포트폴리오”로 표시한다.

두 HTML은 동일 CSS/JavaScript와 생성 JSON을 공유한다. 전국 현황과 매핑·품질을 별도 URL로
유지해 기존 북마크를 깨지 않으면서, 통계 복사본이 서로 달라지는 문제를 없앴다. Leaflet 지도만
CDN/OSM 타일을 사용하며 데이터 표와 검증 결과는 로컬 JSON이다.

```powershell
python app.py build-dashboard
python -m http.server 5500 --directory dashboard
```

대량 데이터 측면에서는 dashboard build도 Gold 전체를 한 번에 메모리에 올리지 않는다. 각 gzip
파티션에서 매핑 검증에 필요한 컬럼과 유일 키만 읽고 버린다. 향후 전국 설비 inventory 원천을
다시 확보하면 별도 `facility_inventory` 계약과 갱신일을 추가하고, 발전실적 확보 자산과 전국
인허가 설비를 같은 숫자로 합치지 않는 것이 안전하다.
