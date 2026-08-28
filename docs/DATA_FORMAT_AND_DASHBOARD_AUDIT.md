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

`solar_dashboard.html`은 학습 데이터 보유 여부와 무관한 전국 태양광 발전설비 현황으로 복원했다.
2026-08-28에 EPSIS 발전기 현황 화면에서 `발전원=태양에너지`로 직접 재수집한 CP949 CSV를
Bronze 원본으로 보존한다. 공식 기준일은 2026-08-05이며 합계행 제외 후 188,594행,
33,059.516180 MW, 17개 시도이다. 세부지역 320개 원본 표기는 구·신 행정구역명을
표준화해 272개 지역으로 통합한다. 메인 지도에는 272개 점을 겹쳐 그리지 않고 17개 시도
단계구분도만 기본 표시하며, 클릭한 시도의 상위 세부지역을 팝업으로 제공한다.

과거 코드는 `발전기명 count`를 `발전소수`로 표기했지만, 공식 물리 발전소 ID가 없어 의미가
확정되지 않는다. 새 화면은 해당 수치를 `태양광 설비 등록 레코드`로 표기하고, 정확히
같은 557개 레코드도 공식 고유키가 없으므로 임의 삭제하지 않는다. 원본 합계행은 행 수·용량
교차검증에만 사용하고 지역 집계에서는 제외해 용량 2배 집계를 방지한다.

두 HTML은 동일 CSS/JavaScript와 생성 JSON을 공유하지만 모집단은 섞지 않는다. 메인 화면의
`national_inventory`는 EPSIS 전국 원본에서, 학습 품질 화면의 `plants/mapping`은 Silver/Gold
registry에서 만든다. 시도 경계와 좌표 캐시는 로컬에서 제공한다. Leaflet 라이브러리는 CDN을,
OSM 배경 타일은 위치 문맥이 필요한 학습 품질 지도에서만 네트워크를 사용한다.

메인 화면에서는 운영자용 원천 해시·스키마 품질 카드를 노출하지 않는다. 해당 정보는 생성 JSON과
이 감사 문서에 유지하고, 화면은 KPI·시도 단계구분도·지역 순위에 집중한다. 전국 현황 지도는
불필요한 도로·해외 지명을 없애기 위해 OSM 배경 타일 없이 로컬 시도 경계만 사용한다.

전국 원본은 SHA-256과 필수 11개 컬럼을 확인한 뒤 각 데이터 행의 `발전원=태양에너지`도
검증한다. 이후 다운로드에서 조회 필터가 풀려 다른 발전원이 섞이면 부분 집계를 게시하지 않고
즉시 실패한다. 시도 경계 파일도 config의 `boundary_path`를 실제 게시 경로에 반영한다.

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
Gold도 각 gzip 파티션에서 매핑 검증에 필요한 컬럼과 유일 키만 읽고 버린다.
`원본 해시 검증 → footer 교차검증 → 지역명 표준화 → 지역 집계 → 대시보드 JSON 원자적 교체`로
lineage를 남긴다. `config/national_solar_inventory.json`은 원천 경로, 제공기관, 기준일,
다운로드 시각, 인코딩, SHA-256, 범위 제약을 보존한다.
