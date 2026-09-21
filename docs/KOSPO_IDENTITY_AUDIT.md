# KOSPO collector 식별자 보강

검토일: 2026-09-08. 기준 main: `820ded3b2f45f9bc7a78412dd81444b730c06c67`.

## 판정과 범위

- `ALREADY_DONE`: collector Silver → admission → registry → Gold 연결(`66837859`),
  모듈러 모놀리스와 versioned job 계약(`2a07f466`).
- `NEW`: KOSPO 조직명과 개별 발전기명을 구별하는 식별자 보강.
- `BLOCKED`: 부산철도 2의 실제 Gold 편입은 공식 좌표 또는 근거 있는 ASOS 검토 매핑이 필요함.
- Windows의 미추적 빈 폴더 `C:\ESG_Project1\xgboost`, `C:\ESG_Project1\cnn_bilstm`은
  이번 실행 환경에서 접근할 수 없어 삭제하지 않음.

## 원본 근거

[공공데이터포털 15156688](https://www.data.go.kr/data/15156688/fileData.do)의 설명은
`발전소명=신재생사업본부`, `발전기명=부산 철도태양광 #2`, `계량구분=KPX`를 구분한다.
조직명만으로 alias를 만들면 같은 본부의 서로 다른 발전기를 합칠 수 있다.

저장소의 `file/solar_data_file/location/한국남부발전(주)_태양광발전기 사양정보_20250630.csv`
(기준 main의 Git blob `5bcad19b59e4588053a9250821289d571565368b`)에는
`부산철도 1`, `부산철도 2`, `부산철도 4`가 별도 행으로 있다. 2호기 주소는
`부산광역시 부산진구 신천대로 215`이고 설치용량·설치각은 공란이다.

`file/KMA_data_file/META_관측지점정보.csv`의 ASOS 159(부산)는 중구 대청동1가에 있다.
따라서 설비 식별자 보정만으로 기존의 구 단위 자동 기상 매핑을 통과하지 않는다.
이번 변경은 기상 매핑 설정을 추가하거나 설비용량·좌표를 추정하지 않는다.

## 변경 계약

`collectors/identity.py`의 `kospo.busan-rail-2.v1`은 회사·본부·발전기명이 모두 일치한
태양광 행만 `부산철도 2`로 식별한다. 발전기명의 공백 차이는 허용하되 #1/#4/#20이나
다른 본부에는 적용하지 않는다. 원래 발전기명은 `unit`에 보존하고, `plant_id`에서는
공백 차이를 정규화한 발전기명과 계량구분을 보존한다. 충돌하는 ID와 같은 파일 안에서
보정 후 발생하는 중복 시간 키는 오류로 처리한다.

- 신규 daily-wide 정규화에서 규칙 적용.
- 기존 Silver의 admission에서 적용 규칙·행 수·근거 URL을 `identity_resolutions`에 기록.
  기존 source 경로·SHA-256은 유지하며 파일을 덮어쓰지 않는다.
- registry와 Gold는 동일한 읽기 함수를 사용한다. `unit` 없는 과거 파티션은 기존처럼 처리한다.
- 기존 snapshot 최신본 선택과 기상·품질 게이트는 유지한다. 미지의 발전기명을 임의 매칭하지 않는다.
- admission manifest의 새 필드는 추가적인 감사 정보이며, generation/job 계약 버전은 유지한다.

## 검증

`tests/test_kospo_identity.py`의 표적 회귀검증 14개 통과(Python 3.12.13, pandas 2.2.3):

- 공식 컬럼 구조를 따른 합성 입력의 시간·kWh→MWh 변환, 원본·계량구분 보존.
- 다른 회사·본부·발전원·호기를 매칭하지 않음, 공백 변형과 재실행의 안정성.
- 충돌 ID 및 동일 계량기 중복 거부, 필수 컬럼 누락 거부, 과거 파티션 호환성.
- 기존 Silver hash 유지, admission 감사 기록, registry와 새·옛 snapshot 중복 제거.
- 공식 사양표에 근거한 부산진구 주소는 기상 검토 없이 격리됨, 용량 공란 유지.
- 테스트 전용 명시적 기상 매핑을 주입한 합성 admission → registry → Gold 파일·파티션·manifest 생성.

실행 환경에는 pytest/XGBoost/Torch가 설치되어 있지 않고 패키지 설치 요청도 승인 취소로
실행되지 않았다. 표적 검증은 표준 `unittest`로 실행했으며, 무관한 모델 의존성을 즉시 불러오는
`collectors/__init__.py`의 편의 export만 namespace package로 건너뛰었다.
HTTP를 호출하지 않는 import에는 설치된 pip의 vendored requests를 사용했다.
정규화·admission·metadata·registry·Gold·기상/품질 모듈은 실제 소스를 실행했으며 계산을
대체하지 않았다. 양성 Gold 테스트의 기상 검토는 테스트 전용 주입이며 운영 승인을 뜻하지 않는다.

전체 pytest, 실제 보관 데이터의 `prepare-data`, `verify-e2e` 및 정식 모델 학습은 재실행하지 않았다.
기존 718,531 Gold 행과 671,731 태양광 학습 적격 행은 이번 변경 이후 재검증한 수치가 아니다.
합성 테스트의 24행을 실제 편입 증가분으로 보고하지 않는다.

프로젝트 의존성이 설치된 환경에서 후속 검증:

```bash
python -m pytest tests/test_kospo_identity.py tests/test_collector_admission.py tests/test_collectors.py tests/test_plant_registry.py
python app.py prepare-data
```

먼저 registry의 부산철도 2 주소·격리 사유를 확인한다. 공식 좌표 또는 검토된 기상 매핑을
확보한 뒤에만 전체 Gold source contribution과 학습 적격 행 수를 다시 산출한다.
설치용량은 공식 자료로 별도 보강할 때까지 결측으로 보존한다.
