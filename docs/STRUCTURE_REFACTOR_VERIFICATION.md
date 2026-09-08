# 코드 구조 정비와 ASOS 연결 검증

이 변경은 KOSPO identity 교정 PR #1의 commit `477b01e9e4944330b7b17695b31d40c800819bf3` 위에서 작성했다. 마이크로서비스 전환은 하지 않았다. 전체 Python 모듈의 소유 영역·이름, CLI 명령 파일, 모델별 구현, 화면 JavaScript 원본을 정리하고 ASOS API 수집 경계를 추가했다.

## 수행한 검증

| 확인 | 결과와 범위 |
| --- | --- |
| 기존·추가 회귀 검증 | 실행 가능한 범위에서 **170 passed, 4 deselected, 18 subtests passed** |
| KOSPO 회귀 | 위 170개에 포함된 14개. identity·admission·registry·합성 Gold 연결 |
| ASOS API 회귀 | 위 170개에 포함된 19개. 인증키 인코딩·페이지·호출 한도·캐시 무결성·부분 실패·연도 병합·잠금·합성 Gold 연결 |
| `tools/check_structure.py` | 100개 Python 모듈의 목적·명명·내부 import와 심볼·소스/산출물 경계·색인·JS bundle 일치 통과 |
| Python 구문 검사 | `python -m compileall -q src tests tools` 통과 |
| JS 구문 검사 | `node --check dashboard/assets/dashboard.js` 통과 |
| 이전 코드와 AST 대조 | 이름·import·docstring 차이를 정규화한 클래스/함수 정의 682개가 동일. 이전 정의의 누락 없음. 차이 19개는 ASOS 연결·CLI 옵션·기본 경로 상수·수집 factory/병합 변경 범위 |
| 대시보드 원본/결과 대조 | Node VM에서 변경 전후 11개 HTML 문자열·예측 시계열 정규화 결과 동일. 전국 현황·예측·비교·성능·이상치 화면 포함 |
| CLI | 도움말, `jobs`, ASOS collect/verify-e2e 옵션 파싱 통과. 도움말은 학습 프레임워크를 import하지 않음 |

Python AST와 Node VM 대조는 작업 직전 snapshot을 기준으로 수행한 일회성 대조다. 모델을 실제로 학습하거나 브라우저·Leaflet 상호작용을 실행한 검증으로 해석하지 않는다. 프로젝트의 계속 사용하는 검증 도구는 `tools/check_structure.py`, `tools/update_code_index.py`, `tools/build_dashboard_assets.py`와 `tests/`다.

## 전체 실행을 하지 못한 항목

현재 실행 환경에는 PyTorch·Optuna·XGBoost·Selenium이 없다. pytest 9.1.1은 사용할 수 있었고 HTTP adapter import에 필요한 실제 requests는 실행 환경 pip의 vendored 패키지를 `sys.path` 끝에 추가해 사용했다. 제품 코드나 계산을 가짜 모듈로 대체하지 않았다. 일반 개발 환경에서는 프로젝트 의존성을 설치한 뒤 `python -m pytest`를 실행한다.

학습 의존성 때문에 이번 실행에서 제외한 테스트 파일:

- `test_cnn_bilstm_workflow.py`
- `test_ensemble.py`
- `test_forecast_pipeline.py`
- `test_model_checkpoints.py`
- `test_model_optimization.py`

가져온 코드 snapshot에 대용량 공식 원본이 없어 제외한 검증 4개:

- `test_actual_dasong_export_reconciles_both_daily_total_unit_segments`
- `test_production_location_review_contract_preserves_totals_and_nesting`
- `test_official_sgis_boundaries_have_complete_unique_regions_and_island_owners`
- `test_boundary_validation_rejects_duplicate_region_ids`

해당 테스트를 삭제하거나 상시 skip 처리하지 않았다. 원래 의존성과 데이터를 보유한 개발 폴더에서 그대로 실행된다. CI의 `Source structure` workflow는 가벼운 구조·구문 검사이며 전체 모델 회귀 통과를 대신하지 않는다.

실제 ASOS API 키는 주입하지 않았다. 라이브 API 성공, 실데이터 Gold 행 증가, 전체 `prepare-data` 또는 `verify-e2e` 완료를 주장하지 않는다. API에서 발전소의 기상 매핑 승인을 유추하지 않으며 기존 registry 게이트를 유지한다.

## 적용과 호환성

모델 ID·학습 설정 JSON 값·checkpoint payload·기존 Gold schema는 유지했다. 기존 source import 경로는 실제 재배치했으므로 별도 스크립트는 [module_moves.json](../config/architecture/module_moves.json)을 따라 수정한다. CLI 명령 이름과 패키지 `main`/`build_parser` 진입점은 유지한다.

새 실행의 기본 평가 결과 위치는 `artifacts/` 아래로 통일했다. 기존 `output/`, 원본, 모델 결과를 물리적으로 이동하거나 삭제하지 않았다. `C:\ESG_Project1`의 잠긴 빈 루트 폴더에 접근하거나 삭제한 변경도 아니다. 데이터·결과 폴더의 역할과 새 생성 규칙은 [PROJECT_STRUCTURE.md](PROJECT_STRUCTURE.md)를 따른다.
