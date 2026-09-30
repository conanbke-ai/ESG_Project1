# ASOS 시간자료 API 설정

대상은 공공데이터포털의 [기상청_지상(종관, ASOS) 시간자료 조회서비스](https://www.data.go.kr/data/15057210/openapi.do), 데이터 ID `15057210`이다. 공식 요청 주소는 `https://apis.data.go.kr/1360000/AsosHourlyInfoService/getWthrDataList`이며 관측소 번호와 시작·종료 날짜/시간을 받는다. 이 서비스는 전일(D-1)까지의 관측자료를 제공한다.

포털에서 안내한 `https://apis.data.go.kr/1360000/AsosHourlyInfoService`는 올바른 **서비스 endpoint**다. 코드의 `collectors/kma_api_client.py`에서 `ASOS_SERVICE_ENDPOINT`로 정의하고, 시간자료 조회 메서드 `/getWthrDataList`를 붙여 `ASOS_HOURLY_ENDPOINT`를 만든다. 사용자가 endpoint를 환경변수에 추가할 필요는 없다. 키를 이미 입력했다면 아래 변수명만 일치하면 된다.

## 키를 입력할 파일과 변수

프로젝트 루트의 **`C:\ESG_Project1\.env.local`**에 아래 항목을 추가한다. 파일이 이미 있으면 다른 설정을 보존하고 이 항목만 편집한다.

```dotenv
KMA_ASOS_SERVICE_KEY=발급받은_일반인증키_Decoding
```

| 위치 / 이름 | 역할 |
| --- | --- |
| `.env.local` / `KMA_ASOS_SERVICE_KEY` | 실제 ASOS 인증키. Git에서 제외된다. |
| `.env.example` / `KMA_ASOS_SERVICE_KEY=` | 공유 가능한 빈 설정 예시. |
| `infrastructure/environment.py` / `load_local_env()` | 로컬 설정을 환경변수에 적용한다. 이미 설정된 프로세스 환경변수가 우선한다. |
| `collectors/kma_api_client.py` / `ASOS_SERVICE_KEY_ENV` | 환경변수 이름 상수 `KMA_ASOS_SERVICE_KEY`. |
| `collectors/kma_api.py` / `KmaAsosApiCollector.collect()` | 환경변수 값을 읽어 ASOS client에 전달한다. |
| `collectors/kma_api_client.py` / `AsosHourlyApiClient.fetch_page()` | 키를 `serviceKey` 쿼리로 한 번 URL 인코딩하여 전송한다. |
| `collectors/komipo_api.py` / `DATA_GO_SERVICE_KEY` | 기존 중부발전 수집 설정. ASOS 설정과 별도로 유지한다. |

이 설정은 ASOS API 수집기가 포함된 변경을 적용한 뒤 동작한다. API 키를 소스에 직접 넣거나 채팅에 전달할 필요가 없다. 이 변경에서 실제 키를 추가하거나 승인 상태를 확인한 것은 아니다.

## 실행

```powershell
python app.py collect --sources kma --kma-mode api --station-ids 108,159 --start-date 2026-01-01 --end-date 2026-01-02 --api-max-calls 10
python app.py prepare-data
```

`--station-ids`에는 수집할 공식 ASOS 지점 번호를 지정한다. 발전소에 사용할 관측소의 승인 여부는 기존 registry와 `config/reviewed_weather_mappings.json`에서 별도로 검증한다. 예시 지점 선택이 특정 발전소의 기상 매핑 승인으로 바뀌지는 않는다.

- `--kma-mode api`: 키나 관측소 목록이 없으면 설정 필요 상태로 종료한다.
- `--kma-mode auto`: 키가 있으면 API를 선택하고, 키가 없으면 기존 브라우저 수집을 사용한다. API 오류가 나면 실패를 보고하며 브라우저 호출로 숨기지 않는다.
- `--kma-mode browser`: 기존 포털 방식. `KMA_CHROME_USER_DATA_DIR`는 이 모드의 로그인 프로필 설정이다.
- `--weather-root`: 연도별 기상 CSV와 관측소 메타데이터 위치. 기본 `file/KMA_data_file`.
- `--api-max-calls`: 한 수집 실행의 호출 상한. 실제 계정의 승인 트래픽과는 별개이며 자동 재시도는 하지 않는다.

`verify-e2e --collect`에도 `--station-ids`, `--kma-mode`, `--weather-root`가 전달된다. API client는 한국 시간의 전일까지만 조회한다. 제공기관이 아직 관측값을 게시하지 않은 구간은 다음 실행에서 다시 조회할 수 있다.

## 저장과 Gold 연결

| 단계 | 파일 / 처리 |
| --- | --- |
| Bronze | `file/raw/kma/asos_hourly/station_<id>/<start>_<end>/responses/<snapshot>/page_<n>.json`에 성공 응답 원문 byte 보존 |
| 수집 checkpoint | 같은 기간 폴더의 `partition_manifest.json`: 요청 범위, 페이지 hash, 반환 행 수, 기대/관측 시간 수 |
| 기상 Silver | `file/KMA_data_file/OBS_ASOS_TIM_<year>.csv`: 관측소·시간 키로 API/브라우저 자료를 병합 |
| 발전량 admission | `datasets/collector_admission.py`가 발전량 Silver를 심사하고 registry에 전달 |
| Gold | `datasets/model_dataset_builder.py`가 승인된 발전소와 기상 관측값·과거 특징을 결합 |

모든 페이지를 확인하기 전에는 해당 기간을 완료로 기록하거나 연도별 CSV에 부분 게시하지 않는다. 성공 응답 전체를 받았다는 의미와 기대 시간 수를 모두 관측했다는 의미를 구분한다. 빈 응답은 완료 캐시로 삼지 않는다. 관측 누락이 있는 기간은 재조회하고, 모든 시간이 있는 기간만 hash 검증 후 재사용한다. 재수집 `--overwrite`는 이전 응답 snapshot을 보존하고 새로운 snapshot을 만든다.

연도별 병합은 다른 관측소·기존 추가 열을 유지한다. API에서 아예 빠진 필드는 이전 관측값을 지우지 않는다. 명시적 결측과 QC 비정상 값은 0으로 위조하지 않으며 QC 필드도 CSV에 보존한다. ASOS 필드는 기존 `KmaAsosNormalizer`의 열 계약으로 이어진다.

API와 브라우저 수집은 같은 기상 폴더의 `.asos_collection.lock`으로 쓰기를 직렬화한다. 비정상 프로세스 종료 후 잠금이 남으면 기록된 PID의 프로세스 종료를 먼저 확인하고 잠금만 제거한다. 정상 완료·예외 종료에서는 잠금이 자동 해제된다. 기본 `.env.local`과 잠금·임시 CSV 파일은 Git에서 제외된다.

ASOS는 관측자료다. 이번 연결이 실제 예보 발행시각을 가진 기상 예측 데이터나 운영 예측 성능 증명을 생성하지는 않는다. 기존 시간 분할·관측 업데이트 기반 평가 계약과 발전소별 기상 승인 게이트를 유지한다.
