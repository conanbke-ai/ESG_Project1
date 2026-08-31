# 수동 보정·캐시의 공식 자료 대체 감사

검증일: 2026-08-31 (Asia/Seoul)

## 결론

수동으로 맞춘 값을 모두 같은 방식으로 없애면 안 된다. 행정코드·경계·관측소 목록처럼 공식
정답이 있는 항목은 스냅샷과 SHA-256으로 자동 재생성하고, 발전소명 통합이나 발전소별 ASOS 선택처럼
공식 공통 식별자가 없는 항목은 공식 자료로 후보를 만든 뒤 유일 일치만 자동 승인한다. 애매한 후보는
근거 URL·거리·승인 사유를 가진 reviewed 예외로 남긴다.

이번 작업에서 전국 시도 경계는 기존 SimpleMaps 파일 대신
[국가데이터처 SGIS 행정구역 경계 2025년 2분기 자료](https://www.data.go.kr/data/15129688/fileData.do)로
교체했다. 원본 archive와 시도 Shapefile의 SHA-256, 기준일, 좌표계, 단순화 오차를 GeoJSON에
포함하고 백령도·대청도·연평도는 인천, 울릉도는 경북에 단독 포함되는지 게시 전 검사한다.
재변환 명령은 archive 자체의 hash를 먼저 검증하고 실제 `.shp/.shx/.dbf/.prj`가 그 archive의
동명 구성요소와 같은지, `BASE_DATE`가 선언한 기준일과 같은지까지 확인한다.

## 가장 먼저 고쳐야 할 경로

`file/merge_data/val.csv`는 공식 발전량 target으로는 더 이상 사용하지 않지만 `prepare-data`에서
발전소별 ASOS 지점 seed로 계속 읽힌다. 이 seed는 과거 `managed_human` 별칭, Kakao 주소 지오코딩,
최근접 관측소 계산에서 만들어졌다.

- 직접 `reviewed_legacy_mapping`: 19개 자산
- 같은 주소로 전파한 `reviewed_colocated_address`: 3개 자산
- 합계 영향: 22개 자산, 표준화 관측 2,400,964행
- 기존 Gold 영향: 1,810,752행(71.70%), 태양광 1,704,144행(70.97%)
- 현재 문제: 매핑별 공식 근거 URL·원본 hash·거리 제한이 없어도 `high`, 검토 불필요로 승격됨
- 특히 부산운동장 태양광과 행원소수력은 주소도 비어 있는데 legacy 지점번호만으로 승인됨

코드는 legacy seed를 학습 가능 근거가 아니라 audit-only 후보로 바꾸고 동일주소 자동 전파를
제거했다. 공식 발전소 주소·좌표와 발전기간 전체를 덮는 KMA 지점 메타데이터로 다시 계산한 뒤,
유일한 행정구역 일치 또는 검증된 거리 조건을 통과한 것만 승인한다. 나머지는
`config/reviewed_weather_mappings.json`에 근거 URL·거리·사유를 기록한 예외로 옮긴 후
`val.csv` 의존을 제거한다.

안전장치 적용 후 `prepare-data`를 전체 재실행한 결과 registry 68개 자산 중 학습 가능 24개,
격리 44개가 됐다. 태양광은 22개 자산·701,011행, 풍력은 원본 보존용 2개 자산·17,520행이며
태양광 모델은 여전히 `energy_source=solar`만 선택한다. 새 모델용 전체 행은 718,531개다.
legacy 방식으로 학습에 들어간 행, `review_required=true`인 행, explicit reviewed mapping의 근거
URL·사유 누락은 모두 0건이다. legacy 후보 19개 중 영월철도부지 1개만 공식 주소와 발전기간에
유효한 영월 ASOS가 독립적으로 일치해 편입됐으며, legacy 지점번호 자체는 승인 근거로 쓰지 않았다.

## 파일·규칙별 대체 가능성

| 현재 항목 | 영향 | 공식 자료 대체 판단 | 처리 방침 |
| --- | --- | --- | --- |
| `config/administrative_regions_20260701.json`의 코드·명칭 | 전국 집계·검색 | 자동 가능 | 행정표준코드 변경이력으로 유효기간 dimension 생성 |
| `map/json/geoJson.json` | 전국 단계구분도 | 자동 완료 | SGIS archive를 WGS84로 변환하고 source hash·도서 회귀검사 유지 |
| `map/json/coord_cache.json` 687키 | 세부지역 표시용 대표점 | 자동 가능 | 공식 시군구 geometry 대표점으로 교체; 발전소 실좌표로 사용 금지 |
| `file/solar_data_file/location/*.csv` 4개 | 용량·각도·주소 보강 | 자동 가능 | 발전사/공공데이터포털 첨부 collector와 metadata manifest 추가 |
| `config/data_scope.json`의 건수 | 문서용 현황 | 자동 가능 | generation/model/quality manifest에서 build 시 생성 |
| `PlantMetadataCatalog.aliases` 19개 | snapshot 간 발전소 identity 통합 | 후보 자동·승인 검토 | 주소·용량·호기·허가번호가 유일 일치할 때만 자동, 나머지는 provenance crosswalk |
| `config/reviewed_weather_mappings.json` 9개 | Gold ASOS 결합 | 최종 검토 필요 | 공식 좌표와 KMA 지점 후보·거리는 자동 계산하고 지형 판단만 reviewed 예외 유지 |
| 영암 용량 3값·군 단위 주소 | 물리 상한·기상 결합 | 자동+검토 | 공식 설비 metadata snapshot으로 용량·주소 보강; 관측소 선택은 별도 검토 |
| KOEN/EWP/KOSPO 단위 예외·위경도 swap | 1,000배 단위 오류 가능 | source별 검토 | 일반 임계값 대신 source SHA/schema version 예외와 물리 상한 검증으로 제한 |
| `managed_human.json` 64쌍 | 현재 consumer 없음 | 제거 가능 | 필요한 항목만 provenance crosswalk로 이관 후 legacy 보관 또는 삭제 |
| `cache_latlon.json`, `cache_location.json` | 현재 consumer 없음 | 제거 가능 | 공식 metadata+KMA registry로 재생성하고 과거 Kakao 산출물 제거 |
| `excluded_bad_plants.csv`, 빈 `val_gap.csv` | 현재 consumer 없음 | 제거 가능 | 현재 quality flag와 시간순 split manifest가 역할 대체 |
| `candidate_sources.json` | 연구 intake 목록 | 일부 자동 | 포털 메타·수정일·첨부 hash는 자동, 학습 승인 판단은 검토 상태로 분리 |

`coord_cache.json`에는 null 좌표 1개와 `NaN`, `nan nan` 오염 키가 있으며 출처일·hash가 없다.
`cache_latlon.json`과 `cache_location.json`도 현재 코드에서 사용하지 않고, 발전소명 중복과 좌표/지점
결측을 포함한다. 이 캐시는 행정구역 또는 발전소 위치의 정답으로 승격하지 않는다.

## 공식 수집원과 역할

1. [행정안전부 행정표준코드](https://www.code.go.kr/stdcode/regCodeL.do)와
   [법정동 전체자료](https://www.data.go.kr/data/15063424/fileData.do): 코드·명칭·시행일·폐지일.
2. [SGIS 행정구역 통계 및 경계](https://www.data.go.kr/data/15129688/fileData.do) 또는
   [국토교통부 행정구역도 API](https://www.data.go.kr/data/15059008/openapi.do): 시도·시군구
   geometry와 point-in-polygon.
3. [전국태양광발전소전기사업허가정보 표준데이터](https://www.data.go.kr/data/15107742/standard.do):
   시설명·주소·좌표·용량·허가일·운영상태. EPSIS 공통 ID가 없으므로 exact unique match만 승인.
4. [기상청 ASOS 시간자료 API](https://www.data.go.kr/tcs/dss/selectApiDataDetailView.do?publicDataPk=15057210)와
   [기상관측 지점정보 API](https://www.data.go.kr/data/15139439/openapi.do): 시간 기상과
   관측소 위치·이력. 지점 이전을 고려해 `station_id + valid_from + valid_to`로 결합.
5. [한국서부발전 AI 학습용 발전실적](https://www.data.go.kr/data/15155017/fileData.do): 발전기,
   일시, 발전량, 용량, 주소가 함께 있는 우선 수집 후보.
6. [한국중부발전 신재생 발전량 API](https://www.data.go.kr/data/15084511/openapi.do)와
   [한국지역난방공사 인버터 API](https://www.data.go.kr/data/15157890/openapi.do): API key로
   증분 Bronze 수집 가능. 단위·과거 조회범위·수정행을 먼저 검증.
7. [전력거래소 지역별 시간별 태양광 집계](https://www.data.go.kr/data/15103243/openapi.do):
   지역 합계 검증용. 개별 발전소 행이나 가짜 `plant_id`로 복제하지 않는다.

## 자동 매칭 승인 계약

공식 자료를 다운로드했다는 사실만으로 기존 행을 덮어쓰지 않는다. 다음을 모두 통과할 때만 자동
승인한다.

1. 원본 snapshot URL·기준일·byte·SHA-256이 기록돼 있다.
2. 공유 공식 ID가 있거나, 정규화 이름+회사+호기+용량+주소 조합이 전국에서 한 건으로 유일하다.
3. 주소 행정코드와 좌표 point-in-polygon 결과가 일치한다.
4. 관측소는 발전소 좌표 시점에 유효하며, 거리와 후보 수가 기록된다.
5. 단위 변환 후 `generation_mwh <= capacity_mw × interval_hours × tolerance` 물리 검사를 통과한다.
6. 매칭 전후 발전량 행 수와 총량이 같고, 바뀐 identity·지역·관측소가 별도 diff로 남는다.

부분 문자열 fuzzy match, 시군구 중심점을 발전소 좌표로 사용, 지역 집계값을 발전소 행으로 복제,
EPSIS 발전기 등록행을 물리 발전소 개소로 간주하는 처리는 금지한다.

## 실행 순서

- 완료: legacy ASOS seed를 audit-only로 낮추고 stale registry의 unsafe method도 downstream에서 차단했다.
- P0: 공식 발전소 metadata collector+manifest를 만든다.
- P0: 22개 legacy 매핑을 공식 주소·좌표와 KMA 지점 이력으로 재산출하고, 승인된 예외만 reviewed
  config에 이관한 뒤 `val.csv` seed를 제거한다.
- P1: 발전소 alias와 source별 단위 예외를 코드 dict에서 source-versioned provenance table로 옮긴다.
- P1: 공식 시군구 geometry로 표시 대표점을 재생성하고 Kakao cache를 제거한다.
- P2: consumer가 없는 legacy JSON/CSV를 제거하고 현황·후보 메타데이터를 manifest에서 자동 생성한다.
