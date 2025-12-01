신재생 에너지 프로젝트 - 날씨 데이터 기반 재생에너지 데이터 분석

2조 : 기상청 기반 데이터 / 재생에너지 발전량 활용 예측 프로그램

* 환경
  파이썬 버전 : python 3.11
  
* 라이브러리 설치
  python -m pip install --upgrade pip
  pip install pandas folium requests tqdm matplotlib branca openpyxl plotly torch scikit-learn optuna

* 용량 큰 파일 처리

  # 1. Git LFS 사용
  # Git LFS 설치
  git lfs install

  # HTML 파일 등록
  git lfs track "*.html"

  # 커밋
  git add .gitattributes
  git add 경로/파일명
  git commit -m "Add large HTML file"
  git push

  # 2. filter 사용
  # 문제 파일 제거
  git filter-branch --force --index-filter \
  "git rm --cached --ignore-unmatch map/generator_map.html" \
  --prune-empty --tag-name-filter cat -- --all

  # 캐시 클린업
  rm -rf .git/refs/original/
  git reflog expire --expire=now --all
  git gc --prune=now --aggressive

  # 강제 푸시
  git push origin main --force

* 출처
  - 지도 경계 데이터 : https://simplemaps.com/gis/country/kr?utm_source=chatgpt.com
  - 발전량 데이터 : https://www.koenergy.kr/kosep/gv/nf/dt/nfdt21/main.do
  - 기상청 데이터 : https://data.kma.go.kr/data/grnd/selectAsosRltmList.do
  - 카카오 API : https://dapi.kakao.com/v2/local/search/address.json
  - XGBoost 모델 : https://github.com/yun-ss97/solar_prediction/blob/main/XGBoost.ipynb
  - CNN_LSTM 모델 : https://github.com/muntasirhsn/CNN-LSTM-model-for-energy-usage-forecasting/blob/main/CNN_LSTM_univariate_multistep_output_github.ipynb

## CNN-BiLSTM 파이프라인 (Optuna + 강화학습)

`cnn_bilstm` 디렉터리에 학습/검증/이상치 분석을 분리한 파이썬 모듈을 추가했습니다.

```python
import pandas as pd
from cnn_bilstm.workflows import train_and_save, compare_checkpoints, evaluate_and_analyze

df = pd.read_csv("your_dataset.csv")

# 1) Optuna로 하이퍼파라미터 탐색 + 학습 및 저장
artifacts = train_and_save(
    df,
    target_column="target",
    feature_columns=[...],
    output_dir="test/checkpoints",  # 출력 기준 경로 (각 실행마다 날짜 디렉터리 생성)
    use_optuna=True,
    use_reinforcement=True,  # 밴딧으로 학습률 자동 탐색
)

# 저장된 아티팩트는 test/checkpoints/yyyymmdd_HHMMSS/ 아래에 정리됩니다.
print(artifacts["output_dir"])

# 2) 저장된 체크포인트 성능 비교
summary = compare_checkpoints("test/checkpoints", df, target_column="target")
summary.to_csv("cnn_bilstm/output/checkpoints/benchmark.csv", index=False)

# 3) 저장 모델 불러오기 + 이상치 탐지
analysis = evaluate_and_analyze(
    artifacts["checkpoint_path"],
    df,
    target_column="target",
    output_dir="test/analysis",  # 결과도 날짜 디렉터리로 저장
)
analysis["anomalies"].head()
```

- `optuna_search.py` : Optuna 스터디 실행 및 최적 하이퍼파라미터 탐색
- `reinforcement.py` : epsilon-greedy 밴딧으로 학습률을 조정하며 강화학습 적용
- `workflows.py` : 학습/저장, 불러오기 후 성능 비교, 이상치 탐지를 담당

### 테스트

`test` 디렉터리에서 기본 동작을 검증하려면 아래 명령으로 실행합니다.

```
python -m pytest test
```

이전 노트북 없이도 동일한 파이프라인을 스크립트로 재사용할 수 있습니다.
