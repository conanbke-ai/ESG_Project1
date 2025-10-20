신재생 에너지 프로젝트 - 날씨 데이터 기반 재생에너지 데이터 분석

2조 : 기상청 기반 데이터 / 재생에너지 발전량 활용 예측 프로그램

* 환경
  파이썬 버전 : python 3.11
  
* 라이브러리 설치
  python -m pip install --upgrade pip
  pip install pandas folium requests tqdm matplotlib branca openpyxl plotly

<<<<<<< HEAD
=======
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

>>>>>>> dcf78bd013fbaf7b782254dd74d834912e313889
* 출처
  - 지도 경계 데이터 : https://simplemaps.com/gis/country/kr?utm_source=chatgpt.com
