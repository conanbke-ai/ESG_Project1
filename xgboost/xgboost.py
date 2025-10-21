import pandas as pd
import numpy as np
from sklearn.metrics import mean_squared_error
from xgboost import XGBRegressor
import matplotlib.pyplot as plt
from io import BytesIO
import base64
import webbrowser
import os
import logging

# ===== 로깅 설정 =====
logging.basicConfig(level=logging.INFO, format='[%(asctime)s]✅ %(message)s')

# ==== 1️⃣ 데이터 로드 & 전처리 ====
pv_file = "pv_data.csv"          
weather_file = "weather_data.csv"

logging.info("데이터 파일 불러오는 중...")
pv_df = pd.read_csv(pv_file)
weather_df = pd.read_csv(weather_file)

logging.info("데이터 통합 중...")
df = pd.merge(pv_df, weather_df,
              on=['광역지역', '세부지역', '발전소명', '날짜'], how='inner')
df['날짜'] = pd.to_datetime(df['날짜'], format="%Y-%m-%d %H:%M")
df = df.sort_values(['광역지역','세부지역','발전소명','날짜'])

df['hour'] = df['날짜'].dt.hour
df['dayofweek'] = df['날짜'].dt.dayofweek

cat_cols = ['광역지역','세부지역','발전소명']
for col in cat_cols:
    df[col] = df[col].astype('category').cat.codes

lag_hours = [1,2,3]
roll_windows = [3,6]

logging.info("지연(lag) 및 이동평균(rolling) 피처 생성 중...")
df_grouped = []
for _, group in df.groupby(['광역지역','세부지역','발전소명']):
    g = group.copy()
    for lag in lag_hours:
        g[f'발전량_lag_{lag}'] = g['발전량'].shift(lag)
    for w in roll_windows:
        g[f'발전량_roll_{w}'] = g['발전량'].shift(1).rolling(w).mean()
    df_grouped.append(g)

df = pd.concat(df_grouped)
df = df.dropna().reset_index(drop=True)

target = '발전량'
features = ['광역지역','세부지역','발전소명','기온','강수량','일조량','일사량','hour','dayofweek']
features += [f'발전량_lag_{lag}' for lag in lag_hours]
features += [f'발전량_roll_{w}' for w in roll_windows]

X = df[features]
y = df[target]

logging.info("데이터 학습/테스트 분리 중...")
train_idx, test_idx = [], []
for _, group in df.groupby(['광역지역','세부지역','발전소명']):
    n = len(group)
    split = int(n*0.8)
    train_idx.extend(group.index[:split])
    test_idx.extend(group.index[split:])

X_train, X_test = X.loc[train_idx], X.loc[test_idx]
y_train, y_test = y.loc[train_idx], y.loc[test_idx]

# ==== 2️⃣ XGBoost 학습 ====
logging.info("XGBoost 모델 학습 시작...")
xgb_model = XGBRegressor(
    n_estimators=1000,
    max_depth=5,
    learning_rate=0.05,
    subsample=0.8,
    colsample_bytree=0.8,
    reg_alpha=0.1,
    reg_lambda=1.0,
    random_state=42
)

xgb_model.fit(
    X_train, y_train,
    eval_set=[(X_test, y_test)],
    eval_metric='rmse',
    early_stopping_rounds=50,
    verbose=False
)
logging.info("모델 학습 완료")

# ==== 3️⃣ 예측 ====
logging.info("테스트 데이터 예측 중...")
df_test = df.loc[test_idx].copy()
df_test['pred'] = xgb_model.predict(X_test)

# ==== 4️⃣ 발전소별 RMSE 계산 ====
logging.info("발전소별 RMSE 계산 중...")
rmse_list = []
for (gname, g), group in df_test.groupby(['광역지역','세부지역','발전소명']):
    rmse = mean_squared_error(group['발전량'], group['pred'], squared=False)
    rmse_list.append([gname[0], gname[1], gname[2], rmse])

rmse_df = pd.DataFrame(rmse_list, columns=['광역지역','세부지역','발전소명','RMSE'])
rmse_df = rmse_df.sort_values('RMSE', ascending=False).reset_index(drop=True)
logging.info(f"RMSE 계산 완료, 상위 5개 발전소:\n{rmse_df.head()}")

# ==== 5️⃣ 시계열 그래프 생성 ====
def plot_timeseries(df_subset):
    plt.figure(figsize=(6,3))
    plt.plot(df_subset['날짜'], df_subset['발전량'], label='실제', marker='o')
    plt.plot(df_subset['날짜'], df_subset['pred'], label='예측', marker='x')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.legend()
    buf = BytesIO()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return base64.b64encode(buf.read()).decode('utf-8')

top10_stations = rmse_df.head(10)
plot_imgs = []
logging.info("상위 10개 발전소 시계열 그래프 생성 중...")
for _, row in top10_stations.iterrows():
    subset = df_test[(df_test['광역지역']==row['광역지역']) &
                     (df_test['세부지역']==row['세부지역']) &
                     (df_test['발전소명']==row['발전소명'])].copy()
    img_b64 = plot_timeseries(subset)
    plot_imgs.append((f"{row['광역지역']} {row['세부지역']} {row['발전소명']}", img_b64))

# ==== 6️⃣ Feature Importance ====
logging.info("Feature Importance 그래프 생성 중...")
plt.figure(figsize=(6,4))
importances = xgb_model.feature_importances_
indices = np.argsort(importances)[::-1]
plt.barh(range(len(indices)), importances[indices], align='center')
plt.yticks(range(len(indices)), [features[i] for i in indices])
plt.gca().invert_yaxis()
plt.title("Feature Importance")
buf = BytesIO()
plt.savefig(buf, format='png', bbox_inches='tight')
plt.close()
buf.seek(0)
feat_b64 = base64.b64encode(buf.read()).decode('utf-8')

# ==== 7️⃣ HTML 생성 ====
logging.info("HTML 대시보드 생성 중...")
html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>태양광 발전량 예측 대시보드</title>
<style>
body {{ font-family: 'Malgun Gothic','Segoe UI',sans-serif; margin:20px; }}
h2 {{ color:#2c3e50; }}
table {{ border-collapse: collapse; width: 100%; margin-bottom:20px; }}
th, td {{ border:1px solid #ddd; padding:8px; text-align:center; }}
th {{ background-color:#4CAF50; color:white; }}
tr:nth-child(even) {{ background-color:#f2f2f2; }}
img {{ width:100%; max-width:600px; margin-bottom:20px; }}
.scroll-container {{
    max-height: 800px; overflow-y: scroll; border:1px solid #ccc; padding:10px;
}}
</style>
</head>
<body>
<h2>1️⃣ 발전소별 RMSE</h2>
{rmse_df.to_html(index=False)}

<h2>2️⃣ 상위 10개 발전소 예측 vs 실제 시계열</h2>
<div class="scroll-container">
"""

for name, img_b64 in plot_imgs:
    html_content += f"<h3>{name}</h3><img src='data:image/png;base64,{img_b64}' />"

html_content += f"""
</div>
<h2>3️⃣ Feature Importance</h2>
<img src='data:image/png;base64,{feat_b64}' />
</body>
</html>
"""

# ==== 8️⃣ HTML 저장 및 브라우저 열기 ====
OUTPUT_HTML = "pv_xgb_dashboard.html"
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_content)

webbrowser.open('file://' + os.path.realpath(OUTPUT_HTML))
logging.info(f"대시보드 생성 완료: {OUTPUT_HTML}")
