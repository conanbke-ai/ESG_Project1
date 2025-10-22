import pandas as pd
import numpy as np
from xgboost import XGBRegressor
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages

# =======================
# 1️⃣ 데이터 불러오기
# =======================
df = pd.read_csv("solar_data.csv", parse_dates=["date"])
df = df.sort_values("date")
df['year'] = df['date'].dt.year

# =======================
# 2️⃣ Feature Engineering
# =======================
for lag in range(1, 4):
    df[f'lag_{lag}'] = df['generation'].shift(lag)
df = df.dropna()
features = ['lag_1','lag_2','lag_3','temperature','sunlight_hours','radiation']

# =======================
# 3️⃣ 학습용/테스트용 분리
# =======================
train_df = df[df['year'].isin([2022, 2023, 2024])]
test_df = df[df['year'] == 2025]

X_train, y_train = train_df[features], train_df['generation']
X_test, y_test = test_df[features], test_df['generation']

# =======================
# 4️⃣ XGBoost 학습
# =======================
model = XGBRegressor(n_estimators=200, learning_rate=0.1)
model.fit(X_train, y_train)

# =======================
# 5️⃣ 예측 및 잔차 계산
# =======================
y_pred = model.predict(X_test)
residual = y_test - y_pred
test_df = test_df.copy()
test_df['residual'] = residual

# =======================
# 6️⃣ 이상치 탐지 (3σ)
# =======================
threshold = 3 * np.std(residual)
test_df['outlier'] = np.abs(residual) > threshold
outliers = test_df[test_df['outlier']]
normal = test_df[~test_df['outlier']]

# =======================
# 7️⃣ Feature Importance
# =======================
importances = model.feature_importances_
feature_imp = pd.DataFrame({'feature': features, 'importance': importances}).sort_values(by='importance', ascending=False)

# =======================
# 8️⃣ 상관분석 (잔차 vs 기상 변수)
# =======================
corrs = residual.corr(pd.DataFrame({
    'temperature': test_df['temperature'],
    'sunlight_hours': test_df['sunlight_hours'],
    'radiation': test_df['radiation']
}))

# =======================
# 9️⃣ PDF 보고서 생성
# =======================
pdf_path = "Solar_Outlier_Report_2025_Enhanced.pdf"
with PdfPages(pdf_path) as pdf:
    
    # === 표 1: 이상치 요약 ===
    fig, ax = plt.subplots(figsize=(8,2))
    ax.axis('off')
    text = f"2025년도 이상치 분석 보고서 (Enhanced)\n총 데이터: {len(test_df)}, 이상치 수: {len(outliers)}"
    ax.text(0.5, 0.5, text, fontsize=14, ha='center', va='center')
    pdf.savefig()
    plt.close()

    # === 표 2: 이상치 데이터 ===
    fig, ax = plt.subplots(figsize=(12,4))
    ax.axis('off')
    tbl = outliers[['date','generation','residual','temperature','sunlight_hours','radiation']]
    ax.table(cellText=tbl.values, colLabels=tbl.columns, loc='center', cellLoc='center').auto_set_font_size(False)
    pdf.savefig()
    plt.close()
    
    # === 표 3: 이상치 vs 정상치 평균 비교 ===
    fig, ax = plt.subplots(figsize=(8,3))
    ax.axis('off')
    summary = ""
    for col in ['temperature','sunlight_hours','radiation']:
        summary += f"{col}: 정상={normal[col].mean():.2f}, 이상치={outliers[col].mean():.2f}\n"
    ax.text(0.5,0.5,summary, fontsize=12, ha='center', va='center')
    pdf.savefig()
    plt.close()

    # === 표 4: 잔차 vs 기상 변수 상관계수 ===
    fig, ax = plt.subplots(figsize=(8,3))
    ax.axis('off')
    corr_text = "Residual vs Weather Variables Correlation:\n" + \
                "\n".join([f"{col}: {residual.corr(test_df[col]):.3f}" for col in ['temperature','sunlight_hours','radiation']])
    ax.text(0.5,0.5,corr_text, fontsize=12, ha='center', va='center')
    pdf.savefig()
    plt.close()

    # === 시각화 1: 발전량 시계열 + 이상치 ===
    plt.figure(figsize=(12,5))
    plt.plot(test_df['date'], test_df['generation'], label='Actual', color='blue')
    plt.scatter(outliers['date'], outliers['generation'], color='red', label='Outlier', s=30)
    plt.xlabel('Date')
    plt.ylabel('Generation')
    plt.title('2025 Solar Generation & Outliers')
    plt.legend()
    pdf.savefig()
    plt.close()

    # === 시각화 2: 잔차 vs 기상 변수 ===
    plt.figure(figsize=(12,5))
    for col, color in zip(['temperature','sunlight_hours','radiation'], ['orange','green','purple']):
        sns.scatterplot(x=col, y='residual', data=test_df, label=col, color=color)
        # 상관계수 주석
        plt.text(x=test_df[col].max()*0.8, y=test_df['residual'].max()*0.8,
                 s=f"corr={residual.corr(test_df[col]):.2f}", fontsize=10, color=color)
    plt.xlabel('Weather Variables')
    plt.ylabel('Residuals')
    plt.title('Residuals vs Weather Variables (2025)')
    plt.legend()
    pdf.savefig()
    plt.close()

    # === 시각화 3: 이상치/정상치 기상 변수 분포 ===
    plt.figure(figsize=(12,5))
    sns.boxplot(data=pd.melt(test_df, id_vars='outlier', value_vars=['temperature','sunlight_hours','radiation']),
                x='variable', y='value', hue='outlier')
    plt.title('Weather Variables Distribution: Outliers vs Normal (2025)')
    pdf.savefig()
    plt.close()

    # === 표 5: Feature Importance ===
    fig, ax = plt.subplots(figsize=(8,3))
    ax.axis('off')
    imp_text = "Feature Importance:\n" + "\n".join([f"{row.feature}: {row.importance:.3f}" for idx,row in feature_imp.iterrows()])
    ax.text(0.5,0.5,imp_text, fontsize=10, ha='center', va='center')
    pdf.savefig()
    plt.close()

print(f"✅ Enhanced PDF 보고서 생성 완료: {pdf_path}")
