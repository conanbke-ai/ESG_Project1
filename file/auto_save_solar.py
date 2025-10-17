import os
import time
import glob
import shutil
import calendar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def move_and_encode_csv(base_path: str, year_dir: str, file_name: str):
    """다운로드된 최신 CSV를 찾아 인코딩 변환 후 이동"""
    list_files = glob.glob(os.path.join(base_path, "*.csv"))
    if not list_files:
        return None

    latest_file = max(list_files, key=os.path.getmtime)
    dest_path = os.path.join(year_dir, file_name)

    # 한글 깨짐 방지: UTF-8-sig로 재저장
    try:
        with open(latest_file, 'r', encoding='cp949', errors='ignore') as f:
            content = f.read()
        with open(dest_path, 'w', encoding='utf-8-sig') as f:
            f.write(content)
    except Exception as e:
        print(f"⚠ 인코딩 변환 실패: {e}")
        shutil.copy(latest_file, dest_path)

    os.remove(latest_file)
    return dest_path


def download_solar_data(base_path: str, year: int, max_retry: int = 3):
    """지정 연도의 태양광 발전량 데이터를 월별로 다운로드"""
    os.makedirs(base_path, exist_ok=True)

    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": base_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)
    driver.get("https://www.koenergy.kr/kosep/gv/nf/dt/nfdt21/main.do")
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.ID, "strDateS")))

    year_dir = os.path.join(base_path, str(year))
    os.makedirs(year_dir, exist_ok=True)

    for month in range(1, 13):
        start_date = f"{year}{month:02d}01"
        last_day = calendar.monthrange(year, month)[1]
        end_date = f"{year}{month:02d}{last_day:02d}"
        file_name = f"남동발전량_{year}_{month:02d}.csv"

        print(f"▶ {year}-{month:02d} ({start_date} ~ {end_date}) 조회 중...")

        success = False
        for attempt in range(1, max_retry + 1):
            try:
                # 날짜 JS 직접 설정
                driver.execute_script(f"document.getElementById('strDateS').value='{start_date}'")
                driver.execute_script(f"document.getElementById('strDateE').value='{end_date}'")

                # 조회 버튼 클릭
                search_btn = driver.find_element(By.XPATH, "//a[contains(@href, 'goSubmit') or contains(text(),'조회')]")
                search_btn.click()
                time.sleep(3)

                # CSV 다운로드 실행
                driver.execute_script("goCsvDown();")
                time.sleep(5)

                # 파일 이동 및 인코딩 처리
                dest_path = move_and_encode_csv(base_path, year_dir, file_name)
                if dest_path:
                    print(f"✅ 저장 완료 → {dest_path}")
                    success = True
                    break
                else:
                    print(f"⚠ 다운로드 파일을 찾지 못했습니다. 재시도 {attempt}/{max_retry}")
                    time.sleep(2)

            except Exception as e:
                print(f"❌ {year}-{month:02d} 오류 발생: {e}. 재시도 {attempt}/{max_retry}")
                time.sleep(2)

        if not success:
            print(f"❌ {year}-{month:02d} 다운로드 실패, 다음 달로 진행합니다.")

    driver.quit()
    print(f"\n🎉 {year}년 CSV 파일이 '{year_dir}'에 저장 완료되었습니다.")