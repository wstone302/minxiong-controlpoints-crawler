from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import urllib.request
import pandas as pd

# 圖片下載函式
def download_images(info, wait, driver, output_dir):
    point_id = info["控制點號"]
    point_folder = output_dir / point_id
    point_folder.mkdir(exist_ok=True)

    img_map = {
        "imgPointRecord": "點之記",
        "imgPointNearPhoto": "點位近照",
        "imgPointFarPhoto": "點位遠照",
        "imgPointEast": "樁標東向",
        "imgPointWest": "樁標西向",
        "imgPointSouth": "樁標南向",
        "imgPointNorth": "樁標北向"
    }

    for img_id, label in img_map.items():
        try:
            # 若無對應「觀看圖片」按鈕，跳過
            if len(driver.find_elements(By.ID, f"hl{img_id[3:]}")) == 0:
                continue
            view_link = wait.until(EC.element_to_be_clickable((By.ID, f"hl{img_id[3:]}")))
            driver.execute_script("arguments[0].click();", view_link)
            time.sleep(1)

            # 若無對應圖片標籤，跳過
            if len(driver.find_elements(By.ID, img_id)) == 0:
                continue
            img_tag = wait.until(EC.presence_of_element_located((By.ID, img_id)))
            img_url = img_tag.get_attribute("src")
            urllib.request.urlretrieve(img_url, point_folder / f"{label}.jpg")
            print(f"✅ {point_id} 已下載：{label}")
        except Exception as e:
            print(f"⚠️ {point_id} 無法下載 {label}：{e}")

# 主流程
output_dir = Path("C:/Users/user/Desktop/checkpoint")
output_dir.mkdir(parents=True, exist_ok=True)

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

driver.get("https://ctl.nlsc.gov.tw/cyhg/Portal/map.aspx")
time.sleep(2)

wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "easyui-west-iframe1")))

Select(driver.find_element(By.ID, "CountyList")).select_by_visible_text("嘉義縣")
time.sleep(1)
Select(driver.find_element(By.ID, "TownList")).select_by_visible_text("民雄鄉")
time.sleep(1)

driver.find_element(By.ID, "CoordSysTWD97_1").click()
driver.find_element(By.ID, "CoordSysTWD97[2020]_1").click()
driver.find_element(By.XPATH, "//label[contains(text(),'加密控制點')]/preceding-sibling::input").click()

driver.find_element(By.ID, "btn_search").click()
driver.switch_to.default_content()
wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ef_panel_iframe_F_1")))
time.sleep(1)

results = []
page = 1 

for _ in range(page - 1):
    try:
        next_btn = driver.find_element(By.XPATH, '//input[contains(@src, "arrow_gorightdown.png") and not(@disabled)]')
        driver.execute_script("arguments[0].click();", next_btn)
        time.sleep(2)
    except:
        print("⚠️ 無法提前跳轉到第 9 頁")
        break

while True:
    print(f"🔍 第 {page} 頁")
    rows = driver.find_elements(By.XPATH, '//table[contains(@class, "winTable")]/tbody/tr')[1:]

    for i in range(len(rows)):
        try:
            rows = driver.find_elements(By.XPATH, '//table[contains(@class, "winTable")]/tbody/tr')[1:]
            row = rows[i]
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 6 or "加密控制點" not in cols[4].text:
                continue

            info = {
                "行政區": cols[1].text.strip(),
                "控制點號": cols[2].text.strip(),
                "控制點名": cols[3].text.strip()
            }

            link = cols[5].find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].click();", link)
            time.sleep(2)

            plan_link = wait.until(EC.presence_of_element_located((By.XPATH, '//a[contains(@class,"BUT_listBOX")]')))
            driver.execute_script("arguments[0].click();", plan_link)
            time.sleep(2)

            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ef_panel_iframe_F_1")))
            iframe = wait.until(EC.presence_of_element_located((By.ID, "iframeResult")))
            driver.switch_to.frame(iframe)

            table = wait.until(EC.presence_of_element_located((By.XPATH, '//table[@class="dataListTable"]')))
            for row in table.find_elements(By.TAG_NAME, "tr"):
                try:
                    th = row.find_element(By.TAG_NAME, "th").text.strip()
                    td = row.find_element(By.TAG_NAME, "td").text.strip()
                    if th in ["測量坐標X", "測量坐標Y"]:
                        info[th] = td
                except:
                    continue

            download_images(info, wait, driver, output_dir)
            results.append(info)

            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ef_panel_iframe_F_1")))
            driver.execute_script("window.history.go(-2);")
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "winTable")))
            time.sleep(1)
        except Exception as e:
            print("⚠️ 錯誤：", e)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ef_panel_iframe_F_1")))
            continue

    try:
        next_btn = driver.find_element(By.XPATH, '//input[contains(@src, "arrow_gorightdown.png") and not(@disabled)]')
        driver.execute_script("arguments[0].click();", next_btn)
        page += 1
        time.sleep(2)
    except:
        print("✅ 沒有下一頁")
        break

df = pd.DataFrame(results)
excel_path = output_dir / "民雄鄉_加密控制點資訊_含圖片_第9頁起.xlsx"
df.to_excel(excel_path, index=False)
print(f"✅ 匯出完成，共 {len(df)} 筆")

driver.quit()
