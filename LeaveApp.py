#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請假小工具 - 自動填寫Google表單搶假系統
版本: 2.0 (加入顯示說明功能)
"""

import time
import ntplib
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import sys
import os
import subprocess

# ==================== 固定設定 ====================
# 表單網址
FORM_URL = "https://docs.google.com/forms/d/1hHHrf19cWw0Nn8C0RIgOwXwhcUQJSuNpqMoQCERuQVI/viewform"

# 固定填寫內容
FIXED_DATA = {
    "姓名": "孫煥然",
    "員工代號": "088624",
    "請假密碼": "dgKzhixM",
    "近假長假類型": "近假",
    "假別": "特休",
    "補充事項": "",
    "長假號碼牌": "",
}

# 台灣 NTP 伺服器
NTP_SERVER = "time.stdtime.gov.tw"

# ==================== 功能函數 ====================

def show_manual():
    """顯示使用手冊"""
    manual_path = os.path.join(os.path.dirname(__file__), "Manual.txt")

    if os.path.exists(manual_path):
        try:
            # 使用記事本開啟 Manual.txt
            subprocess.Popen(['notepad.exe', manual_path])
            print("[說明] 已開啟使用手冊")
        except Exception as e:
            print(f"[錯誤] 無法開啟使用手冊: {e}")
            print(f"手冊位置: {manual_path}")
    else:
        print(f"[錯誤] 找不到使用手冊檔案")
        print(f"預期位置: {manual_path}")

def print_banner():
    """顯示程式標題"""
    print("=" * 60)
    print("           線上請假自動填寫工具 v2.0")
    print("=" * 60)
    print()

def show_menu():
    """顯示主選單"""
    print("請選擇功能：")
    print("  [1] 開始執行請假程式")
    print("  [2] 查看使用說明")
    print("  [3] 退出程式")
    print()
    choice = input("請輸入選項 (1/2/3) > ").strip()
    return choice

def get_ntp_time():
    """從國家標準時間伺服器獲取精確時間"""
    try:
        client = ntplib.NTPClient()
        response = client.request(NTP_SERVER, version=3, timeout=5)
        ntp_time = datetime.fromtimestamp(response.tx_time)
        return ntp_time, True
    except Exception as e:
        print(f"[警告] 無法連接到 NTP 伺服器: {e}")
        print("[警告] 將使用系統時間，可能不夠精確")
        return datetime.now(), False

def parse_target_time(time_str):
    """解析目標時間字串
    格式: YYYY-MM-DD HH:MM:SS.sss
    例如: 2025-12-25 09:59:59.999
    """
    try:
        # 分離日期時間和毫秒
        if '.' in time_str:
            dt_part, ms_part = time_str.rsplit('.', 1)
            ms = int(ms_part)
        else:
            dt_part = time_str
            ms = 0

        # 解析日期時間
        target_dt = datetime.strptime(dt_part.strip(), "%Y-%m-%d %H:%M:%S")
        # 加上毫秒
        target_dt = target_dt.replace(microsecond=ms * 1000)

        return target_dt
    except ValueError as e:
        print(f"[錯誤] 時間格式錯誤: {e}")
        print("正確格式: YYYY-MM-DD HH:MM:SS.sss")
        print("例如: 2025-12-25 09:59:59.999")
        return None

def get_user_input():
    """取得使用者輸入"""
    print("[輸入] 請填寫以下資訊：")
    print("-" * 60)

    # 送出時間
    while True:
        print("\n請輸入送出時間 (格式: YYYY-MM-DD HH:MM:SS.sss)")
        print("例如: 2025-12-25 09:59:59.999")
        submit_time = input("送出時間 > ").strip()
        target_time = parse_target_time(submit_time)
        if target_time:
            break

    # 請假起始日期
    while True:
        print("\n請輸入請假起始日期 (格式: YYYY-MM-DD)")
        print("例如: 2025-12-25")
        start_date = input("起始日期 > ").strip()
        try:
            datetime.strptime(start_date, "%Y-%m-%d")
            break
        except ValueError:
            print("[錯誤] 日期格式錯誤，請重新輸入")

    # 請假結束日期
    while True:
        print("\n請輸入請假結束日期 (格式: YYYY-MM-DD)")
        print("例如: 2025-12-25")
        end_date = input("結束日期 > ").strip()
        try:
            datetime.strptime(end_date, "%Y-%m-%d")
            break
        except ValueError:
            print("[錯誤] 日期格式錯誤，請重新輸入")

    return target_time, start_date, end_date

def confirm_info(target_time, start_date, end_date):
    """確認資訊"""
    print("\n" + "=" * 60)
    print("             資訊確認")
    print("=" * 60)
    print(f"姓名：{FIXED_DATA['姓名']}")
    print(f"員工代號：{FIXED_DATA['員工代號']}")
    print(f"請假類型：{FIXED_DATA['近假長假類型']}")
    print(f"假別：{FIXED_DATA['假別']}")
    print(f"請假期間：{start_date} ~ {end_date}")
    print(f"送出時間：{target_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}")
    print("=" * 60)

    confirm = input("\n確認以上資訊正確？(Y/N) > ").strip().upper()
    return confirm == 'Y'

def setup_driver():
    """設定 Chrome WebDriver"""
    print("\n[初始化] 正在啟動瀏覽器...")
    options = Options()
    # options.add_argument('--headless')  # 無頭模式（背景執行）
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver

def fill_form(driver, start_date, end_date):
    """填寫表單（不送出）"""
    print("\n[處理中] 正在開啟表單...")
    driver.get(FORM_URL)

    wait = WebDriverWait(driver, 10)

    try:
        print("[處理中] 正在填寫表單...")

        # 等待表單載入
        time.sleep(2)

        # 1. 姓名 (第一個文字輸入框)
        name_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@type='text' and contains(@aria-label, '姓名')]")
        ))
        name_input.clear()
        name_input.send_keys(FIXED_DATA['姓名'])
        print(f"  ✓ 姓名: {FIXED_DATA['姓名']}")

        # 2. 員工代號
        emp_id_input = driver.find_element(By.XPATH, "//input[@type='text' and contains(@aria-label, '員工代號')]")
        emp_id_input.clear()
        emp_id_input.send_keys(FIXED_DATA['員工代號'])
        print(f"  ✓ 員工代號: {FIXED_DATA['員工代號']}")

        # 3. 近假/長假 (單選題 - 選擇"近假")
        radio_option = driver.find_element(By.XPATH, "//span[contains(text(), '近假')]/ancestor::div[@role='radio']")
        radio_option.click()
        print(f"  ✓ 請假類型: {FIXED_DATA['近假長假類型']}")

        # 4. 假別 (下拉選單 - 選擇"特休")
        dropdown = driver.find_element(By.XPATH, "//div[@role='listbox']")
        dropdown.click()
        time.sleep(0.5)
        special_leave = driver.find_element(By.XPATH, "//span[contains(text(), '特休')]/ancestor::div[@role='option']")
        special_leave.click()
        print(f"  ✓ 假別: {FIXED_DATA['假別']}")

        # 5. 請假起點日期
        start_date_input = driver.find_element(By.XPATH, "//input[@type='date' and contains(@aria-label, '起點')]")
        start_date_input.send_keys(start_date)
        print(f"  ✓ 起始日期: {start_date}")

        # 6. 請假終點日期
        end_date_input = driver.find_element(By.XPATH, "//input[@type='date' and contains(@aria-label, '終點')]")
        end_date_input.send_keys(end_date)
        print(f"  ✓ 結束日期: {end_date}")

        # 7. 確認勾選
        checkbox = driver.find_element(By.XPATH, "//span[contains(text(), '我確認了')]/ancestor::div[@role='checkbox']")
        if checkbox.get_attribute('aria-checked') != 'true':
            checkbox.click()
        print("  ✓ 已勾選確認")

        # 8. 請假密碼
        password_input = driver.find_element(By.XPATH, "//input[@type='password']")
        password_input.clear()
        password_input.send_keys(FIXED_DATA['請假密碼'])
        print("  ✓ 密碼: ********")

        print("\n[完成] 表單填寫完畢，等待送出時間...")
        return True

    except Exception as e:
        print(f"\n[錯誤] 填寫表單時發生錯誤: {e}")
        return False

def wait_and_submit(driver, target_time):
    """等待到指定時間並送出表單"""
    print("\n[同步中] 正在同步國家標準時間...")
    current_time, is_ntp = get_ntp_time()

    if is_ntp:
        print(f"[成功] 已同步國家標準時間: {current_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}")
    else:
        print(f"[警告] 使用系統時間: {current_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}")

    print(f"[目標] 送出時間: {target_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}")

    # 計算等待時間
    time_diff = (target_time - current_time).total_seconds()

    if time_diff < 0:
        print("\n[錯誤] 目標時間已過，無法執行")
        return False

    print(f"\n[等待中] 距離送出還有 {int(time_diff)} 秒...")

    # 如果等待時間超過 60 秒，先粗略等待
    if time_diff > 60:
        wait_time = time_diff - 60
        print(f"[等待中] 粗略等待 {int(wait_time)} 秒...")
        time.sleep(wait_time)
        time_diff = 60

    # 倒數 60 秒
    if time_diff > 10:
        wait_time = time_diff - 10
        for remaining in range(int(wait_time), 0, -1):
            print(f"\r[倒數中] {remaining} 秒...", end='', flush=True)
            time.sleep(1)
        print()

    # 最後 10 秒精確倒數
    print("\n[最後倒數]")
    while True:
        current_time, _ = get_ntp_time()
        remaining = (target_time - current_time).total_seconds()

        if remaining <= 0:
            break

        if remaining <= 10:
            print(f"\r  >>> {remaining:.3f} 秒 <<<", end='', flush=True)

        # 精確等待
        if remaining < 0.01:
            break
        elif remaining < 0.1:
            time.sleep(0.001)
        else:
            time.sleep(0.01)

    # 送出表單
    print("\n\n[送出!] 正在提交表單...")
    try:
        submit_button = driver.find_element(By.XPATH, "//span[contains(text(), '提交') or contains(text(), '送出')]/ancestor::div[@role='button']")
        submit_button.click()

        actual_time, _ = get_ntp_time()
        print(f"[成功] 表單已於 {actual_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} 送出！")

        # 等待確認頁面
        time.sleep(3)

        # 檢查是否成功
        if "已記錄您的回應" in driver.page_source or "您的回應已記錄" in driver.page_source:
            print("[成功] ✓ 表單提交成功！")
            return True
        else:
            print("[警告] 無法確認提交狀態，請手動檢查")
            return True

    except Exception as e:
        print(f"[錯誤] 送出表單時發生錯誤: {e}")
        return False

def run_leave_form():
    """執行請假表單填寫"""
    # 取得使用者輸入
    target_time, start_date, end_date = get_user_input()

    # 確認資訊
    if not confirm_info(target_time, start_date, end_date):
        print("\n[取消] 使用者取消操作")
        return

    driver = None
    try:
        # 設定瀏覽器
        driver = setup_driver()

        # 填寫表單
        if not fill_form(driver, start_date, end_date):
            print("\n[失敗] 表單填寫失敗")
            return

        # 等待並送出
        if wait_and_submit(driver, target_time):
            print("\n" + "=" * 60)
            print("             任務完成！")
            print("=" * 60)
        else:
            print("\n[失敗] 表單送出失敗")

        # 保持瀏覽器開啟 5 秒讓使用者查看結果
        print("\n瀏覽器將在 5 秒後關閉...")
        time.sleep(5)

    except KeyboardInterrupt:
        print("\n\n[中斷] 使用者中斷程式")
    except Exception as e:
        print(f"\n[錯誤] 程式執行錯誤: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if driver:
            driver.quit()
            print("[關閉] 瀏覽器已關閉")

# ==================== 主程式 ====================

def main():
    """主程式"""
    while True:
        print_banner()
        choice = show_menu()

        if choice == '1':
            # 執行請假程式
            print("\n" + "=" * 60)
            print("             開始執行請假程式")
            print("=" * 60 + "\n")
            run_leave_form()
            print("\n")
            input("按 Enter 鍵返回主選單...")
            print("\n" * 2)

        elif choice == '2':
            # 查看使用說明
            print("\n[說明] 正在開啟使用手冊...")
            show_manual()
            print("\n")
            input("按 Enter 鍵返回主選單...")
            print("\n" * 2)

        elif choice == '3':
            # 退出程式
            print("\n感謝使用，再見！")
            break

        else:
            print("\n[錯誤] 無效的選項，請重新選擇\n")
            time.sleep(1)
            print("\n" * 2)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程式已中斷")
    except Exception as e:
        print(f"\n程式發生錯誤: {e}")
        import traceback
        traceback.print_exc()
    finally:
        input("\n按 Enter 鍵退出...")
