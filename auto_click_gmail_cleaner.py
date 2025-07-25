"""
⚠️  初次使用要下：
    "C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="D:\ChromeProfile"
    到該視窗登入Google帳號
    不然程式會被登入環節卡住
"""
"""
pip install selenium webdriver-manager
python auto_click_gmail_cleaner.py
"""
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- 設定區 ---
CHROME_PROFILE_PATH = r"D:\ChromeProfile"
APPS_SCRIPT_URL = "https://script.google.com/home/projects/10qWTgnpfnPJB46itFCkBQtERibZvmNiTzh95VG6im71PLz3pgPbE-nGg/edit"
SUCCESS_MESSAGE = "郵件清理任務執行完畢"
TIMEOUT_MESSAGE = "Exceeded maximum execution time"

# --- 腳本開始 ---

options = webdriver.ChromeOptions()
options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}")
options.add_argument("--start-maximized")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

service = ChromeService(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

run_count = 1
is_all_done = False

try:
    print(f"正在前往您的 Apps Script 專案: {APPS_SCRIPT_URL}")
    driver.get(APPS_SCRIPT_URL)
    
    wait = WebDriverWait(driver, 60)
    short_wait = WebDriverWait(driver, 5)

    while not is_all_done:
        print(f"\n--- 第 {run_count} 次嘗試執行 ---")
        
        # 1. [關鍵修正] 使用更精準的 XPath 等待並點擊「執行」按鈕
        try:
            # 找到一個 button，它的後代元素中有一個 span，其文字是'執行'
            run_button_xpath = "//button[.//span[text()='執行']]"
            run_button = wait.until(EC.element_to_be_clickable((By.XPATH, run_button_xpath)))
            
            print("偵測到『執行』按鈕，點擊執行...")
            run_button.click()
            run_count += 1
            
        except TimeoutException:
            print("錯誤：在 60 秒內找不到可點擊的『執行』按鈕。請確認頁面已載入或按鈕文字是否正確。")
            break

        # 2. 監控執行記錄
        print("腳本執行中，開始監控執行記錄...(最長等待 7 分鐘)")
        
        monitoring_start_time = time.time()
        log_pane_found = False
        
        while time.time() - monitoring_start_time < 420:
            try:
                # 尋找日誌面板容器
                log_pane = driver.find_element(By.ID, "log-container")
                if not log_pane_found:
                    print("已偵測到日誌面板。")
                    log_pane_found = True

                log_text = log_pane.text

                if SUCCESS_MESSAGE in log_text:
                    print(f"\n🎉 偵測到成功訊息: '{SUCCESS_MESSAGE}'")
                    print("所有郵件已清理完畢，任務終止！")
                    is_all_done = True
                    break

                if TIMEOUT_MESSAGE in log_text:
                    print(f"\n🕒 偵測到超時訊息: '{TIMEOUT_MESSAGE}'")
                    print("準備進行下一次執行...")
                    try:
                        # 點擊「清除所有日誌」按鈕，比執行 JS 更穩定
                        clear_log_button_xpath = "//button[@aria-label='清除所有日誌']"
                        clear_button = short_wait.until(EC.element_to_be_clickable((By.XPATH, clear_log_button_xpath)))
                        clear_button.click()
                        print("已點擊清除日誌按鈕。")
                    except:
                        print("警告：找不到清除日誌按鈕，將嘗試用 JS 清除。")
                        driver.execute_script("document.getElementById('log-container').innerHTML = '';")
                    
                    time.sleep(3)
                    break
                
                time.sleep(5)

            except NoSuchElementException:
                # 如果日誌面板還沒出現，是很正常的，繼續等待
                if not log_pane_found:
                    print("等待日誌面板出現...")
                time.sleep(5)
            except Exception as e:
                print(f"監控日誌時發生輕微錯誤: {e}")
                time.sleep(5)

        # 檢查是否是因為監控超時而跳出迴圈
        if not is_all_done and (time.time() - monitoring_start_time >= 420):
            print("\n錯誤：監控超過 7 分鐘，但未偵測到成功或超時訊息。腳本可能已卡住，終止執行。")
            break

except Exception as e:
    print(f"\n腳本執行過程中發生嚴重錯誤: {e}")
    driver.save_screenshot("supervisor_error.png")
    print("已儲存錯誤截圖: supervisor_error.png")

finally:
    print("\n自動監工腳本執行結束。")
    time.sleep(10)
    driver.quit()