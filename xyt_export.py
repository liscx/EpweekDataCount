from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'Data')

# 先用硬编码的本地路径，不存在再下载
LOCAL_DRIVER = r"C:\Users\Epoint\.wdm\drivers\chromedriver\win64\148.0.7778.168\chromedriver-win32\chromedriver.exe"
if os.path.exists(LOCAL_DRIVER):
    driver_path = LOCAL_DRIVER
else:
    os.environ["WDM_SSL_VERIFY"] = "0"
    try:
        driver_path = ChromeDriverManager().install()
    except Exception:
        raise FileNotFoundError(f"本地chromedriver不存在且下载失败: {LOCAL_DRIVER}")


def wait_for_download(download_dir, existing_files, timeout=180):
    """等待新文件下载完成，忽略已有文件，返回新文件路径"""
    end_time = time.time() + timeout
    while time.time() < end_time:
        time.sleep(2)
        downloading = glob.glob(os.path.join(download_dir, '*.crdownload'))
        xlsx_files = glob.glob(os.path.join(download_dir, '*.xlsx'))
        new_files = [f for f in xlsx_files if f not in existing_files]
        if not downloading and new_files:
            return max(new_files, key=os.path.getmtime)
    raise TimeoutError("下载超时，未检测到新文件")


def main():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
    wait = WebDriverWait(driver, 180)

    # 每次页面加载后设置缩放为75%
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "document.addEventListener('DOMContentLoaded', () => { document.body.style.zoom = '75%'; });"
    })

    try:
        # 1. 登录
        print("[1/7] 打开登录页...")
        driver.get("https://xyt.etrading.cn/qytpframe/customframe4bid/login_TP")
        driver.execute_script("document.body.style.zoom = '75%'")

        username_input = wait.until(EC.presence_of_element_located((By.ID, "txtUserName")))
        password_input = driver.find_element(By.ID, "txtPwd")

        print("[2/7] 输入账号密码...")
        username_input.send_keys("ycyyzy")
        password_input.send_keys("Ep@2024Better")

        print("[3/7] 点击登录...")
        login_button = wait.until(EC.presence_of_element_located((By.ID, "btn_submit")))
        time.sleep(3)
        login_url = driver.current_url
        for _ in range(30):
            try:
                login_button.click()
                break
            except Exception:
                time.sleep(2)

        # 2. 导航：统计分析 -> 运营看板
        print("      等待页面跳转...")
        wait.until(EC.url_changes(login_url))

        print("[4/7] 导航: 统计分析 -> 运营看板...")
        stats_menu = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@class, 'menu-link')][.//span[text()='统计分析']]")))
        time.sleep(3)
        for _ in range(30):
            try:
                stats_menu.click()
                break
            except Exception:
                time.sleep(2)

        dashboard_menu = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@class, 'menu-link')][.//span[text()='运营看板']]")))
        time.sleep(3)
        for _ in range(30):
            try:
                dashboard_menu.click()
                break
            except Exception:
                time.sleep(2)

        # 3. 等待iframe出现
        print("[5/7] 等待看板加载...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, 'iframe[src*="shujutongjinew"]')))
        print("      已切入iframe")

        # 4. 记录已有文件，点击导出
        existing = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.xlsx')))
        print("[6/7] 点击导出订单...")
        export_button = wait.until(EC.visibility_of_element_located((By.ID, "btnExportOrder")))
        time.sleep(3)
        for _ in range(30):
            try:
                driver.execute_script("arguments[0].click();", export_button)
                print("      导出按钮已点击")
                break
            except Exception as e:
                print(f"      点击失败: {e}")
                time.sleep(2)

        # 5. 等待新文件下载完成
        print("[7/7] 等待文件下载...")
        downloaded = wait_for_download(DOWNLOAD_DIR, existing)

        # 重命名为 source_data.xlsx
        target = os.path.join(DOWNLOAD_DIR, 'source_data.xlsx')
        if os.path.exists(target):
            os.remove(target)
        os.rename(downloaded, target)
        print(f"文件已保存: {target}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
