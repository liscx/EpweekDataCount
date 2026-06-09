from playwright.sync_api import sync_playwright, expect
import os
import glob
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'Data')


def main():
    with sync_playwright() as p:
        # 启动浏览器，设置下载目录
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True
        )
        # 设置下载路径
        page = context.new_page()

        # 每次页面加载后设置缩放为75%
        page.add_init_script("""
            document.addEventListener('DOMContentLoaded', () => {
                document.body.style.zoom = '75%';
            });
        """)

        try:
            # 1. 登录
            print("[1/7] 打开登录页...")
            page.goto("https://xyt.etrading.cn/qytpframe/customframe4bid/login_TP")
            page.evaluate("document.body.style.zoom = '75%'")

            print("[2/7] 输入账号密码...")
            page.wait_for_selector("#txtUserName", timeout=180000)
            page.fill("#txtUserName", "ycyyzy")
            page.fill("#txtPwd", "Ep@2024Better")

            print("[3/7] 点击登录...")
            login_button = page.locator("#btn_submit")
            page.wait_for_timeout(3000)
            login_url = page.url

            for _ in range(30):
                try:
                    login_button.click()
                    break
                except Exception:
                    page.wait_for_timeout(2000)

            # 2. 导航：统计分析 -> 运营看板
            print("      等待页面跳转...")
            page.wait_for_url(lambda url: url != login_url, timeout=180000)

            print("[4/7] 导航: 统计分析 -> 运营看板...")
            stats_menu = page.locator("a.menu-link:has(span:text('统计分析'))")
            page.wait_for_timeout(3000)
            for _ in range(30):
                try:
                    stats_menu.click()
                    break
                except Exception:
                    page.wait_for_timeout(2000)

            dashboard_menu = page.locator("a.menu-link:has(span:text('运营看板'))")
            page.wait_for_timeout(3000)
            for _ in range(30):
                try:
                    dashboard_menu.click()
                    break
                except Exception:
                    page.wait_for_timeout(2000)

            # 3. 等待iframe出现
            print("[5/7] 等待看板加载...")
            iframe = page.frame_locator("iframe[src*='shujutongjinew']")
            print("      已定位到iframe")

            # 4. 记录已有文件，点击导出
            existing = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.xlsx')))
            print("[6/7] 点击导出订单...")
            export_button = iframe.locator("#btnExportOrder")
            page.wait_for_timeout(3000)

            # 使用 Playwright 的下载事件
            with page.expect_download(timeout=180000) as download_info:
                for _ in range(30):
                    try:
                        # 先滚动到元素位置，再用 JS 点击
                        export_button.scroll_into_view_if_needed()
                        page.evaluate("document.querySelector('iframe[src*=\"shujutongjinew\"]').contentDocument.querySelector('#btnExportOrder').click()")
                        print("      导出按钮已点击")
                        break
                    except Exception as e:
                        print(f"      点击失败: {e}")
                        page.wait_for_timeout(2000)

            # 5. 等待新文件下载完成
            print("[7/7] 等待文件下载...")
            download = download_info.value
            # 保存到目标路径
            target = os.path.join(DOWNLOAD_DIR, 'source_data.xlsx')
            download.save_as(target)
            print(f"文件已保存: {target}")

        finally:
            browser.close()


if __name__ == "__main__":
    main()
