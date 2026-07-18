#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modrinth 批量注册 + 收藏夹管理工具 - GUI版

修复：Windows无GPU环境下多开浏览器卡死问题
策略：串行启动浏览器 + 最小化 + 并发执行后续流程

修复内容：
1. 双重driver.quit()保护
2. 任务完成后从browsers字典移除，防止内存泄漏
3. 启动失败时进度正确显示
4. API阶段响应暂停事件
"""

import os
import sys
import time
import random
import string
import requests
import threading
import queue
import warnings
import unicodedata
import shutil
import tempfile
import subprocess
from threading import Thread, Event
from datetime import datetime
from pathlib import Path
from copy import copy
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib3.exceptions import InsecureRequestWarning
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import tkinter as tk
    from tkinter import ttk, scrolledtext, messagebox, filedialog
except ImportError:
    print("GUI模式需要tkinter，当前环境不支持")
    sys.exit(1)

import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from urllib.parse import quote
from selenium.common.exceptions import TimeoutException, NoSuchElementException


# 屏蔽证书警告
warnings.filterwarnings("ignore", category=InsecureRequestWarning)


# ====================== 全局配置 ======================
DEFAULT_THREAD_COUNT = 3  # 建议2，Windows无GPU环境稳定
MAX_COLLECTIONS_PER_USER = 32

# 全局信号量：控制同时"活跃窗口"数量（人工操作hCaptcha时需要）
ACTIVE_WINDOW = threading.Semaphore(1)


def init_browser(task_id: int, attempt: int = 1):
    """
    跨平台浏览器初始化 - 修复Windows无GPU多开卡死
    策略：小窗口 + 最小化启动 + 卡死检测
    """
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    local_driver = os.path.join(base_dir, "chromedriver.exe")
    portable_chrome = os.path.join(base_dir, "chrome", "chrome.exe")

    print(f"📁 程序目录：{base_dir}")
    print(f"🔍 查找 ChromeDriver：{local_driver}")
    print(f"🔍 查找 Chrome：{portable_chrome}")

    if not os.path.exists(local_driver):
        raise FileNotFoundError(f"找不到 chromedriver.exe，请确保与 exe 放在同一目录\n查找路径：{local_driver}")

    if not os.path.exists(portable_chrome):
        raise FileNotFoundError(
            f"找不到 chrome.exe，请确保 chrome 文件夹与 exe 放在同一目录\n查找路径：{portable_chrome}")

    print(f"✅ 使用本地 ChromeDriver：{local_driver}")
    print(f"✅ 使用本地 Chrome：{portable_chrome}")

    # ========== 关键修复：使用临时目录 + 强制清理 ==========
    user_data_dir = os.path.join(
        tempfile.gettempdir(),
        f"modrinth_chrome_{task_id}_{threading.current_thread().ident}_{int(time.time()*1000)}"
    )

    # 强制清理残留
    if os.path.exists(user_data_dir):
        try:
            shutil.rmtree(user_data_dir, ignore_errors=True)
            time.sleep(0.3)
        except Exception as e:
            print(f"⚠️ 清理旧数据目录失败: {e}")

    # 杀死残留进程
    try:
        subprocess.run(
            f'taskkill /F /IM chrome.exe /FI "COMMANDLINE LIKE %{user_data_dir}%"',
            shell=True, capture_output=True
        )
        time.sleep(0.3)
    except:
        pass

    os.makedirs(user_data_dir, exist_ok=True)

    # 欺骗SingletonLock
    try:
        with open(os.path.join(user_data_dir, "SingletonLock"), "w") as f:
            f.write("fake_lock")
    except:
        pass

    options = webdriver.ChromeOptions()
    options.binary_location = portable_chrome

    # ========== 关键：窗口控制参数（防卡死） ==========
    # 小窗口，分散位置
    x_pos = (task_id % 3) * 400
    y_pos = (task_id // 3) * 50
    options.add_argument(f"--window-size=900,600")
    options.add_argument(f"--window-position={x_pos},{y_pos}")

    # 禁用GPU和特效（无GPU环境）
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-gpu-sandbox")
    options.add_argument("--disable-features=SmoothScrolling,OverlayScrollbar,CanvasOopRasterization")
    options.add_argument("--disable-animations")
    options.add_argument("--disable-lcd-text")
    options.add_argument("--disable-font-subpixel-positioning")

    # 基础稳定参数
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-renderer-backgrounding")

    # 禁用预创建和通知
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")

    # 移除有问题的参数
    # ❌ 无 --start-maximized
    # ❌ 无 --disable-background-networking
    # ❌ 无 --remote-debugging-port

    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    options.add_argument(f"--user-data-dir={user_data_dir}")

    service = Service(local_driver)

    # 关键：启动间隔，错开桌面堆分配
    wait_time = 6 if attempt == 1 else 3
    time.sleep(wait_time)

    driver = webdriver.Chrome(service=service, options=options)

    # 关键：启动后立即最小化，减少GDI占用
    driver.minimize_window()

    # 隐藏webdriver标志
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined})
        """
    })

    wait = WebDriverWait(driver, 15)
    short_wait = WebDriverWait(driver, 3)
    return driver, wait, short_wait, user_data_dir


def safe_init_browser(task_id: int, log_callback=None):
    """
    带卡死检测和自动重试的浏览器启动
    """
    for attempt in range(1, 3):
        driver = None
        user_data_dir = None
        try:
            if log_callback:
                log_callback(f"[用户{task_id}] 启动浏览器（尝试 {attempt}/2）...")

            driver, wait, short_wait, user_data_dir = init_browser(task_id, attempt)

            # 关键检测：访问about:blank测试渲染是否正常
            driver.set_page_load_timeout(10)
            driver.get("about:blank")

            # 执行简单JS确认渲染进程活着
            result = driver.execute_script("return document.readyState")
            if result != "complete":
                raise Exception("渲染进程无响应")

            if log_callback:
                log_callback(f"[用户{task_id}] ✅ 浏览器启动成功，已最小化")

            return driver, wait, short_wait, user_data_dir

        except Exception as e:
            if log_callback:
                log_callback(f"[用户{task_id}] ⚠️ 启动失败: {str(e)[:100]}")
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            # 清理user_data_dir
            if user_data_dir and os.path.exists(user_data_dir):
                try:
                    shutil.rmtree(user_data_dir, ignore_errors=True)
                except:
                    pass
            if attempt == 2:
                raise Exception(f"浏览器启动失败，已重试: {e}")
            time.sleep(5)

    raise Exception("浏览器启动失败，已重试")


# ===================== 辅助函数 =====================
def retry_click(driver, element, max_retries=3, delay=0.5):
    """带重试的点击"""
    from selenium.webdriver.common.action_chains import ActionChains
    last_error = None
    for attempt in range(max_retries):
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});", element
            )
            time.sleep(0.2)
            element.click()
            return True
        except Exception as e1:
            last_error = e1
            try:
                ActionChains(driver).move_to_element(element).click().perform()
                return True
            except Exception as e2:
                last_error = e2
                try:
                    driver.execute_script("arguments[0].click();", element)
                    return True
                except Exception as e3:
                    last_error = e3
                    time.sleep(delay)
    return False


def random_qq_email():
    """生成随机QQ邮箱"""
    chars = string.ascii_letters + string.digits
    prefix = ''.join(random.choice(chars) for _ in range(8))
    return f"{prefix}@qq.com"


def display_width(text):
    """中文/全角字符按2宽度计算"""
    return sum(2 if unicodedata.east_asian_width(c) in ('F','W') else 1 for c in str(text or ''))


def auto_fit_columns(ws, min_w=8, max_w=50, padding=3):
    for col_cells in ws.columns:
        letter = col_cells[0].column_letter
        w = max((display_width(c.value) for c in col_cells
                 if not isinstance(c, openpyxl.cell.cell.MergedCell) and c.value is not None), default=0)
        ws.column_dimensions[letter].width = max(min_w, min(w * 1.1 + padding, max_w))


def append_link_to_txt(link: str, file_path: str = "links.txt"):
    """将单个链接追加写入txt文件"""
    with open(file_path, "a", encoding="utf-8") as f:
        f.write(link + "\n")


# ===================== 单用户注册 + 创建收藏夹任务 =====================
def single_user_task(task_id: int, user_titles: list, user_intros: list,
                     output_dir: str, log_callback=None, browser_tuple=None,
                     stop_event=None, pause_event=None):
    """
    单个用户完整流程：使用已启动的浏览器 -> 执行注册流程
    browser_tuple: (driver, wait, short_wait, user_data_dir) 如果为None则自行启动
    """
    driver = None
    session = None
    token = None
    failed_titles = []
    failed_intros = []
    success_count = 0
    user_data_dir = None

    try:
        if log_callback:
            log_callback(f"[用户{task_id}] ===== 任务开始 =====")
            log_callback(f"[用户{task_id}] 需要创建收藏夹: {len(user_titles)} 个")

        # ========== 使用预启动的浏览器或自行启动 ==========
        if browser_tuple:
            driver, wait, short_wait, user_data_dir = browser_tuple
            if log_callback:
                log_callback(f"[用户{task_id}] 使用预启动的浏览器")
        else:
            if log_callback:
                log_callback(f"[用户{task_id}] 启动浏览器...")
            driver, wait, short_wait, user_data_dir = safe_init_browser(task_id, log_callback)

        long_wait = WebDriverWait(driver, 6000)

        # ========== 访问目标网址 ==========
        if browser_tuple is None:
            if log_callback:
                log_callback(f"[用户{task_id}] 访问 modrinth.com...")
            driver.set_page_load_timeout(30)
            driver.get("https://modrinth.com")
        else:
            if log_callback:
                log_callback(f"[用户{task_id}] 访问 modrinth.com...")
            driver.set_page_load_timeout(30)
            driver.get("https://modrinth.com")

        # 恢复窗口（需要人工操作hCaptcha时）
        with ACTIVE_WINDOW:
            driver.maximize_window()
            if log_callback:
                log_callback(f"[用户{task_id}] 窗口已恢复，开始注册流程...")

            # 1. 点击注册
            signup_btn = long_wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@href="/auth/sign-up"]')))
            if not retry_click(driver, signup_btn):
                raise Exception(f"点击注册按钮失败")
            if log_callback:
                log_callback(f"[用户{task_id}] 点击注册按钮")

            # 2. 输入随机邮箱
            email_input = long_wait.until(EC.visibility_of_element_located((By.ID, "email")))
            random_email = random_qq_email()
            email_input.clear()
            email_input.send_keys(random_email)
            if log_callback:
                log_callback(f"[用户{task_id}] 输入随机邮箱: {random_email}")

            # 3. 输入密码
            pwd_input = long_wait.until(EC.visibility_of_element_located((By.ID, "password")))
            pwd_input.clear()
            pwd_input.send_keys("Admin@coc1")
            if log_callback:
                log_callback(f"[用户{task_id}] 输入密码")

            # 4. 点击 Continue with Email
            continue_btn = long_wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(text(), 'Continue with Email')]")
                )
            )
            if not retry_click(driver, continue_btn):
                raise Exception(f"点击 Continue with Email 失败")
            if log_callback:
                log_callback(f"[用户{task_id}] 点击 Continue with Email")

            # 5. 选择生日
            picker_wrap = long_wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "modrinth-date-picker")))
            if not retry_click(driver, picker_wrap):
                raise Exception(f"点击日期选择器失败")
            time.sleep(0.6)

            month_select = long_wait.until(EC.element_to_be_clickable((
                By.CSS_SELECTOR, "select.modrinth-monthDropdown-months"
            )))
            if not retry_click(driver, month_select):
                raise Exception(f"点击月份选择失败")
            time.sleep(0.2)
            august_option = long_wait.until(EC.element_to_be_clickable((
                By.CSS_SELECTOR, "select.modrinth-monthDropdown-months option[value='7']"
            )))
            if not retry_click(driver, august_option):
                raise Exception(f"选择八月失败")
            time.sleep(0.3)

            year_input = long_wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR, "input.numInput.cur-year"
            )))
            year_input.clear()
            year_input.send_keys("1998")
            time.sleep(0.3)

            day23 = long_wait.until(EC.element_to_be_clickable((
                By.XPATH, '//span[@aria-label="August 23, 1998"]'
            )))
            if not retry_click(driver, day23):
                raise Exception(f"选择日期失败")
            time.sleep(0.4)

            blank_target = long_wait.until(EC.element_to_be_clickable((
                By.XPATH, "//*[contains(text(), 'Date of birth')]"
            )))
            if not retry_click(driver, blank_target):
                raise Exception(f"点击空白处关闭日期选择器失败")
            if log_callback:
                log_callback(f"[用户{task_id}] 生日选择完成")

            # 6. hCaptcha 验证（原有代码，不动）
            hcaptcha_iframe = long_wait.until(
                EC.presence_of_element_located((
                    By.CSS_SELECTOR,
                    "iframe[src*='newassets.hcaptcha.com'][src*='frame=checkbox']"
                ))
            )

            print(f"[hCaptcha] 发现 checkbox iframe")

            driver.switch_to.frame(hcaptcha_iframe)
            print("[hCaptcha] 已切换到 iframe 内部")

            checkbox = long_wait.until(
                EC.presence_of_element_located((By.ID, "checkbox"))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                checkbox
            )
            time.sleep(0.3)

            print(f"[hCaptcha] 点击前 aria-checked: {checkbox.get_attribute('aria-checked')}")

            actions = ActionChains(driver)
            actions.move_to_element(checkbox)
            actions.click()
            actions.perform()

            print("✅ [hCaptcha] ActionChains 点击完成")

            driver.switch_to.default_content()
            print("[hCaptcha] 已切回主文档")

            print("\n⏳ 等待手动完成 hCaptcha 验证...")

            max_wait_time = 600
            poll_interval = 2
            elapsed = 0
            verified = False

            while elapsed < max_wait_time:
                time.sleep(poll_interval)
                elapsed += poll_interval

                try:
                    checkbox_iframes = driver.find_elements(
                        By.CSS_SELECTOR,
                        "iframe[src*='newassets.hcaptcha.com'][src*='frame=checkbox']"
                    )

                    if checkbox_iframes:
                        driver.switch_to.frame(checkbox_iframes[0])
                        try:
                            cb = driver.find_element(By.ID, "checkbox")
                            if cb.get_attribute("aria-checked") == "true":
                                verified = True
                        except:
                            pass
                        driver.switch_to.default_content()
                    else:
                        challenge_iframes = driver.find_elements(
                            By.CSS_SELECTOR,
                            "iframe[src*='newassets.hcaptcha.com'][src*='frame=challenge']"
                        )
                        if not challenge_iframes:
                            verified = True
                        else:
                            if int(elapsed) % 10 == 0 and log_callback:
                                log_callback(f"[用户{task_id}]    ...等待完成图片挑战...")
                            driver.switch_to.default_content()
                            continue

                    if verified:
                        if log_callback:
                            log_callback(f"[用户{task_id}] ✅ hCaptcha 验证通过!")
                        break

                except Exception as e:
                    driver.switch_to.default_content()
                    pass

                if int(elapsed) % 10 == 0 and log_callback:
                    log_callback(f"[用户{task_id}]    ...已等待 {int(elapsed)} 秒，请手动点击验证框...")

            else:
                raise TimeoutError("hCaptcha 验证等待超时（5分钟未检测到通过）")

            # 7. 勾选邮件订阅 + 完成注册
            keep_check = long_wait.until(
                EC.element_to_be_clickable((By.XPATH, '//span[contains(@class, "checkbox-shadow")]'))
            )
            if not retry_click(driver, keep_check):
                raise Exception(f"勾选邮件订阅失败")

            finish_register_btn = long_wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Complete sign up')]"))
            )
            if not retry_click(driver, finish_register_btn):
                raise Exception(f"点击完成注册按钮失败")
            if log_callback:
                log_callback(f"[用户{task_id}] 注册完成!")
            time.sleep(5)

            # hCaptcha完成后，最小化窗口释放GDI
            driver.minimize_window()
            if log_callback:
                log_callback(f"[用户{task_id}] 窗口已最小化，继续API操作...")

        # 8. 提取 Token
        cookies = driver.get_cookies()
        for ck in cookies:
            if ck["name"] == "auth-token":
                token = ck["value"]
                break
        if not token:
            raise Exception("无法获取 auth-token")
        if log_callback:
            log_callback(f"[用户{task_id}] 获取 Token 成功")

        # 9. 创建 requests session
        retry_strategy = Retry(
            total=5, backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET", "POST", "PATCH"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session = requests.Session()
        session.verify = False
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        session.headers.update({
            "Authorization": f"Bearer {token}",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })

        # 10. 创建收藏夹（支持暂停）
        collection_ids = []
        for i, (title, intro) in enumerate(zip(user_titles, user_intros)):
            # 检查暂停
            if pause_event and pause_event.is_set():
                if log_callback:
                    log_callback(f"[用户{task_id}] ⏸ 暂停中...")
                while pause_event.is_set() and (not stop_event or not stop_event.is_set()):
                    time.sleep(0.5)
                if log_callback:
                    log_callback(f"[用户{task_id}] ▶ 继续执行...")

            # 检查停止
            if stop_event and stop_event.is_set():
                if log_callback:
                    log_callback(f"[用户{task_id}] ⏹ 收到停止信号，中断任务")
                break

            if log_callback:
                log_callback(f"[用户{task_id}] 创建收藏夹 {i+1}/{len(user_titles)}: {title[:30]}...")

            create_payload = {
                "name": title,
                "description": intro,
                "projects": []
            }
            resp = session.post("https://api.modrinth.com/v3/collection", json=create_payload)
            time.sleep(random.uniform(0.1, 0.5))

            if resp.status_code == 200:
                collection_id = resp.json()["id"]
                collection_ids.append(collection_id)
                success_count += 1
                if log_callback:
                    log_callback(f"[用户{task_id}] 收藏夹创建成功! ID: {collection_id}")
            else:
                failed_titles.append(title)
                failed_intros.append(intro)
                if log_callback:
                    log_callback(f"[用户{task_id}] 创建收藏夹失败: {resp.status_code} - {resp.text[:100]}")

        # 11. 搜索并关注项目（支持暂停）
        if log_callback:
            log_callback(f"[用户{task_id}] 搜索热门模组...")

        # 检查暂停
        if pause_event and pause_event.is_set():
            while pause_event.is_set() and (not stop_event or not stop_event.is_set()):
                time.sleep(0.5)

        if stop_event and stop_event.is_set():
            if log_callback:
                log_callback(f"[用户{task_id}] ⏹ 收到停止信号，跳过后续操作")
        else:
            search_resp = session.get(
                "https://api.modrinth.com/v2/search",
                params={"limit": 20, "index": "relevance", "new_filters": "project_types = `mod`"}
            )
            time.sleep(random.uniform(0.1, 0.5))

            if search_resp.status_code == 200:
                hits = search_resp.json().get("hits", [])
                if hits:
                    target_id = hits[0]['project_id']
                    session.post(f"https://api.modrinth.com/v2/project/{target_id}/follow")
                    if log_callback:
                        log_callback(f"[用户{task_id}] 已关注项目: {target_id}")

                    # 批量加入收藏
                    for cid in collection_ids:
                        # 检查暂停
                        if pause_event and pause_event.is_set():
                            while pause_event.is_set() and (not stop_event or not stop_event.is_set()):
                                time.sleep(0.5)
                        if stop_event and stop_event.is_set():
                            break

                        update_resp = session.patch(
                            f"https://api.modrinth.com/v3/collection/{cid}",
                            json={"new_projects": [target_id]}
                        )
                        time.sleep(random.uniform(0.1, 0.5))
                        if update_resp.status_code in [200, 204]:
                            link = f"https://modrinth.com/collection/{cid}"
                            append_link_to_txt(link, os.path.join(output_dir, "collection_links.txt"))
                            if log_callback:
                                log_callback(f"[用户{task_id}] 项目已加入收藏夹: {cid}")

        if log_callback:
            if failed_titles:
                log_callback(f"[用户{task_id}] 部分完成! 成功 {success_count} 个, 失败 {len(failed_titles)} 个收藏夹")
            else:
                log_callback(f"[用户{task_id}] 全部完成! 创建了 {len(collection_ids)} 个收藏夹")
        return f"用户{task_id} 成功 {success_count}/{len(user_titles)}"

    except Exception as e:
        import traceback
        error_msg = f"[用户{task_id}] 错误: {str(e)}\n{traceback.format_exc()}"
        if log_callback:
            log_callback(error_msg)
        else:
            print(error_msg)

        if not failed_titles and user_titles:
            failed_titles = list(user_titles)
            failed_intros = list(user_intros)

        if failed_titles:
            try:
                failed_title_path = os.path.join(output_dir, "失败标题.txt")
                failed_intro_path = os.path.join(output_dir, "失败简介.txt")
                with open(failed_title_path, "a", encoding="utf-8") as ft:
                    for t in failed_titles:
                        ft.write(t + "\n")
                with open(failed_intro_path, "a", encoding="utf-8") as fi:
                    for intro in failed_intros:
                        fi.write(intro + "\n")
                if log_callback:
                    log_callback(f"[用户{task_id}] 已写入 {len(failed_titles)} 条失败记录到 {output_dir}")
            except Exception as write_err:
                if log_callback:
                    log_callback(f"[用户{task_id}] 写入失败文件出错: {write_err}")

        return f"用户{task_id} 失败: {str(e)}"
    finally:
        # 清理资源
        if session and token:
            try:
                session.delete(f"https://api.modrinth.com/v2/session/{token}")
            except:
                pass
        if session:
            try:
                session.close()
            except:
                pass
        if driver:
            try:
                driver.quit()
            except:
                pass
        # 清理user_data_dir
        if user_data_dir and os.path.exists(user_data_dir):
            try:
                shutil.rmtree(user_data_dir, ignore_errors=True)
            except:
                pass


# ====================================================================
# ===================== GUI 主程序 =====================


class ModrinthCollector:
    """Modrinth 收藏夹分配与注册引擎"""
    MAX_PER_USER = 32

    def __init__(self, title_files, intro_files, output_dir, thread_count,
                 log_callback=None, progress_callback=None):
        self.title_files = title_files
        self.intro_files = intro_files
        self.output_dir = output_dir
        self.thread_count = thread_count
        self.log_callback = log_callback
        self.progress_callback = progress_callback
        self.stop_event = Event()
        self.pause_event = Event()
        self.lock = threading.Lock()
        self.total_processed = 0
        self.browsers = {}  # 预启动的浏览器 {task_id: (driver, wait, short_wait, user_data_dir)}

    def _log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_msg = f"[{timestamp}] {message}"
        print(log_msg)
        if self.log_callback:
            self.log_callback(log_msg)

    def _read_lines_from_files(self, file_paths):
        """从多个txt文件中按行读取内容"""
        lines = []
        for fp in file_paths:
            if not os.path.exists(fp):
                self._log(f"⚠️ 文件不存在，跳过: {fp}")
                continue
            try:
                with open(fp, "r", encoding="utf-8") as f:
                    file_lines = []
                    for line in f:
                        line = line.strip()
                        if line:
                            file_lines.append(line)
                    lines.extend(file_lines)
                self._log(f"   从 [{os.path.basename(fp)}] 读取 {len(file_lines)} 行")
            except Exception as e:
                self._log(f"⚠️ 读取文件失败 [{fp}]: {e}")
        return lines

    def _distribute_to_users(self, titles):
        """
        将标题按顺序分发给用户，每个用户最多32个收藏夹，
        且同一个用户的32个标题不能重复。
        """
        users = []
        current_user_titles = []
        current_user_seen = set()
        user_idx = 0

        for title in titles:
            if self.stop_event.is_set():
                break
            while self.pause_event.is_set() and not self.stop_event.is_set():
                time.sleep(0.1)

            if title in current_user_seen:
                if current_user_titles:
                    users.append({
                        "user_index": user_idx,
                        "titles": current_user_titles.copy()
                    })
                    self._log(f"   用户 #{user_idx + 1} 分配完成，共 {len(current_user_titles)} 个标题")
                user_idx += 1
                current_user_titles = [title]
                current_user_seen = {title}
                self._log(f"   标题重复 [{title[:30]}...] -> 创建新用户 #{user_idx + 1}")
            else:
                if len(current_user_titles) >= self.MAX_PER_USER:
                    users.append({
                        "user_index": user_idx,
                        "titles": current_user_titles.copy()
                    })
                    self._log(f"   用户 #{user_idx + 1} 已满32个，创建新用户 #{user_idx + 2}")
                    user_idx += 1
                    current_user_titles = [title]
                    current_user_seen = {title}
                else:
                    current_user_titles.append(title)
                    current_user_seen.add(title)

        if current_user_titles and not self.stop_event.is_set():
            users.append({
                "user_index": user_idx,
                "titles": current_user_titles.copy()
            })
            self._log(f"   用户 #{user_idx + 1} 分配完成，共 {len(current_user_titles)} 个标题")

        return users

    def _launch_all_browsers(self, users):
        """
        阶段一：串行启动所有浏览器
        每个浏览器启动后访问about:blank检测存活，然后最小化
        """
        self._log("\n🚀 阶段一：串行启动浏览器...")
        self.browsers = {}

        for i, u in enumerate(users):
            if self.stop_event.is_set():
                break
            while self.pause_event.is_set() and not self.stop_event.is_set():
                time.sleep(0.1)

            task_id = u["user_index"] + 1

            try:
                self._log(f"   [{i+1}/{len(users)}] 启动浏览器 #{task_id}...")
                browser = safe_init_browser(task_id, self._log)
                self.browsers[task_id] = browser
                self._log(f"   ✅ 浏览器 #{task_id} 启动成功")

            except Exception as e:
                self._log(f"   ❌ 浏览器 #{task_id} 启动失败: {str(e)[:100]}")
                continue

            # 间隔等待，让桌面堆回收
            if i < len(users) - 1:
                self._log(f"   等待 6 秒，让系统回收资源...")
                time.sleep(6)

        self._log(f"\n📊 浏览器启动结果: {len(self.browsers)}/{len(users)} 成功")
        return len(self.browsers) > 0

    def run(self):
        self._log("=" * 60)
        self._log("🚀 Modrinth 收藏夹分配分析启动")
        self._log(f"   标题文件: {len(self.title_files)} 个")
        self._log(f"   简介文件: {len(self.intro_files)} 个")
        self._log(f"   输出目录: {self.output_dir}")
        self._log(f"   浏览器最大数: {self.thread_count}")
        self._log("=" * 60)

        # 1. 读取标题和简介
        self._log("\n📖 步骤1: 读取标题文件...")
        raw_titles = self._read_lines_from_files(self.title_files)
        self._log(f"   标题总行数: {len(raw_titles)}")

        self._log("\n📖 步骤2: 读取简介文件...")
        raw_intros = self._read_lines_from_files(self.intro_files)
        self._log(f"   简介总行数: {len(raw_intros)}")

        if not raw_titles or not raw_intros:
            self._log("\n❌ 标题或简介为空，无法继续")
            return

        # 2. 对齐标题和简介数量
        self._log("\n📊 步骤3: 对齐标题和简介数量...")
        title_count = len(raw_titles)
        intro_count = len(raw_intros)

        if title_count > intro_count:
            diff = title_count - intro_count
            self._log(f"   标题({title_count}) > 简介({intro_count})，追加 {diff} 个简介")
            extended_intros = raw_intros.copy()
            for i in range(diff):
                extended_intros.append(raw_intros[i % intro_count])
            raw_intros = extended_intros
        elif title_count < intro_count:
            diff = intro_count - title_count
            self._log(f"   标题({title_count}) < 简介({intro_count})，追加 {diff} 个标题")
            extended_titles = raw_titles.copy()
            for i in range(diff):
                extended_titles.append(raw_titles[i % title_count])
            raw_titles = extended_titles
        else:
            self._log(f"   标题({title_count}) = 简介({intro_count})，无需追加")

        total_folders = len(raw_titles)
        self._log(f"\n📊 最终对齐: 标题数={total_folders}, 简介数={len(raw_intros)}")
        self._log(f"   需要创建的收藏夹总数: {total_folders}")

        # 3. 分发标题到用户
        self._log("\n👤 步骤4: 按用户分发标题（每个用户最多32个，标题不重复）...")
        users = self._distribute_to_users(raw_titles)
        total_users = len(users)

        self._log(f"\n📊 分配结果汇总:")
        self._log(f"   总收藏夹数: {total_folders}")
        self._log(f"   总用户数: {total_users}")
        for u in users:
            self._log(f"   用户 #{u['user_index'] + 1}: {len(u['titles'])} 个收藏夹")

        # 4. 生成分配方案文件
        self._log("\n💾 步骤5: 生成分配方案文件...")
        plan_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plans")
        os.makedirs(plan_dir, exist_ok=True)

        output_lines = []
        output_lines.append("=" * 60)
        output_lines.append("Modrinth 收藏夹分配方案")
        output_lines.append("=" * 60)
        output_lines.append(f"总收藏夹数: {total_folders}")
        output_lines.append(f"总用户数: {total_users}")
        output_lines.append(f"每个用户最多收藏夹: {self.MAX_PER_USER}")
        output_lines.append("=" * 60)
        output_lines.append("")

        global_collection_idx = 0
        for u in users:
            output_lines.append(f"--- 用户 #{u['user_index'] + 1} ({len(u['titles'])} 个收藏夹) ---")
            for idx, t in enumerate(u["titles"], 1):
                intro = raw_intros[global_collection_idx] if global_collection_idx < len(raw_intros) else ""
                output_lines.append(f"  收藏夹 {idx}: 标题={t} | 简介={intro}")
                global_collection_idx += 1
            output_lines.append("")

        plan_path = os.path.join(plan_dir, f"collection_plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(plan_path, "w", encoding="utf-8") as f:
            f.write("\n".join(output_lines))
        self._log(f"   分配方案: {plan_path}")

        # 5. 阶段一：串行启动浏览器
        if not self._launch_all_browsers(users):
            self._log("❌ 没有浏览器成功启动，任务终止")
            return

        # 6. 阶段二：并发执行注册流程
        self._log("\n🚀 阶段二：并发执行注册流程...")
        self._log(f"   并发数: {self.thread_count}")
        self._log(f"   已启动浏览器: {len(self.browsers)} 个")

        completed = 0
        executor = ThreadPoolExecutor(max_workers=self.thread_count)
        futures = {}

        try:
            for u in users:
                if self.stop_event.is_set():
                    self._log("   收到停止信号，终止提交新任务")
                    break

                task_id = u["user_index"] + 1
                if task_id not in self.browsers:
                    self._log(f"   [跳过] 用户 #{task_id} - 浏览器未启动")
                    continue

                user_titles = u["titles"]
                start_idx = sum(len(users[i]["titles"]) for i in range(u["user_index"]))
                user_intros = [raw_intros[start_idx + idx] if start_idx + idx < len(raw_intros) else ""
                              for idx in range(len(user_titles))]

                self._log(f"   [提交] 用户 #{task_id} - {len(user_titles)} 个收藏夹")
                future = executor.submit(
                    single_user_task,
                    task_id=task_id,
                    user_titles=user_titles,
                    user_intros=user_intros,
                    output_dir=self.output_dir,
                    log_callback=self.log_callback,
                    browser_tuple=self.browsers[task_id],
                    stop_event=self.stop_event,
                    pause_event=self.pause_event
                )
                futures[future] = task_id

            self._log(f"   已提交 {len(futures)} 个任务，等待执行...")

            for future in as_completed(futures):
                if self.stop_event.is_set():
                    self._log("   收到停止信号，终止等待")
                    break
                while self.pause_event.is_set() and not self.stop_event.is_set():
                    time.sleep(0.1)

                task_id = futures[future]
                try:
                    result = future.result(timeout=600)
                    self._log(f"   [完成] 用户 #{task_id}: {result}")
                except Exception as e:
                    self._log(f"   [错误] 用户 #{task_id}: {str(e)[:200]}")

                completed += 1
                if self.progress_callback:
                    self.progress_callback({
                        "current": completed,
                        "total": len(futures),
                        "status": f"已完成 {completed}/{len(futures)} 个用户"
                    })

        finally:
            self._log("   正在关闭线程池...")
            executor.shutdown(wait=False)

            # 清理残留浏览器（带try/except防止双重quit报错）
            for task_id, browser_info in list(self.browsers.items()):
                try:
                    driver, _, _, user_data_dir = browser_info
                    try:
                        driver.quit()
                    except:
                        pass
                    if user_data_dir and os.path.exists(user_data_dir):
                        try:
                            shutil.rmtree(user_data_dir, ignore_errors=True)
                        except:
                            pass
                except:
                    pass
                # 移除引用，释放内存
                self.browsers.pop(task_id, None)

            self._log("   线程池已关闭")

        self._log("\n" + "=" * 60)
        self._log("✅ 全部完成!")
        self._log(f"   总收藏夹数: {total_folders}")
        self._log(f"   总用户数: {total_users}")
        self._log(f"   完成用户: {completed}/{len(futures)}")
        self._log(f"   分配方案: {plan_path}")
        self._log("=" * 60)

    def stop(self):
        self.stop_event.set()

    def pause(self):
        self.pause_event.set()

    def resume(self):
        self.pause_event.clear()


def run_gui():
    root = tk.Tk()
    root.title("Modrinth 批量注册工具")
    root.geometry("1100x900")
    root.minsize(1000, 800)

    log_queue = queue.Queue()
    engine = [None]

    title_dir_var = tk.StringVar(value="")
    intro_dir_var = tk.StringVar(value="")
    output_dir_var = tk.StringVar(value="")
    thread_count_var = tk.StringVar(value=str(DEFAULT_THREAD_COUNT))
    title_list = []
    intro_list = []
    title_check_vars = {}
    intro_check_vars = {}

    def log(msg):
        log_queue.put(msg)

    def update_progress(data):
        log_queue.put(("progress", data))

    # 标题栏
    title_frame = tk.Frame(root, bg="#2c5aa0")
    title_frame.pack(fill=tk.X)
    tk.Label(title_frame, text="📝 Modrinth 批量注册工具", font=("微软雅黑", 16, "bold"),
             fg="white", bg="#2c5aa0", pady=12).pack()

    main = tk.Frame(root, padx=15, pady=10)
    main.pack(fill=tk.BOTH, expand=True)

    # 配置区
    cfg = tk.LabelFrame(main, text="配置选项", font=("微软雅黑", 10, "bold"))
    cfg.pack(fill=tk.X, pady=5)

    # 线程数
    thread_frame = tk.Frame(cfg)
    thread_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(thread_frame, text="并发数:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    thread_spin = tk.Spinbox(thread_frame, from_=1, to=8, textvariable=thread_count_var,
                              width=8, font=("微软雅黑", 10))
    thread_spin.pack(side=tk.LEFT, padx=5)
    tk.Label(thread_frame, text="（建议3-5，Windows无GPU环境串行启动后并发执行）", font=("微软雅黑", 9), fg="#666").pack(side=tk.LEFT)

    # 标题目录
    title_dir_frame = tk.Frame(cfg)
    title_dir_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(title_dir_frame, text="标题目录:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    tk.Entry(title_dir_frame, textvariable=title_dir_var, width=50, font=("微软雅黑", 9), state="readonly").pack(side=tk.LEFT, padx=5)

    def choose_title_dir():
        d = filedialog.askdirectory(title="选择标题文件所在目录")
        if d:
            title_dir_var.set(d)
            refresh_title_list(d)

    tk.Button(title_dir_frame, text="浏览...", command=choose_title_dir,
              font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)

    # 简介目录
    intro_dir_frame = tk.Frame(cfg)
    intro_dir_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(intro_dir_frame, text="简介目录:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    tk.Entry(intro_dir_frame, textvariable=intro_dir_var, width=50, font=("微软雅黑", 9), state="readonly").pack(side=tk.LEFT, padx=5)

    def choose_intro_dir():
        d = filedialog.askdirectory(title="选择简介文件所在目录")
        if d:
            intro_dir_var.set(d)
            refresh_intro_list(d)

    tk.Button(intro_dir_frame, text="浏览...", command=choose_intro_dir,
              font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)

    # 输出目录
    output_dir_frame = tk.Frame(cfg)
    output_dir_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(output_dir_frame, text="输出目录:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    tk.Entry(output_dir_frame, textvariable=output_dir_var, width=50, font=("微软雅黑", 9), state="readonly").pack(side=tk.LEFT, padx=5)

    def choose_output_dir():
        d = filedialog.askdirectory(title="选择结果文件存放目录")
        if d:
            output_dir_var.set(d)

    tk.Button(output_dir_frame, text="浏览...", command=choose_output_dir,
              font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)

    # 文件选择区
    files_frame = tk.Frame(main)
    files_frame.pack(fill=tk.X, pady=5)

    # 标题文件列表
    title_list_frame = tk.LabelFrame(files_frame, text="标题文件列表（勾选添加）", font=("微软雅黑", 10, "bold"), height=200)
    title_list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
    title_list_frame.pack_propagate(False)

    title_canvas = tk.Canvas(title_list_frame, bg="#1e1e1e", highlightthickness=0)
    title_scrollbar = tk.Scrollbar(title_list_frame, orient=tk.VERTICAL, command=title_canvas.yview)
    title_scrollable_frame = tk.Frame(title_canvas, bg="#1e1e1e")

    title_scrollable_frame.bind(
        "<Configure>",
        lambda e: title_canvas.configure(scrollregion=title_canvas.bbox("all"))
    )
    title_canvas.create_window((0, 0), window=title_scrollable_frame, anchor="nw", width=480)
    title_canvas.configure(yscrollcommand=title_scrollbar.set)
    title_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    title_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # 简介文件列表
    intro_list_frame = tk.LabelFrame(files_frame, text="简介文件列表（勾选添加）", font=("微软雅黑", 10, "bold"), height=200)
    intro_list_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
    intro_list_frame.pack_propagate(False)

    intro_canvas = tk.Canvas(intro_list_frame, bg="#1e1e1e", highlightthickness=0)
    intro_scrollbar = tk.Scrollbar(intro_list_frame, orient=tk.VERTICAL, command=intro_canvas.yview)
    intro_scrollable_frame = tk.Frame(intro_canvas, bg="#1e1e1e")

    intro_scrollable_frame.bind(
        "<Configure>",
        lambda e: intro_canvas.configure(scrollregion=intro_canvas.bbox("all"))
    )
    intro_canvas.create_window((0, 0), window=intro_scrollable_frame, anchor="nw", width=480)
    intro_canvas.configure(yscrollcommand=intro_scrollbar.set)
    intro_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    intro_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def refresh_title_list(directory):
        for widget in title_scrollable_frame.winfo_children():
            widget.destroy()
        title_check_vars.clear()
        title_list.clear()
        title_input.delete("1.0", tk.END)

        if not directory or not os.path.isdir(directory):
            tk.Label(title_scrollable_frame, text="请先选择有效目录", font=("微软雅黑", 10),
                     bg="#1e1e1e", fg="#888888").pack(pady=20)
            return

        txt_files = sorted([f for f in os.listdir(directory) if f.lower().endswith(".txt")])
        if not txt_files:
            tk.Label(title_scrollable_frame, text="目录中未找到 .txt 文件", font=("微软雅黑", 10),
                     bg="#1e1e1e", fg="#888888").pack(pady=20)
            return

        for fname in txt_files:
            display_name = os.path.splitext(fname)[0]
            var = tk.IntVar(value=0)
            full_path = os.path.join(directory, fname)
            title_check_vars[display_name] = (var, full_path)

            cb = tk.Checkbutton(
                title_scrollable_frame,
                text=f"  {display_name}",
                variable=var,
                font=("微软雅黑", 10),
                fg="white",
                bg="#1e1e1e",
                selectcolor="#333333",
                activebackground="#1e1e1e",
                activeforeground="white",
                anchor=tk.W,
                command=lambda dn=display_name: on_title_toggle(dn)
            )
            cb.pack(fill=tk.X, padx=5, pady=2)

    def on_title_toggle(display_name):
        var, full_path = title_check_vars[display_name]
        if var.get() == 1:
            if full_path not in title_list:
                title_list.append(full_path)
        else:
            if full_path in title_list:
                title_list.remove(full_path)
        title_input.delete("1.0", tk.END)
        title_input.insert(tk.END, "\n".join(title_list))

    def refresh_intro_list(directory):
        for widget in intro_scrollable_frame.winfo_children():
            widget.destroy()
        intro_check_vars.clear()
        intro_list.clear()
        intro_input.delete("1.0", tk.END)

        if not directory or not os.path.isdir(directory):
            tk.Label(intro_scrollable_frame, text="请先选择有效目录", font=("微软雅黑", 10),
                     bg="#1e1e1e", fg="#888888").pack(pady=20)
            return

        txt_files = sorted([f for f in os.listdir(directory) if f.lower().endswith(".txt")])
        if not txt_files:
            tk.Label(intro_scrollable_frame, text="目录中未找到 .txt 文件", font=("微软雅黑", 10),
                     bg="#1e1e1e", fg="#888888").pack(pady=20)
            return

        for fname in txt_files:
            display_name = os.path.splitext(fname)[0]
            var = tk.IntVar(value=0)
            full_path = os.path.join(directory, fname)
            intro_check_vars[display_name] = (var, full_path)

            cb = tk.Checkbutton(
                intro_scrollable_frame,
                text=f"  {display_name}",
                variable=var,
                font=("微软雅黑", 10),
                fg="white",
                bg="#1e1e1e",
                selectcolor="#333333",
                activebackground="#1e1e1e",
                activeforeground="white",
                anchor=tk.W,
                command=lambda dn=display_name: on_intro_toggle(dn)
            )
            cb.pack(fill=tk.X, padx=5, pady=2)

    def on_intro_toggle(display_name):
        var, full_path = intro_check_vars[display_name]
        if var.get() == 1:
            if full_path not in intro_list:
                intro_list.append(full_path)
        else:
            if full_path in intro_list:
                intro_list.remove(full_path)
        intro_input.delete("1.0", tk.END)
        intro_input.insert(tk.END, "\n".join(intro_list))

    # 输入框区域
    input_frame = tk.Frame(main)
    input_frame.pack(fill=tk.X, pady=5)

    title_input_frame = tk.LabelFrame(input_frame, text="已选标题文件路径", font=("微软雅黑", 10, "bold"))
    title_input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

    title_input = tk.Text(title_input_frame, font=("Consolas", 9), wrap=tk.WORD,
                          height=3, bg="#1e1e1e", fg="#d4d4d4", insertbackground="white")
    title_input.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    title_input_scroll = tk.Scrollbar(title_input_frame, command=title_input.yview)
    title_input_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    title_input.config(yscrollcommand=title_input_scroll.set)

    intro_input_frame = tk.LabelFrame(input_frame, text="已选简介文件路径", font=("微软雅黑", 10, "bold"))
    intro_input_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

    intro_input = tk.Text(intro_input_frame, font=("Consolas", 9), wrap=tk.WORD,
                          height=3, bg="#1e1e1e", fg="#d4d4d4", insertbackground="white")
    intro_input.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    intro_input_scroll = tk.Scrollbar(intro_input_frame, command=intro_input.yview)
    intro_input_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    intro_input.config(yscrollcommand=intro_input_scroll.set)

    # 控制按钮
    btn_frame = tk.Frame(main)
    btn_frame.pack(fill=tk.X, pady=10)

    start_btn = tk.Button(btn_frame, text="▶ 启动", bg="#4CAF50", fg="white",
                          font=("微软雅黑", 12, "bold"), width=15, height=1)
    start_btn.pack(side=tk.LEFT, padx=5)

    pause_btn = tk.Button(btn_frame, text="⏸ 暂停", bg="#FF9800", fg="white",
                          font=("微软雅黑", 12, "bold"), width=12, height=1, state=tk.DISABLED)
    pause_btn.pack(side=tk.LEFT, padx=5)

    # 进度区
    prog_frame = tk.LabelFrame(main, text="处理进度", font=("微软雅黑", 10, "bold"))
    prog_frame.pack(fill=tk.X, pady=5)

    task_label = tk.Label(prog_frame, text="就绪", font=("微软雅黑", 11, "bold"),
                          fg="#333", anchor=tk.W)
    task_label.pack(fill=tk.X, padx=10, pady=5)

    progress_frame = tk.Frame(prog_frame)
    progress_frame.pack(fill=tk.X, padx=10, pady=2)
    tk.Label(progress_frame, text="用户进度:", font=("微软雅黑", 9), width=12, anchor=tk.W).pack(side=tk.LEFT)
    progress_var = tk.DoubleVar(value=0)
    progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, length=750)
    progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
    progress_label = tk.Label(progress_frame, text="0/0", font=("微软雅黑", 9), width=8)
    progress_label.pack(side=tk.LEFT, padx=5)

    stats_label = tk.Label(prog_frame, text="已处理: 0 个用户 | 状态: 就绪",
                           font=("微软雅黑", 9), fg="#666", anchor=tk.W)
    stats_label.pack(fill=tk.X, padx=10, pady=5)

    # 日志区
    log_frame = tk.LabelFrame(main, text="运行日志", font=("微软雅黑", 10, "bold"))
    log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    log_text = scrolledtext.ScrolledText(log_frame, font=("Consolas", 9), wrap=tk.WORD,
                                         state=tk.DISABLED, bg="#1e1e1e", fg="#d4d4d4",
                                         insertbackground="white")
    log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    status_bar = tk.Label(root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=("微软雅黑", 9))
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def add_log(msg):
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)

    def refresh():
        try:
            while True:
                item = log_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "progress":
                    d = item[1]
                    pct = (d["current"] / d["total"]) * 100 if d["total"] > 0 else 0
                    progress_var.set(pct)
                    progress_label.config(text=f"{d['current']}/{d['total']}")
                    task_label.config(text=f"状态: {d['status']}")
                    stats_label.config(text=f"已处理: {d['current']} 个用户 | 状态: {d['status']}")
                else:
                    add_log(item)
        except queue.Empty:
            pass
        root.after(200, refresh)

    def start_processing():
        if not title_dir_var.get():
            messagebox.showerror("错误", "请选择标题目录")
            return
        if not intro_dir_var.get():
            messagebox.showerror("错误", "请选择简介目录")
            return
        if not output_dir_var.get():
            messagebox.showerror("错误", "请选择结果文件存放目录")
            return
        if not title_list:
            messagebox.showerror("错误", "请至少勾选一个标题文件")
            return
        if not intro_list:
            messagebox.showerror("错误", "请至少勾选一个简介文件")
            return

        out_dir = output_dir_var.get()
        os.makedirs(out_dir, exist_ok=True)

        try:
            tc = int(thread_count_var.get())
            if tc > 8:
                tc = 5  # 限制最大5
            if tc < 1:
                tc = 1
        except ValueError:
            tc = DEFAULT_THREAD_COUNT

        engine[0] = ModrinthCollector(
            title_files=title_list.copy(),
            intro_files=intro_list.copy(),
            output_dir=out_dir,
            thread_count=tc,
            log_callback=log,
            progress_callback=update_progress
        )
        start_btn.config(state=tk.DISABLED)
        pause_btn.config(state=tk.NORMAL)
        status_bar.config(text="处理中...")
        Thread(target=lambda: engine[0].run(), daemon=True).start()

    def pause_processing():
        if not engine[0]:
            return
        if pause_btn.cget("text") == "⏸ 暂停":
            engine[0].pause()
            pause_btn.config(text="▶ 继续")
            status_bar.config(text="已暂停")
        else:
            engine[0].resume()
            pause_btn.config(text="⏸ 暂停")
            status_bar.config(text="处理中...")

    start_btn.config(command=start_processing)
    pause_btn.config(command=pause_processing)

    add_log("Modrinth 批量注册工具已启动")
    add_log("修复：Windows无GPU环境下多开浏览器卡死问题")
    add_log("策略：串行启动浏览器 → 最小化 → 并发执行注册流程")
    add_log("请依次选择：标题目录 -> 简介目录 -> 输出目录")
    add_log("勾选需要的文件后，点击「启动」开始")
    refresh()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
