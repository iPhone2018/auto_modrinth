#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modrinth 批量注册 + 收藏夹管理工具 - GUI版
"""

import os
import sys
import time
import random
import string
import socket
import json
import requests
import threading
import queue
import warnings
import unicodedata
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
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from urllib.parse import quote
from selenium.common.exceptions import TimeoutException, NoSuchElementException

warnings.filterwarnings("ignore", category=InsecureRequestWarning)

MAX_COLLECTIONS_PER_USER = 32
_file_write_lock = threading.Lock()

_active_drivers = {}
_drivers_lock = threading.Lock()


def register_driver(task_id, driver):
    with _drivers_lock:
        _active_drivers[task_id] = driver


def unregister_driver(task_id):
    with _drivers_lock:
        _active_drivers.pop(task_id, None)


def close_all_drivers():
    """立刻全部quit所有浏览器"""
    with _drivers_lock:
        drivers = list(_active_drivers.values())
        _active_drivers.clear()
    for driver in drivers:
        try:
            driver.quit()
        except Exception:
            pass


def _is_port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(('127.0.0.1', port)) == 0


def _cleanup_chrome_locks(user_data_dir: str):
    lock_files = ["SingletonLock", "SingletonSocket", "SingletonCookie"]
    for lock_name in lock_files:
        lock_path = os.path.join(user_data_dir, lock_name)
        if os.path.exists(lock_path):
            try:
                if os.path.isfile(lock_path):
                    os.remove(lock_path)
                elif os.path.islink(lock_path):
                    os.unlink(lock_path)
            except Exception:
                pass
    cache_dirs = ["GPUCache", "Code Cache", "Service Worker"]
    for cache_name in cache_dirs:
        cache_path = os.path.join(user_data_dir, cache_name)
        if os.path.exists(cache_path):
            try:
                import shutil
                shutil.rmtree(cache_path, ignore_errors=True)
            except Exception:
                pass


def _find_available_port(start_port: int, max_attempts: int = 20) -> int:
    for offset in range(max_attempts):
        port = start_port + offset
        if not _is_port_in_use(port):
            return port
    raise RuntimeError(f"无法找到可用端口，已尝试 {max_attempts} 个端口（从 {start_port} 开始）")


# def get_local_chromedriver_path(base_dir: str) -> str:
#     """读取项目.wdm缓存里已下载的mac-arm64 chromedriver，自动选最新版本"""
#     wdm_root = Path(base_dir) / ".wdm" / "drivers" / "chromedriver" / "mac-arm64"
#     if not wdm_root.exists():
#         raise FileNotFoundError(f"未找到wdm驱动缓存目录：{wdm_root}\n请先下载对应版本chromedriver")
#
#     version_dirs = []
#     for child in wdm_root.iterdir():
#         if child.is_dir() and child.name.count(".") >= 3:
#             try:
#                 ver_tuple = tuple(map(int, child.name.split(".")))
#                 version_dirs.append((ver_tuple, child))
#             except ValueError:
#                 continue
#
#     if not version_dirs:
#         raise FileNotFoundError("wdm缓存内无任何chromedriver版本文件夹")
#
#     version_dirs.sort(reverse=True, key=lambda x: x[0])
#     latest_driver = version_dirs[0][1] / "chromedriver-mac-arm64" / "chromedriver"
#
#     if not latest_driver.exists():
#         raise FileNotFoundError(f"驱动文件缺失：{latest_driver}")
#     return str(latest_driver)
#
#
# def find_mac_system_chrome() -> str | None:
#     """自动查找Mac系统自带Chrome可执行文件路径"""
#     chrome_candidates = [
#         "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
#         "/Applications/Google Chrome Beta.app/Contents/MacOS/Google Chrome Beta",
#         "/Applications/Google Chrome Dev.app/Contents/MacOS/Google Chrome Dev",
#         "/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary",
#     ]
#     for path in chrome_candidates:
#         if Path(path).exists():
#             return path
#     return None
#
#
# def init_browser(task_id: int):
#     """
#     跨平台浏览器初始化
#     Mac：自动读取系统Chrome + 复用项目.wdm缓存驱动
#     Windows：使用根目录chromedriver.exe + 项目内chrome便携包
#     """
#     if getattr(sys, 'frozen', False):
#         base_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
#     else:
#         base_dir = os.path.dirname(os.path.abspath(__file__))
#
#     if sys.platform == "win32":
#         driver_name = "chromedriver.exe"
#         local_driver = os.path.join(base_dir, driver_name)
#         chrome_bin_path = os.path.join(base_dir, "chrome", "chrome.exe")
#
#         if not os.path.exists(local_driver):
#             raise FileNotFoundError(f"找不到 chromedriver.exe，放在程序同级目录\n{local_driver}")
#         if not os.path.exists(chrome_bin_path):
#             raise FileNotFoundError(f"找不到便携chrome.exe，chrome文件夹放程序同级\n{chrome_bin_path}")
#     else:
#         local_driver = get_local_chromedriver_path(base_dir)
#         chrome_bin_path = find_mac_system_chrome()
#         if chrome_bin_path is None:
#             portable_chrome = os.path.join(base_dir, "chrome", "Chrome")
#             if os.path.exists(portable_chrome):
#                 chrome_bin_path = portable_chrome
#             else:
#                 raise FileNotFoundError(
#                     "未检测到系统Google Chrome，同时项目内无便携chrome文件夹\n"
#                     "解决方法：1.安装系统Chrome；2.在程序根目录放置chrome便携文件夹"
#                 )
#         os.chmod(local_driver, 0o755)
#
#     options = webdriver.ChromeOptions()
#     options.binary_location = chrome_bin_path
#     options.add_argument("--no-sandbox")
#     options.add_argument("--disable-dev-shm-usage")
#     options.add_argument("--disable-extensions")
#     options.add_argument("--disable-background-networking")
#     options.add_argument("--start-maximized")
#     options.add_experimental_option("excludeSwitches", ["enable-logging"])
#     options.add_argument("--disable-animations")
#     options.add_experimental_option("excludeSwitches", ["enable-automation"])
#     options.add_experimental_option("useAutomationExtension", False)
#
#     # ===== 窗口配置：全屏最大化 =====
#     options.add_argument("--start-maximized")
#     # =====================================================
#
#     user_data_dir = os.path.join(base_dir, f"chrome_user_data_task_{task_id}")
#     os.makedirs(user_data_dir, exist_ok=True)
#     options.add_argument(f"--user-data-dir={user_data_dir}")
#     debug_port = 9222 + task_id * 10
#     options.add_argument(f"--remote-debugging-port={debug_port}")
#
#     service = Service(local_driver)
#     driver = webdriver.Chrome(service=service, options=options)
#     driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
#         "source": """
#             Object.defineProperty(navigator, 'webdriver', {get: () => undefined})
#         """
#     })
#     wait = WebDriverWait(driver, 15)
#     short_wait = WebDriverWait(driver, 3)
#     return driver, wait, short_wait


def init_browser(task_id):
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

    options = webdriver.ChromeOptions()
    options.binary_location = portable_chrome
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-background-networking")

    user_data_dir = os.path.join(base_dir, f"chrome_user_data_task_{task_id}")
    os.makedirs(user_data_dir, exist_ok=True)
    _cleanup_chrome_locks(user_data_dir)
    options.add_argument(f"--user-data-dir={user_data_dir}")

    debug_port = _find_available_port(9222 + task_id)
    options.add_argument(f"--remote-debugging-port={debug_port}")

    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--disable-animations")

    driver = webdriver.Chrome(service=Service(local_driver), options=options)
    wait = WebDriverWait(driver, 15)
    short_wait = WebDriverWait(driver, 3)
    return driver, wait, short_wait


def retry_click(driver, element, max_retries=3, delay=0.5):
    from selenium.webdriver.common.action_chains import ActionChains
    for attempt in range(max_retries):
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});", element
            )
            time.sleep(0.2)
            element.click()
            return True
        except Exception as e1:
            try:
                ActionChains(driver).move_to_element(element).click().perform()
                return True
            except Exception as e2:
                try:
                    driver.execute_script("arguments[0].click();", element)
                    return True
                except Exception as e3:
                    time.sleep(delay)
    return False


def display_width(text):
    return sum(2 if unicodedata.east_asian_width(c) in ('F', 'W') else 1 for c in str(text or ''))


def auto_fit_columns(ws, min_w=8, max_w=50, padding=3):
    for col_cells in ws.columns:
        letter = col_cells[0].column_letter
        w = max((display_width(c.value) for c in col_cells
                 if not isinstance(c, openpyxl.cell.cell.MergedCell) and c.value is not None), default=0)
        ws.column_dimensions[letter].width = max(min_w, min(w * 1.1 + padding, max_w))


def append_link_to_txt(link: str, file_path: str = "links.txt"):
    with _file_write_lock:
        with open(file_path, "a", encoding="utf-8") as f:
            f.write(link + "\n")


# === Modified：增加stop_event，子任务可以接收停止信号 ===
def single_user_task(task_id: int, email: str, user_titles: list, user_intros: list,
                     output_dir: str, interval: float, stop_event: Event, log_callback=None,
                     on_collection_created=None):
    driver = None
    session = None
    token = None
    success_count = 0

    def _check_stop():
        """一旦收到停止信号直接抛出，跳出所有逻辑，进入finally"""
        if stop_event.is_set():
            raise InterruptedError("收到停止信号，立即终止当前用户会话")

    try:
        _check_stop()
        if log_callback:
            log_callback(f"[用户{task_id}] ===== 任务开始 =====")
            log_callback(f"[用户{task_id}] 需要创建收藏夹: {len(user_titles)} 个")
            log_callback(f"[用户{task_id}] 启动浏览器，准备注册...")

        _check_stop()
        if log_callback:
            log_callback(f"[用户{task_id}] 初始化 Chrome 浏览器...")
        driver, wait, short_wait = init_browser(task_id)
        register_driver(task_id, driver)
        _check_stop()
        if log_callback:
            log_callback(f"[用户{task_id}] 浏览器初始化成功")
        long_wait = WebDriverWait(driver, 6000)

        driver.get("https://modrinth.com")
        _check_stop()
        # 等待header区域出现，代表桌面版导航栏渲染
        long_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "header.desktop-only")))

        # 精准定位：header.desktop-only 右侧，带log‑in图标的Sign in按钮
        sign_in_btn = long_wait.until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                'header.desktop-only > div.flex.items-center.gap-1 a[data-button]:has(svg.lucide-log-in)'
            ))
        )

        # 优先JS点击，绕过selenium合成鼠标事件
        driver.execute_script("arguments[0].click();", sign_in_btn)
        _check_stop()

        email_input = long_wait.until(EC.visibility_of_element_located((By.ID, "email")))
        email_input.clear()
        email_input.send_keys(email)
        if log_callback:
            log_callback(f"[用户{task_id}] 输入邮箱: {email}")

        pwd_input = long_wait.until(EC.visibility_of_element_located((By.ID, "password")))
        pwd_input.clear()
        pwd_input.send_keys("zqx1314520.")
        if log_callback:
            log_callback(f"[用户{task_id}] 输入密码")

        # 1.等待hCaptcha iframe出现（主文档）
        iframe = long_wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'iframe[src*="newassets.hcaptcha.com"][src*="frame=checkbox"]')
            )
        )

        # 切入iframe！非常关键
        driver.switch_to.frame(iframe)

        # 现在在iframe内部，才可以拿到 #checkbox
        checkbox = long_wait.until(
            EC.element_to_be_clickable((By.ID, "checkbox"))
        )

        # 滚动到可视区域（iframe内部滚动）
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            checkbox
        )
        time.sleep(0.4)

        print(f"[hCaptcha] 点击前 aria-checked: {checkbox.get_attribute('aria-checked')}")

        # 优先JS点击，iframe内ActionChains经常失效
        driver.execute_script("arguments[0].click();", checkbox)
        print("✅ [hCaptcha] iframe内JS点击checkbox完成")

        # 切回主页面，交给你后面的轮询检测逻辑
        driver.switch_to.default_content()
        print("[hCaptcha] 已切回主文档，开始轮询检测验证状态")

        print("\n⏳ 等待手动完成 hCaptcha 验证...")

        max_wait_time = 600
        poll_interval = 2
        elapsed = 0
        verified = False

        while elapsed < max_wait_time:
            _check_stop()  # 等待人机验证的循环也要检测停止信号
            time.sleep(poll_interval)
            elapsed += poll_interval

            try:
                resp_textareas = driver.find_elements(By.CSS_SELECTOR, "textarea[name='h-captcha-response']")
                if resp_textareas:
                    token_val = resp_textareas[0].get_attribute("value")
                    if token_val and len(token_val.strip()) > 20:
                        verified = True

                if verified:
                    if log_callback:
                        log_callback(f"[用户{task_id}] ✅ hCaptcha 检测到有效token，验证通过!")
                    break

            except Exception:
                try:
                    driver.switch_to.default_content()
                except:
                    pass
        else:
            raise TimeoutError("hCaptcha 验证等待超时（5分钟未检测到通过）")

        _check_stop()
        continue_btn = long_wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//button[normalize-space()='Continue with Email']")
            )
        )

        for _ in range(20):
            _check_stop()
            aria_disabled = continue_btn.get_attribute("aria-disabled")
            if aria_disabled != "true":
                break
            time.sleep(0.25)
        else:
            raise TimeoutError("Continue with Email 按钮长时间处于禁用状态，hCaptcha可能未生效")

        driver.execute_script("arguments[0].click();", continue_btn)
        if log_callback:
            log_callback(f"[用户{task_id}] ✅ 点击 Continue with Email")

        time.sleep(5)
        _check_stop()

        for attempt in range(5):
            _check_stop()
            cookies = driver.get_cookies()
            for ck in cookies:
                if ck["name"] == "auth-token":
                    token = ck["value"]
                    break
            if token:
                if log_callback:
                    log_callback(f"[用户{task_id}] 获取 Token 成功")
                break
            if attempt < 4:
                time.sleep(5)
        else:
            raise Exception("无法获取 auth-token")

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
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
        })

        collection_ids = []
        # === Modified：创建收藏夹循环内每次迭代检测停止信号 ===
        for i, (title, intro) in enumerate(zip(user_titles, user_intros)):
            _check_stop()
            if log_callback:
                log_callback(f"[用户{task_id}] 创建收藏夹 {i + 1}/{len(user_titles)}: {title[:30]}...")

            create_payload = {
                "name": title,
                "description": intro,
                "projects": []
            }
            resp = session.post("https://api.modrinth.com/v3/collection", json=create_payload)
            time.sleep(random.uniform(0.1, 0.5))
            _check_stop()

            if resp.status_code == 200:
                collection_id = resp.json()["id"]
                collection_ids.append(collection_id)
                success_count += 1
                if on_collection_created:
                    on_collection_created(title)
                if log_callback:
                    log_callback(f"[用户{task_id}] 收藏夹创建成功! ID: {collection_id}")
                time.sleep(interval)
            else:
                if log_callback:
                    log_callback(f"[用户{task_id}] 创建收藏夹失败: {resp.status_code} - {resp.text}")

        _check_stop()
        if log_callback:
            log_callback(f"[用户{task_id}] 搜索热门模组...")
        search_resp = session.get(
            "https://api.modrinth.com/v2/search",
            params={"limit": 20, "index": "relevance", "new_filters": "project_types = `mod`"}
        )
        time.sleep(random.uniform(0.1, 0.5))
        _check_stop()

        if search_resp.status_code == 200:
            hits = search_resp.json().get("hits", [])
            if hits:
                target_id = hits[0]['project_id']
                session.post(f"https://api.modrinth.com/v2/project/{target_id}/follow")
                if log_callback:
                    log_callback(f"[用户{task_id}] 已关注项目: {target_id}")

                for cid in collection_ids:
                    _check_stop()
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

        _check_stop()
        if log_callback:
            log_callback(f"[用户{task_id}] 全部完成! 创建了 {success_count} 个收藏夹")
        return f"用户{task_id} 成功 {success_count}/{len(user_titles)}"

    except InterruptedError as ie:
        msg = f"[用户{task_id}] ✋ {str(ie)}"
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
        return f"用户{task_id} 被外部停止: {str(ie)}"
    except Exception as e:
        error_msg = f"[用户{task_id}] 错误: {str(e)}"
        if log_callback:
            log_callback(error_msg)
        else:
            print(error_msg)
        return f"用户{task_id} 失败: {str(e)}"
    finally:
        unregister_driver(task_id)
        # === 无论成功、异常、被中断，都执行登出会话 ===
        if session and token:
            try:
                session.delete(f"https://api.modrinth.com/v2/session/{token}", timeout=3)
            except Exception:
                pass
        if session:
            try:
                session.close()
            except Exception:
                pass
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        try:
            base_dir = os.path.dirname(os.path.realpath(sys.argv[0])) if getattr(sys, 'frozen',
                                                                                 False) else os.path.dirname(
                os.path.abspath(__file__))
            user_data_dir = os.path.join(base_dir, f"chrome_user_data_task_{task_id}")
            _cleanup_chrome_locks(user_data_dir)
        except Exception:
            pass


class ModrinthCollector:
    MAX_PER_USER = 32

    def __init__(self, title_files, intro_files, email_list, output_dir, thread_count, interval=0,
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
        self.completed_collections = 0
        self._is_running = False
        self.email_list = email_list
        self.interval = interval
        self.email_index = 0

    def _log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_msg = f"[{timestamp}] {message}"
        print(log_msg)
        if self.log_callback:
            self.log_callback(log_msg)

    def run(self):
        with self.lock:
            if self._is_running:
                self._log("已有任务在运行中，跳过")
                return
            self._is_running = True
            self.stop_event.clear()
            self.pause_event.clear()

        try:
            self._log("=" * 60)
            self._log("🚀 Modrinth 无限循环创建启动")
            self._log(f"   标题文件: {len(self.title_files)} 个")
            self._log(f"   简介文件: {len(self.intro_files)} 个")
            self._log(f"   输出目录: {self.output_dir}")
            self._log(f"   浏览器最大数: {self.thread_count}")
            self._log("=" * 60)

            self._log("\n📖 读取标题池...")
            title_pool = []
            for fp in self.title_files:
                if not os.path.exists(fp):
                    self._log(f"⚠️ 标题文件不存在，跳过: {fp}")
                    continue
                try:
                    with open(fp, "r", encoding="utf-8") as f:
                        for line in f:
                            line = line.strip()
                            if line:
                                title_pool.append(line)
                    self._log(f"   从 [{os.path.basename(fp)}] 读取")
                except Exception as e:
                    self._log(f"⚠️ 读取标题文件失败 [{fp}]: {e}")
            title_pool = list(dict.fromkeys(title_pool))
            self._log(f"   标题池去重后: {len(title_pool)} 个")

            if not title_pool:
                self._log("\n❌ 标题池为空，无法继续")
                return

            self._log("\n📖 读取简介池...")
            intro_files_data = []
            for fp in self.intro_files:
                lines = []
                if not os.path.exists(fp):
                    self._log(f"⚠️ 简介文件不存在，跳过: {fp}")
                    continue
                try:
                    with open(fp, "r", encoding="utf-8") as f:
                        for line in f:
                            line = line.strip()
                            if line:
                                lines.append(line)
                    self._log(f"   从 [{os.path.basename(fp)}] 读取 {len(lines)} 行")
                except Exception as e:
                    self._log(f"⚠️ 读取简介文件失败 [{fp}]: {e}")
                if lines:
                    intro_files_data.append(lines)

            if not intro_files_data:
                self._log("\n❌ 没有有效的简介文件，无法继续")
                return

            self._log(f"   简介文件数: {len(intro_files_data)}")

            if not self.email_list:
                self._log("\n❌ 邮箱列表为空，无法继续")
                return
            self._log("\n💾 生成分配方案文件...")
            plan_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plans")
            os.makedirs(plan_dir, exist_ok=True)

            output_lines = []
            output_lines.append("=" * 60)
            output_lines.append("Modrinth 收藏夹分配方案")
            output_lines.append("=" * 60)
            output_lines.append(f"标题池数量: {len(title_pool)}")
            output_lines.append(f"简介文件数: {len(intro_files_data)}")
            output_lines.append(f"每个用户最多收藏夹: {self.MAX_PER_USER}")
            output_lines.append("=" * 60)
            output_lines.append("")

            plan_path = os.path.join(plan_dir, f"collection_plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            with open(plan_path, "w", encoding="utf-8") as f:
                f.write("\n".join(output_lines))
            self._log(f"   分配方案: {plan_path}")

            self._log("\n🚀 开始无限循环创建...")
            self._log(f"   并发浏览器: {self.thread_count}")

            def on_collection_created(title):
                with self.lock:
                    self.completed_collections += 1
                if self.progress_callback:
                    self.progress_callback({
                        "current": self.completed_collections,
                        "status": f"已完成 {self.completed_collections} 个收藏夹"
                    })

            if self.progress_callback:
                self.progress_callback({
                    "current": self.completed_collections,
                    "status": f"已完成 {self.completed_collections} 个收藏夹"
                })

            with ThreadPoolExecutor(max_workers=self.thread_count) as executor:
                futures = {}
                next_user_idx = 0

                while not self.stop_event.is_set():
                    done_futures = [f for f in list(futures.keys()) if f.done()]
                    for f in done_futures:
                        user_idx = futures.pop(f)
                        try:
                            result = f.result()
                            self._log(f"   [完成] 用户 #{user_idx}: {result}")
                        except Exception as e:
                            self._log(f"   [错误] 用户 #{user_idx}: {str(e)}")

                    if self.pause_event.is_set():
                        time.sleep(0.5)
                        continue

                    if len(futures) >= self.thread_count:
                        time.sleep(0.5)
                        continue

                    next_user_idx += 1
                    count = min(self.MAX_PER_USER, len(title_pool))
                    titles = random.sample(title_pool, count)

                    intros = []
                    for i in range(count):
                        parts = [random.choice(lines) for lines in intro_files_data]
                        intros.append("".join(parts))

                    email = self.email_list[self.email_index % len(self.email_list)]
                    self.email_index += 1

                    self._log(f"   [提交] 用户 #{next_user_idx} - {len(titles)} 个收藏夹")
                    # === Modified：把stop_event传给任务函数 ===
                    future = executor.submit(
                        single_user_task,
                        task_id=next_user_idx,
                        email=email,
                        interval=self.interval,
                        user_titles=titles,
                        user_intros=intros,
                        output_dir=self.output_dir,
                        stop_event=self.stop_event,
                        log_callback=self.log_callback,
                        on_collection_created=on_collection_created
                    )
                    futures[future] = next_user_idx
                    time.sleep(2)

                self._log("   收到停止信号，终止提交新任务")
                # 不能cancel正在运行的线程，设置event让子线程内部自行中断
                close_all_drivers()
                self._log("   已调用close_all_drivers，全部浏览器强制退出，子任务将自行执行会话登出")

            self._log("\n" + "=" * 60)
            self._log("✅ 任务结束")
            self._log(f"   已完成收藏夹: {self.completed_collections} 个")
            self._log("=" * 60)

        finally:
            with self.lock:
                self._is_running = False

    def stop(self):
        self.stop_event.set()
        self.pause_event.set()
        close_all_drivers()
        self._log("🛑 stop触发：设置停止事件，强制关闭所有浏览器，子任务会自动登出会话")

    def pause(self):
        """暂停等价于完全停止，移除原来复杂的resume重建逻辑"""
        self.stop()
        self._log("⏸ pause触发：已执行stop，全部会话退出，浏览器关闭")

    def resume(self):
        """resume不再做任何内部重启，交给GUI层重新run"""
        self.stop_event.clear()
        self.pause_event.clear()
        self._log("▶ resume准备，等待GUI调用run重新启动")


def load_emails_from_dir(email_dir: str):
    """遍历目录下所有txt，按行读取邮箱，去重后返回列表"""
    emails = []
    if not os.path.isdir(email_dir):
        return emails
    for fname in sorted(os.listdir(email_dir)):
        if not fname.lower().endswith(".txt"):
            continue
        fpath = os.path.join(email_dir, fname)
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and "@" in line:
                        emails.append(line)
        except Exception:
            continue
    # 去重并保持顺序
    seen = set()
    result = []
    for e in emails:
        if e not in seen:
            seen.add(e)
            result.append(e)
    return result


CONFIG_FILE = "modrinth_gui_config.json"


def load_gui_config():
    """加载上次界面配置"""
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_gui_config(cfg):
    """保存界面配置到json"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


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
    email_dir_var = tk.StringVar(value="")
    interval_var = tk.StringVar(value="5")
    title_list = []
    intro_list = []
    email_list = []
    title_check_vars = {}
    intro_check_vars = {}

    # ========== 加载历史配置 ==========
    saved_cfg = load_gui_config()
    title_dir_var.set(saved_cfg.get("title_dir", ""))
    intro_dir_var.set(saved_cfg.get("intro_dir", ""))
    output_dir_var.set(saved_cfg.get("output_dir", ""))
    email_dir_var.set(saved_cfg.get("email_dir", ""))
    interval_var.set(saved_cfg.get("interval", "5"))
    saved_thread = saved_cfg.get("thread_count", 3)
    # 新增：读取上次勾选的文件名（不带后缀）
    saved_selected_titles = saved_cfg.get("selected_title_files", [])
    saved_selected_intros = saved_cfg.get("selected_intro_files", [])

    def log(msg):
        log_queue.put(msg)

    def update_progress(data):
        log_queue.put(("progress", data))

    def on_closing():
        """窗口关闭：先保存当前界面配置，再stop引擎"""
        # 提取当前勾选的文件名（无后缀）用于持久化
        curr_selected_titles = []
        for disp_name, (var, _) in title_check_vars.items():
            if var.get() == 1:
                curr_selected_titles.append(disp_name)

        curr_selected_intros = []
        for disp_name, (var, _) in intro_check_vars.items():
            if var.get() == 1:
                curr_selected_intros.append(disp_name)

        current_config = {
            "title_dir": title_dir_var.get(),
            "intro_dir": intro_dir_var.get(),
            "output_dir": output_dir_var.get(),
            "email_dir": email_dir_var.get(),
            "interval": interval_var.get(),
            "thread_count": int(thread_spin.get()),
            "selected_title_files": curr_selected_titles,
            "selected_intro_files": curr_selected_intros
        }
        save_gui_config(current_config)

        if engine[0]:
            engine[0].stop()
        root.after(300, root.destroy)

    root.protocol("WM_DELETE_WINDOW", on_closing)

    title_frame = tk.Frame(root, bg="#2c5aa0")
    title_frame.pack(fill=tk.X)
    tk.Label(title_frame, text="📝 Modrinth 批量注册工具", font=("微软雅黑", 16, "bold"),
             fg="white", bg="#2c5aa0", pady=12).pack()

    main = tk.Frame(root, padx=15, pady=10)
    main.pack(fill=tk.BOTH, expand=True)

    cfg = tk.LabelFrame(main, text="配置选项", font=("微软雅黑", 10, "bold"))
    cfg.pack(fill=tk.X, pady=5)

    thread_frame = tk.Frame(cfg)
    thread_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(thread_frame, text="浏览器数:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    thread_spin = tk.Spinbox(thread_frame, from_=1, to=6, width=8, font=("微软雅黑", 10))
    thread_spin.pack(side=tk.LEFT, padx=5)
    tk.Label(thread_frame, text="(同时打开的最大浏览器数量，建议 1~6)", font=("微软雅黑", 9), fg="#666").pack(
        side=tk.LEFT)
    thread_spin.delete(0, tk.END)
    thread_spin.insert(0, str(saved_thread))

    # ===== 发布间隔 =====
    interval_frame = tk.Frame(cfg)
    interval_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(interval_frame, text="发布间隔:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    interval_entry = tk.Entry(interval_frame, textvariable=interval_var, width=8, font=("微软雅黑", 10))
    interval_entry.pack(side=tk.LEFT, padx=5)
    tk.Label(interval_frame, text="(秒，每创建一个收藏夹后的等待时间)", font=("微软雅黑", 9), fg="#666").pack(
        side=tk.LEFT)

    title_dir_frame = tk.Frame(cfg)
    title_dir_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(title_dir_frame, text="标题目录:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    tk.Entry(title_dir_frame, textvariable=title_dir_var, width=50, font=("微软雅黑", 9), state="readonly").pack(
        side=tk.LEFT, padx=5)

    def choose_title_dir():
        d = filedialog.askdirectory(title="选择标题文件所在目录")
        if d:
            title_dir_var.set(d)
            refresh_title_list(d)

    tk.Button(title_dir_frame, text="浏览...", command=choose_title_dir,
              font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)

    intro_dir_frame = tk.Frame(cfg)
    intro_dir_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(intro_dir_frame, text="简介目录:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    tk.Entry(intro_dir_frame, textvariable=intro_dir_var, width=50, font=("微软雅黑", 9), state="readonly").pack(
        side=tk.LEFT, padx=5)

    def choose_intro_dir():
        d = filedialog.askdirectory(title="选择简介文件所在目录")
        if d:
            intro_dir_var.set(d)
            refresh_intro_list(d)

    tk.Button(intro_dir_frame, text="浏览...", command=choose_intro_dir,
              font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)

    output_dir_frame = tk.Frame(cfg)
    output_dir_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(output_dir_frame, text="输出目录:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(
        side=tk.LEFT)
    tk.Entry(output_dir_frame, textvariable=output_dir_var, width=50, font=("微软雅黑", 9), state="readonly").pack(
        side=tk.LEFT, padx=5)

    def choose_output_dir():
        d = filedialog.askdirectory(title="选择结果文件存放目录")
        if d:
            output_dir_var.set(d)

    tk.Button(output_dir_frame, text="浏览...", command=choose_output_dir,
              font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)

    # ===== 邮箱目录 =====
    email_dir_frame = tk.Frame(cfg)
    email_dir_frame.pack(fill=tk.X, pady=5, padx=10)
    tk.Label(email_dir_frame, text="邮箱目录:", font=("微软雅黑", 10, "bold"), width=10, anchor=tk.W).pack(side=tk.LEFT)
    tk.Entry(email_dir_frame, textvariable=email_dir_var, width=50, font=("微软雅黑", 9), state="readonly").pack(
        side=tk.LEFT, padx=5)

    def choose_email_dir():
        d = filedialog.askdirectory(title="选择邮箱文件所在目录")
        if d:
            email_dir_var.set(d)

    tk.Button(email_dir_frame, text="浏览...", command=choose_email_dir,
              font=("微软雅黑", 9), width=8).pack(side=tk.LEFT)

    files_frame = tk.Frame(main)
    files_frame.pack(fill=tk.X, pady=5)

    title_list_frame = tk.LabelFrame(files_frame, text="标题文件列表（勾选添加）", font=("微软雅黑", 10, "bold"),
                                     height=200)
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

    intro_list_frame = tk.LabelFrame(files_frame, text="简介文件列表（勾选添加）", font=("微软雅黑", 10, "bold"),
                                     height=200)
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

            # 恢复上次勾选状态
            if display_name in saved_selected_titles:
                var.set(1)
                if full_path not in title_list:
                    title_list.append(full_path)

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

        # 刷新底部文本框
        title_input.delete("1.0", tk.END)
        title_input.insert(tk.END, "\n".join(title_list))

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

            # 恢复上次勾选状态
            if display_name in saved_selected_intros:
                var.set(1)
                if full_path not in intro_list:
                    intro_list.append(full_path)

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

        # 刷新底部文本框
        intro_input.delete("1.0", tk.END)
        intro_input.insert(tk.END, "\n".join(intro_list))

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

    btn_frame = tk.Frame(main)
    btn_frame.pack(fill=tk.X, pady=10)

    start_btn = tk.Button(btn_frame, text="▶ 启动", bg="#4CAF50", fg="white",
                          font=("微软雅黑", 12, "bold"), width=15, height=1)
    start_btn.pack(side=tk.LEFT, padx=5)

    pause_btn = tk.Button(btn_frame, text="⏸ 暂停", bg="#FF9800", fg="white",
                          font=("微软雅黑", 12, "bold"), width=12, height=1, state=tk.DISABLED)
    pause_btn.pack(side=tk.LEFT, padx=5)

    prog_frame = tk.LabelFrame(main, text="处理进度", font=("微软雅黑", 10, "bold"))
    prog_frame.pack(fill=tk.X, pady=5)

    task_label = tk.Label(prog_frame, text="就绪", font=("微软雅黑", 11, "bold"),
                          fg="#333", anchor=tk.W)
    task_label.pack(fill=tk.X, padx=10, pady=5)

    progress_frame = tk.Frame(prog_frame)
    progress_frame.pack(fill=tk.X, padx=10, pady=2)
    tk.Label(progress_frame, text="收藏夹进度:", font=("微软雅黑", 9), width=12, anchor=tk.W).pack(side=tk.LEFT)
    progress_var = tk.DoubleVar(value=0)
    progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, length=750)
    progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
    progress_label = tk.Label(progress_frame, text="0", font=("微软雅黑", 9), width=8)
    progress_label.pack(side=tk.LEFT, padx=5)

    stats_label = tk.Label(prog_frame, text="已完成: 0 个收藏夹 | 状态: 就绪",
                           font=("微软雅黑", 9), fg="#666", anchor=tk.W)
    stats_label.pack(fill=tk.X, padx=10, pady=5)

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
                    progress_var.set(d["current"] % 100)
                    progress_label.config(text=f"{d['current']}")
                    task_label.config(text=f"状态: {d['status']}")
                    stats_label.config(text=f"已完成: {d['current']} 个收藏夹 | 状态: {d['status']}")
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
        if not email_dir_var.get():
            messagebox.showerror("错误", "请选择邮箱目录")
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
            tc = int(thread_spin.get())
            if not 1 <= tc <= 6:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "浏览器数量必须是 1~6 的整数")
            return

        try:
            interval = float(interval_var.get())
            if interval < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "发布间隔必须是大于等于0的数字")
            return

        email_list = load_emails_from_dir(email_dir_var.get())
        if not email_list:
            messagebox.showerror("错误", "邮箱目录下未找到有效的邮箱账号")
            return

        engine[0] = ModrinthCollector(
            title_files=title_list,
            intro_files=intro_list,
            output_dir=out_dir,
            thread_count=tc,
            log_callback=log,
            progress_callback=update_progress,
            email_list=email_list,
            interval=interval,
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
            old_count = engine[0].completed_collections if engine[0] else 0

            def do_resume():
                if engine[0] and getattr(engine[0], '_is_running', False):
                    root.after(100, do_resume)
                    return

                out_dir = output_dir_var.get()
                try:
                    tc = int(thread_spin.get())
                    if not 1 <= tc <= 6:
                        raise ValueError
                except ValueError:
                    tc = 3

                resume_email_list = load_emails_from_dir(email_dir_var.get())
                try:
                    resume_interval = float(interval_var.get())
                    if resume_interval < 0:
                        raise ValueError
                except ValueError:
                    resume_interval = 0

                engine[0] = ModrinthCollector(
                    title_files=title_list,
                    intro_files=intro_list,
                    output_dir=out_dir,
                    thread_count=tc,
                    log_callback=log,
                    progress_callback=update_progress,
                    email_list=resume_email_list,
                    interval=resume_interval,
                )
                engine[0].completed_collections = old_count
                engine[0].resume()
                pause_btn.config(text="⏸ 暂停")
                status_bar.config(text="处理中...")
                Thread(target=lambda: engine[0].run(), daemon=True).start()

            status_bar.config(text="等待当前任务结束...")
            do_resume()

    start_btn.config(command=start_processing)
    pause_btn.config(command=pause_processing)

    add_log("Modrinth 批量注册工具已启动")
    add_log("请依次选择：标题目录 -> 简介目录 -> 输出目录")
    add_log("勾选需要的文件后，点击「启动」开始")
    refresh()

    # 程序启动完毕，如果配置里保存了目录，自动刷新列表恢复勾选
    if title_dir_var.get() and os.path.isdir(title_dir_var.get()):
        refresh_title_list(title_dir_var.get())
    if intro_dir_var.get() and os.path.isdir(intro_dir_var.get()):
        refresh_intro_list(intro_dir_var.get())

    root.mainloop()


if __name__ == "__main__":
    run_gui()
