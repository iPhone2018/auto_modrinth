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
   # 10位时间戳（秒级）
   prefix = str(int(time.time()))
   return f"{prefix}@qq.com"


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


def single_user_task(task_id: int, user_titles: list, user_intros: list,
                     output_dir: str, log_callback=None, on_collection_created=None):
    driver = None
    session = None
    token = None
    success_count = 0

    try:
        if log_callback:
            log_callback(f"[用户{task_id}] ===== 任务开始 =====")
            log_callback(f"[用户{task_id}] 需要创建收藏夹: {len(user_titles)} 个")
            log_callback(f"[用户{task_id}] 启动浏览器，准备注册...")

        if log_callback:
            log_callback(f"[用户{task_id}] 初始化 Chrome 浏览器...")
        driver, wait, short_wait = init_browser(task_id)
        register_driver(task_id, driver)
        if log_callback:
            log_callback(f"[用户{task_id}] 浏览器初始化成功")
        long_wait = WebDriverWait(driver, 6000)

        driver.get("https://modrinth.com")
        signup_btn = long_wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@href="/auth/sign-up"]')))
        if not retry_click(driver, signup_btn):
            raise Exception(f"点击注册按钮失败")
        if log_callback:
            log_callback(f"[用户{task_id}] 点击注册按钮")

        email_input = long_wait.until(EC.visibility_of_element_located((By.ID, "email")))
        random_email = random_qq_email()
        email_input.clear()
        email_input.send_keys(random_email)
        if log_callback:
            log_callback(f"[用户{task_id}] 输入随机邮箱: {random_email}")

        pwd_input = long_wait.until(EC.visibility_of_element_located((By.ID, "password")))
        pwd_input.clear()
        pwd_input.send_keys("Admin@coc1")
        if log_callback:
            log_callback(f"[用户{task_id}] 输入密码")

        continue_btn = long_wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(text(), 'Continue with Email')]")
            )
        )
        if not retry_click(driver, continue_btn):
            raise Exception(f"点击 Continue with Email 失败")
        if log_callback:
            log_callback(f"[用户{task_id}] 点击 Continue with Email")

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

        for attempt in range(5):
            cookies = driver.get_cookies()
            for ck in cookies:
                if ck["name"] == "auth-token":
                    token = ck["value"]
                    break
            if token:
                if log_callback:
                    log_callback(f"[用户{task_id}] 获取 Token 成功")
                break
            if attempt < 4:  # 前4次失败才等待
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
        for i, (title, intro) in enumerate(zip(user_titles, user_intros)):
            if log_callback:
                log_callback(f"[用户{task_id}] 创建收藏夹 {i + 1}/{len(user_titles)}: {title[:30]}...")

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
                if on_collection_created:
                    on_collection_created(title)
                if log_callback:
                    log_callback(f"[用户{task_id}] 收藏夹创建成功! ID: {collection_id}")
            else:
                if log_callback:
                    log_callback(f"[用户{task_id}] 创建收藏夹失败: {resp.status_code} - {resp.text}")

        if log_callback:
            log_callback(f"[用户{task_id}] 搜索热门模组...")
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

                for cid in collection_ids:
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
            log_callback(f"[用户{task_id}] 全部完成! 创建了 {success_count} 个收藏夹")
        return f"用户{task_id} 成功 {success_count}/{len(user_titles)}"

    except Exception as e:
        error_msg = f"[用户{task_id}] 错误: {str(e)}"
        if log_callback:
            log_callback(error_msg)
        else:
            print(error_msg)
        return f"用户{task_id} 失败: {str(e)}"
    finally:
        unregister_driver(task_id)
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
        try:
            base_dir = os.path.dirname(os.path.realpath(sys.argv[0])) if getattr(sys, 'frozen',
                                                                                 False) else os.path.dirname(
                os.path.abspath(__file__))
            user_data_dir = os.path.join(base_dir, f"chrome_user_data_task_{task_id}")
            _cleanup_chrome_locks(user_data_dir)
        except:
            pass


class ModrinthCollector:
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
        self.completed_collections = 0
        self._is_running = False

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

                    self._log(f"   [提交] 用户 #{next_user_idx} - {len(titles)} 个收藏夹")
                    future = executor.submit(
                        single_user_task,
                        task_id=next_user_idx,
                        user_titles=titles,
                        user_intros=intros,
                        output_dir=self.output_dir,
                        log_callback=self.log_callback,
                        on_collection_created=on_collection_created
                    )
                    futures[future] = next_user_idx
                    time.sleep(2)

                self._log("   收到停止信号，终止提交新任务")
                for f in list(futures.keys()):
                    f.cancel()
                close_all_drivers()
                self._log("   已关闭所有浏览器")

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
        self._log("🛑 已停止")

    def pause(self):
        self.stop_event.set()
        self.pause_event.set()
        close_all_drivers()
        self._log("⏸ 已暂停，关闭所有浏览器...")

    def resume(self):
        self.stop_event.clear()
        self.pause_event.clear()
        self._log("▶ 继续运行")


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
    title_list = []
    intro_list = []
    title_check_vars = {}
    intro_check_vars = {}

    def log(msg):
        log_queue.put(msg)

    def update_progress(data):
        log_queue.put(("progress", data))

    def on_closing():
        if engine[0]:
            engine[0].stop()
        root.destroy()

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
    thread_spin.insert(0, "3")

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

        engine[0] = ModrinthCollector(
            title_files=title_list,
            intro_files=intro_list,
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

                engine[0] = ModrinthCollector(
                    title_files=title_list,
                    intro_files=intro_list,
                    output_dir=out_dir,
                    thread_count=tc,
                    log_callback=log,
                    progress_callback=update_progress
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
    root.mainloop()


if __name__ == "__main__":
    run_gui()
