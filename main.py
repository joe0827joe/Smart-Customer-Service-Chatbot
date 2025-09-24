'''version:1.2.0'''
import sys
import os
import logging
import functools
import time
import datetime
import csv
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from apscheduler.schedulers.blocking import BlockingScheduler
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, InvalidSessionIdException
import sqlite3
import json
import subprocess
import traceback
import undetected_chromedriver as uc

from PyQt5.QtCore import QSettings, QStandardPaths, QObject, pyqtSignal, QTime, pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5 import uic
from threading import Thread


ACCOUNT_NUMBER = None
ACCOUNT_PASSWORD = None
LOGIN_STATE = 0
UI_STATUS_DATA = None

def log_decorator(func):
    """
    日誌裝飾器，用於記錄函數的執行狀態。

    :param func: 被裝飾的函數
    :return: 包裝後的函數
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        logging.info(f'正在執行: {func.__name__}')
        try:
            result = func(*args, **kwargs)
            logging.info(f'成功執行: {func.__name__}')
            return result
        except Exception as e:
            # 詳細紀錄 traceback 到日誌與控制台
            logging.exception(f'在 {func.__name__} 中發生錯誤')
            try:
                print(traceback.format_exc())
            except:
                pass
            raise
    return wrapper


@log_decorator
def verify_cpu_serial(expected_serial):
    """
    驗證當前CPU序列號是否與預期的序列號匹配。

    :param expected_serial: 預期的CPU序列號
    :return: 如果匹配則返回True，否則返回False
    """
    
    try:
        # 方法1: 嘗試使用 wmic (舊版 Windows)
        try:
            command = "wmic cpu get ProcessorId"
            result = subprocess.check_output(command, shell=True, timeout=10).decode()
            cpu_serial_number = result.split('\n')[1].strip()
            if cpu_serial_number:
                return cpu_serial_number == expected_serial
        except (subprocess.CalledProcessError, FileNotFoundError, IndexError):
            pass
        
        # 方法2: 使用 PowerShell (新版 Windows)
        try:
            command = 'powershell "Get-WmiObject -Class Win32_Processor | Select-Object -ExpandProperty ProcessorId"'
            result = subprocess.check_output(command, shell=True, timeout=10).decode().strip()
            if result:
                return result == expected_serial
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass
        
        # 方法3: 使用 Python 的 platform 模組
        try:
            import platform
            system_info = f"{platform.processor()}_{platform.machine()}_{platform.node()}"
            return system_info == expected_serial
        except:
            pass
        
        # 如果所有方法都失敗，返回 False
        print("警告：無法獲取 CPU 序列號，跳過硬體驗證")
        return False
        
    except Exception as e:
        print(f"CPU 序列號驗證時發生錯誤: {e}")
        return False


@log_decorator
def read_whitelist():
    """
    讀取回覆名單，從CSV文件中加載客戶名稱及其最後回覆時間。

    :return: 包含客戶名稱及其最後回覆時間的字典
    """
    whitelist = {}
    try:
        global FILE_PATH
        with open(os.path.join(FILE_PATH, 'reply_whitelist.csv'), mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            whitelist = {rows[0]: rows[1] for rows in reader if rows}
    except FileNotFoundError:
        pass
    return whitelist

@log_decorator
def update_whitelist(customer_name):
    """
    更新回覆名單，將客戶名稱及當前時間寫入CSV文件。

    :param customer_name: 客戶名稱
    """
    whitelist = read_whitelist()
    now = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')
    whitelist[customer_name] = now
    global FILE_PATH
    with open(os.path.join(FILE_PATH, 'reply_whitelist.csv'), mode='w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        for name, last_reply in whitelist.items():
            writer.writerow([name, last_reply])

def read_database(db_path, column):
    """
    從 SQLite 資料庫中讀取指定欄位的單個值。

    :param db_path: 資料庫路徑
    :param column: 要讀取的欄位名稱
    :return: 指定欄位的值
    """
    # 連接到資料庫
    conn = sqlite3.connect(db_path, timeout=10)
    conn.execute('PRAGMA journal_mode=WAL;')
    cursor = conn.cursor()

    # 構建查詢語句
    query = f"SELECT {column} FROM data WHERE id = 1"
    
    # 執行查詢
    cursor.execute(query)
    row = cursor.fetchone()
    
    # 關閉連接
    cursor.close()
    conn.close()
    
    if row:
        value = row[0]
        try:
            # 嘗試將 JSON 字符串轉換回字典
            return json.loads(value)
        except (TypeError, json.JSONDecodeError):
            return value
    else:
        return None

#瀏覽器對象
class WebDriverManager():
    def __init__(self):
        try:
            # 可使用既有的真實 Chrome 個人資料夾 (環境變數 REAL_CHROME_PROFILE)
            profile_dir = os.environ.get('REAL_CHROME_PROFILE') or os.path.join(FILE_PATH, 'chrome_profile')
            os.makedirs(profile_dir, exist_ok=True)
            self.profile_dir = profile_dir

            # 修復可能壞掉的 Chrome 偏好檔，避免 uc 載入 JSON 失敗
            try:
                self._sanitize_chrome_profile()
            except Exception:
                pass
            # 保存 UA 並建立 options
            self.user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
            self.options = self._build_options()
            self.driver = uc.Chrome(options=self.options)

            self.driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": """
                  Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                  Object.defineProperty(navigator, 'languages', {get: () => ['zh-TW','zh','en-US']});
                  Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
                  Object.defineProperty(navigator, 'platform', {get: () => 'Win32'});
                  Object.defineProperty(navigator, 'vendor', {get: () => 'Google Inc.'});
                  Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
                  Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});
                """}
            )
            try:
                # 時區與在地化
                self.driver.execute_cdp_cmd('Emulation.setTimezoneOverride', { 'timezoneId': 'Asia/Taipei' })
            except Exception:
                pass
            try:
                self.driver.execute_cdp_cmd('Emulation.setLocaleOverride', { 'locale': 'zh-TW' })
            except Exception:
                pass
            try:
                self.driver.execute_cdp_cmd('Network.setUserAgentOverride', {
                    'userAgent': self.user_agent,
                    'acceptLanguage': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
                    'platform': 'Windows'
                })
            except Exception:
                pass

            # 導覽紀錄
            self.last_url = None
            self.nav_log_path = os.path.join(FILE_PATH, 'nav_history.log')
            try:
                with open(self.nav_log_path, 'a', encoding='utf-8') as f:
                    f.write(f"===== 啟動瀏覽器 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====\n")
            except Exception:
                pass
        except Exception:
            logging.exception('初始化 WebDriver 失敗')
            print(traceback.format_exc())
            raise

    def _sanitize_chrome_profile(self):
        try:
            default_dir = os.path.join(self.profile_dir, 'Default')
            os.makedirs(default_dir, exist_ok=True)
            candidates = [
                os.path.join(default_dir, 'Preferences'),
                os.path.join(default_dir, 'Secure Preferences'),
                os.path.join(self.profile_dir, 'Local State'),
            ]
            for path in candidates:
                if not os.path.exists(path):
                    continue
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        json.load(f)
                except Exception:
                    # JSON 損壞：備份並以空物件取代
                    try:
                        bak = path + '.bak'
                        try:
                            os.replace(path, bak)
                        except Exception:
                            pass
                        with open(path, 'w', encoding='utf-8') as f:
                            f.write('{}')
                    except Exception:
                        pass
        except Exception:
            pass

    def _build_options(self):
        options = uc.ChromeOptions()
        prefs = {"profile.default_content_setting_values.notifications": 2}
        options.add_experimental_option("prefs", prefs)
        options.add_argument(f'--user-agent={self.user_agent}')
        options.add_argument(f'--user-data-dir={self.profile_dir}')
        options.add_argument('--profile-directory=Default')
        options.add_argument('--lang=zh-TW')
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--window-size=1280,860')
        return options
    
    def restart_driver(self, target_after: str | None = None):
        try:
            try:
                self.driver.quit()
            except Exception:
                pass
            # 重啟時不可重用 options，需重建
            self.options = self._build_options()
            self.driver = uc.Chrome(options=self.options)
            self.driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": """
                  Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                  Object.defineProperty(navigator, 'languages', {get: () => ['zh-TW','zh','en-US']});
                  Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
                """}
            )
            if target_after:
                self.driver.get(target_after)
        except Exception:
            logging.exception('重啟 WebDriver 失敗')
            print(traceback.format_exc())
            raise

    def ensure_alive(self, target_after: str | None = None):
        try:
            _ = self.driver.current_url
            self.driver.execute_script("return 1")
        except (InvalidSessionIdException, WebDriverException):
            self.restart_driver(target_after)
        
    @log_decorator
    def wait_for_element(self, element, by, value, timeout=10):
        """等待元素出現並返回該元素"""
        try:
            return WebDriverWait(element, timeout).until(lambda d: d.find_element(by, value))
        except (InvalidSessionIdException, WebDriverException):
            self.ensure_alive()
            return WebDriverWait(self.driver, timeout).until(lambda d: d.find_element(by, value))
        except TimeoutException:
            try:
                self.record_nav(f"wait_for_element timeout by={by} value={value}")
            except Exception:
                pass
            raise

    @log_decorator
    def wait_for_elements(self, element, by, value, timeout=10):
        """等待元素出現並返回所有元素"""
        try:
            return WebDriverWait(element, timeout).until(lambda d: d.find_elements(by, value))
        except (InvalidSessionIdException, WebDriverException):
            self.ensure_alive()
            return WebDriverWait(self.driver, timeout).until(lambda d: d.find_elements(by, value))
        except TimeoutException:
            try:
                self.record_nav(f"wait_for_elements timeout by={by} value={value}")
            except Exception:
                pass
            raise

    def record_nav(self, note=""):
        try:
            url = self.driver.current_url
            title = ""
            try:
                title = self.driver.title
            except Exception:
                pass
            referrer = ""
            try:
                referrer = self.driver.execute_script("return document.referrer || ''")
            except Exception:
                pass
            ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            if url != self.last_url or note:
                line = f"[{ts}] url={url} title={title} ref={referrer} note={note}\n"
                try:
                    with open(self.nav_log_path, 'a', encoding='utf-8') as f:
                        f.write(line)
                except Exception:
                    pass
                self.last_url = url
        except Exception:
            pass

    def dump_cookies(self, path: str):
        try:
            cookies = self.driver.get_cookies()
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(cookies, f, ensure_ascii=False)
            print(f"Cookies 已匯出: {path}")
        except Exception:
            print("Cookies 匯出失敗")

    def load_cookies(self, path: str, domain_hint: str = ".shopee.tw"):
        try:
            if not os.path.exists(path):
                print("找不到 cookies 檔，略過匯入")
                return
            with open(path, 'r', encoding='utf-8') as f:
                cookies = json.load(f)
            # Selenium 需先到對應網域一頁才能設 cookie
            self.driver.get("https://seller.shopee.tw")
            time.sleep(1)
            for ck in cookies:
                # 清理無效欄位
                ck.pop('sameSite', None)
                # 如果 domain 不含網域，補上提示
                if 'domain' not in ck:
                    ck['domain'] = domain_hint
                try:
                    self.driver.add_cookie(ck)
                except Exception:
                    continue
            self.driver.get("https://seller.shopee.tw/new-webchat/conversations")
            print("Cookies 已匯入，嘗試直達聊天頁")
        except Exception:
            print("Cookies 匯入失敗")

    def click_all_conversations(self, timeout: int = 10) -> bool:
        """點擊『全部聊聊』過濾器。回傳是否成功。"""
        try:
            try:
                self.record_nav("click_all: enter")
            except Exception:
                pass
            # 前置等待：確認頁面已渲染出可用的 UI
            try:
                end_ts = time.time() + max(6, timeout)
                while time.time() < end_ts:
                    try:
                        ready = self.driver.execute_script(
                            """
                            const hasDataCy = !!document.querySelector('[data-cy^=\'webchat-conversation-filter\']');
                            const textAll = (document.body.innerText||'').includes('全部聊聊') || (document.body.innerText||'').trim()==='全部' || (document.body.innerText||'').includes('All');
                            const rvList = !!document.querySelector('.ReactVirtualized__List');
                            return hasDataCy || textAll || rvList;
                            """
                        )
                        if ready:
                            break
                    except Exception:
                        pass
                    time.sleep(0.4)
                try:
                    self.record_nav("click_all: prewait done")
                except Exception:
                    pass
            except Exception:
                pass

            # 先確保過濾器切換器存在，若有先點開
            try:
                try:
                    self.record_nav("click_all: try root filter container")
                except Exception:
                    pass
                root = self.wait_for_element(self.driver, By.CSS_SELECTOR, "[data-cy='webchat-conversation-filter-root-new']", timeout=3)
                try:
                    root.click()
                except Exception:
                    try:
                        # 按下其中的切換按鈕
                        toggle = self.wait_for_element(root, By.CSS_SELECTOR, "[data-cy^='webchat-conversation-filter-filter-']", timeout=2)
                        try:
                            toggle.click()
                        except Exception:
                            self.driver.execute_script("arguments[0].click();", toggle)
                    except Exception:
                        pass
            except Exception:
                pass

            # 直接找 All 過濾器按鈕
            try:
                try:
                    self.record_nav("click_all: try data-cy direct")
                except Exception:
                    pass
                all_btn = self.wait_for_element(self.driver, By.CSS_SELECTOR, "[data-cy='webchat-conversation-filter-filter-all']", timeout=3)
                try:
                    all_btn.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", all_btn)
                return True
            except Exception:
                pass

            # 若按鈕不存在，先打開下拉，再選文字/或 data-cy
            try:
                try:
                    self.record_nav("click_all: try open toggle")
                except Exception:
                    pass
                toggle = self.wait_for_element(self.driver, By.CSS_SELECTOR, "[data-cy^='webchat-conversation-filter-filter-']", timeout=timeout)
                try:
                    toggle.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", toggle)
                # 備援：嘗試用鍵盤開啟下拉
                try:
                    toggle.send_keys(Keys.ALT, Keys.ARROW_DOWN)
                except Exception:
                    pass
            except Exception:
                pass

            # 先試 data-cy 再試文字
            try:
                try:
                    self.record_nav("click_all: try data-cy after toggle")
                except Exception:
                    pass
                item = self.wait_for_element(self.driver, By.CSS_SELECTOR, "[data-cy='webchat-conversation-filter-filter-all']", timeout=6)
                try:
                    item.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", item)
                return True
            except Exception:
                pass

            # 使用你提供的 sidebar 絕對 XPath（優先度高於純文字備援）
            try:
                try:
                    self.record_nav("click_all: try sidebar absolute xpath")
                except Exception:
                    pass
                sidebar_all = self.wait_for_element(
                    self.driver,
                    By.XPATH,
                    "//div[@id='sidebar']/div[3]/div/div/div/div/div[2]/div",
                    timeout=6
                )
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sidebar_all)
                    sidebar_all.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", sidebar_all)
                return True
            except Exception:
                pass

            # 先確保 sidebar 容器存在，再於其中尋找包含文字的項目
            try:
                try:
                    self.record_nav("click_all: try find text in #sidebar")
                except Exception:
                    pass
                sidebar = self.wait_for_element(self.driver, By.ID, "sidebar", timeout=6)
                # 在 sidebar 內用文字匹配找「全部聊聊 / 全部 / All」
                candidates = sidebar.find_elements(
                    By.XPATH,
                    ".//div[contains(normalize-space(.), '全部聊聊') or contains(normalize-space(.), '全部') or contains(normalize-space(.), 'All')]"
                )
                try:
                    self.record_nav(f"click_all: #sidebar text candidates={len(candidates)}")
                except Exception:
                    pass
                for el in candidates:
                    try:
                        if not el.is_displayed():
                            continue
                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        try:
                            el.click()
                        except Exception:
                            # 嘗試點擊可點擊的父層
                            clickable_parent = self._find_clickable_parent(el, max_depth=5)
                            target = clickable_parent or el
                            try:
                                target.click()
                            except Exception:
                                self.driver.execute_script("arguments[0].click();", target)
                        return True
                    except Exception:
                        continue
            except Exception:
                pass

            # 若主頁未找到，嘗試在 iframe 內尋找並點擊
            try:
                try:
                    frs = self.driver.find_elements(By.TAG_NAME, 'iframe')
                    self.record_nav(f"click_all: try iframes; count={len(frs)}")
                except Exception:
                    frs = self.driver.find_elements(By.TAG_NAME, 'iframe')
                iframes = frs
                iframes = self.driver.find_elements(By.TAG_NAME, 'iframe')
                for fr in iframes:
                    try:
                        try:
                            self.record_nav("click_all: switch into iframe")
                        except Exception:
                            pass
                        self.driver.switch_to.frame(fr)
                        # 優先 data-cy
                        try:
                            try:
                                self.record_nav("click_all: iframe try data-cy")
                            except Exception:
                                pass
                            el = WebDriverWait(self.driver, 3).until(lambda d: d.find_element(By.CSS_SELECTOR, "[data-cy='webchat-conversation-filter-filter-all']"))
                            try:
                                el.click()
                            except Exception:
                                self.driver.execute_script("arguments[0].click();", el)
                            return True
                        except Exception:
                            pass
                        # 使用你的 sidebar 結構 XPath（若該 frame 內有 sidebar）
                        try:
                            try:
                                self.record_nav("click_all: iframe try sidebar xpath")
                            except Exception:
                                pass
                            el2 = WebDriverWait(self.driver, 3).until(lambda d: d.find_element(By.XPATH, "//div[@id='sidebar']/div[3]/div/div/div/div/div[2]/div"))
                            try:
                                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el2)
                                el2.click()
                            except Exception:
                                self.driver.execute_script("arguments[0].click();", el2)
                            return True
                        except Exception:
                            pass
                        # 文字匹配備援
                        try:
                            try:
                                self.record_nav("click_all: iframe try text match")
                            except Exception:
                                pass
                            nodes = self.driver.find_elements(By.XPATH, "//div[contains(normalize-space(.), '全部聊聊') or contains(normalize-space(.), '全部') or contains(normalize-space(.), 'All')]")
                            for node in nodes:
                                try:
                                    if not node.is_displayed():
                                        continue
                                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", node)
                                    try:
                                        node.click()
                                    except Exception:
                                        # 點父層
                                        try:
                                            parent = node.find_element(By.XPATH, "./ancestor::div[contains(@class,'qK2REYutol') or contains(@class,'tab') or contains(@class,'menu')][1]")
                                        except Exception:
                                            parent = None
                                        target = parent or node
                                        try:
                                            target.click()
                                        except Exception:
                                            self.driver.execute_script("arguments[0].click();", target)
                                    return True
                                except Exception:
                                    continue
                        except Exception:
                            pass
                    except Exception:
                        continue
                    finally:
                        try:
                            self.driver.switch_to.default_content()
                        except Exception:
                            pass
            except Exception:
                pass

            # 最終備援：跨 Shadow DOM 搜尋 data-cy 或文字
            try:
                try:
                    self.record_nav("click_all: shadow DOM crawl")
                except Exception:
                    pass
                el = self.driver.execute_script(
                    """
                    function crawl(root){
                      try {
                        const btn = root.querySelector("[data-cy='webchat-conversation-filter-filter-all']");
                        if (btn) return btn;
                        const divs = Array.from(root.querySelectorAll('div'));
                        for (const d of divs){
                          const t=(d.innerText||'').trim();
                          if (t.includes('全部聊聊') || t === '全部' || t.includes('All')) return d;
                        }
                        const all = root.querySelectorAll('*');
                        for (const e of all){
                          if (e.shadowRoot){
                            const r = crawl(e.shadowRoot);
                            if (r) return r;
                          }
                        }
                      } catch(e){}
                      return null;
                    }
                    return crawl(document);
                    """
                )
                if el:
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    except Exception:
                        pass
                    try:
                        self.driver.execute_script("arguments[0].click();", el)
                    except Exception:
                        pass
                    return True
            except Exception:
                pass
            try:
                try:
                    self.record_nav("click_all: try dropdown item xpath fallback")
                except Exception:
                    pass
                item = self.wait_for_element(
                    self.driver,
                    By.XPATH,
                    "//div[contains(@class,'shopee-react-dropdown-item') or @data-cy][contains(., '全部') or contains(., 'All')]",
                    timeout=3
                )
                try:
                    item.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", item)
                return True
            except Exception:
                pass

            # 精準定位：直接找 LxWGUqZsIZ 文字為「全部聊聊」的父層 qK2REYutol
            try:
                try:
                    self.record_nav("click_all: try LxWGUqZsIZ parent")
                except Exception:
                    pass
                # 直接找 div.LxWGUqZsIZ 文字為「全部聊聊」的最近父層 div.qK2REYutol
                parent = self.wait_for_element(
                    self.driver,
                    By.XPATH,
                    "//div[contains(@class,'LxWGUqZsIZ')][normalize-space()='全部聊聊']/ancestor::div[contains(@class,'qK2REYutol')][1]",
                    timeout=6
                )
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", parent)
                    parent.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", parent)
                return True
            except Exception:
                pass

            # 有些版面『全部聊聊』是頁籤，非下拉（修正括號優先順序）
            try:
                try:
                    self.record_nav("click_all: try tab-like structure xpath")
                except Exception:
                    pass
                # 使用你提供的結構：div.qK2REYutol.ArmcNxhBwZ > div.LxWGUqZsIZ (文字：全部聊聊)
                tab = self.wait_for_element(
                    self.driver,
                    By.XPATH,
                    "//div[contains(@class,'xPjxpCOEQp') and contains(@class,'ArmcNxhBwZ')]//div[(contains(@class,'qK2REYutol') and contains(@class,'ArmcNxhBwZ')) or contains(@class,'zFVr2WzGcI')]//div[contains(@class,'LxWGUqZsIZ')][contains(., '全部聊聊') or contains(., '全部') or contains(., 'All')]",
                    timeout=3
                )
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", tab)
                    tab.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", tab)
                return True
            except Exception:
                pass

            # 再試任何 data-cy 可能的鍵名
            candidates = [
                "webchat-conversation-left-tab-all",
                "webchat-conversation-tab-all",
                "webchat-conversation-filter-all",
            ]
            for key in candidates:
                try:
                    try:
                        self.record_nav(f"click_all: try candidate data-cy={key}")
                    except Exception:
                        pass
                    el = self.wait_for_element(self.driver, By.CSS_SELECTOR, f"[data-cy='{key}']", timeout=2)
                    try:
                        el.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", el)
                    return True
                except Exception:
                    continue

            # 文字匹配最終備援：normalize-space + 可視判斷 + 向上找可點擊父層 + JS 強制點
            try:
                try:
                    self.record_nav("click_all: final text-match fallback")
                except Exception:
                    pass
                labels = ["全部聊聊", "全部", "All"]
                for label in labels:
                    elems = self.driver.find_elements(By.XPATH, f"//div[contains(normalize-space(.), '{label}')]")
                    for el in elems:
                        try:
                            if not el.is_displayed():
                                continue
                        except Exception:
                            pass
                        target = el
                        try:
                            parent = el.find_element(By.XPATH, "./ancestor::div[contains(@class,'qK2REYutol')][1]")
                            if parent:
                                target = parent
                        except Exception:
                            pass
                        try:
                            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
                            try:
                                target.click()
                            except Exception:
                                self.driver.execute_script("arguments[0].click();", target)
                            return True
                        except Exception:
                            continue
            except Exception:
                pass
        except Exception:
            pass
        return False

#蝦皮登入
@log_decorator
def autologin(drivermanager, login_status, account_number, account_password):
    global LOGIN_STATE
    try:
        targets = [
            "https://seller.shopee.tw/new-webchat/conversations",
        ]
        target = targets[0]
        login_url = "https://accounts.shopee.tw/seller/login?next=" + target
        drivermanager.driver.get(login_url)
        drivermanager.driver.maximize_window()
        drivermanager.record_nav("open login")

        def human_sleep(a, b):
            time.sleep(random.uniform(a, b))

        def wait_first_available(selectors, timeout=30):
            last_err = None
            for by, value in selectors:
                try:
                    return WebDriverWait(drivermanager.driver, timeout).until(
                        EC.visibility_of_element_located((by, value))
                    )
                except Exception as e:
                    last_err = e
            if last_err:
                raise last_err

        def is_chat_ready(timeout=8):
            try:
                drivermanager.ensure_alive()
            except Exception:
                pass
            try:
                # 新聊聊頁的過濾器或清單任一就緒即視為成功
                drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.CSS_SELECTOR,
                    "[data-cy^='webchat-conversation-filter-']",
                    timeout=timeout
                )
                return True
            except Exception:
                pass
            # 新聊聊頁 filter 容器
            try:
                drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.CSS_SELECTOR,
                    "[data-cy='webchat-conversation-filter-root-new']",
                    timeout=timeout
                )
                return True
            except Exception:
                pass
            # conversation cell 也代表頁面載入
            try:
                drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.CSS_SELECTOR,
                    "[data-cy='webchat-conversation-cell-root']",
                    timeout=timeout
                )
                return True
            except Exception:
                pass
            # 再試訊息容器或列表
            try:
                drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.ID,
                    "messagesContainer",
                    timeout=timeout
                )
                return True
            except Exception:
                pass
            try:
                drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.CSS_SELECTOR,
                    ".ReactVirtualized__List",
                    timeout=timeout
                )
                return True
            except Exception:
                pass
            # 檢查是否有「全部聊聊」的結構（LxWGUqZsIZ + qK2REYutol）
            try:
                drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.XPATH,
                    "//div[contains(@class,'LxWGUqZsIZ')][normalize-space()='全部聊聊']/ancestor::div[contains(@class,'qK2REYutol')][1]",
                    timeout=timeout
                )
                return True
            except Exception:
                pass
            # 檢查是否為伺服器錯誤頁
            try:
                _ = drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.XPATH,
                    "//*[contains(text(),'伺服器錯誤') or contains(text(),'服务器错误') or contains(text(),'Server Error')]",
                    timeout=2
                )
            except Exception:
                pass
            return False

        verify_entry = "https://shopee.tw/verify/ivs?is_initial=true"

        def switch_to_verify_tab():
            try:
                handles = drivermanager.driver.window_handles
                for h in reversed(handles):
                    try:
                        drivermanager.driver.switch_to.window(h)
                        time.sleep(0.2)
                        u = drivermanager.driver.current_url
                        if ("/verify" in u) or ("/challenge" in u) or ("captcha" in u) or ("/portal/sgw" in u):
                            drivermanager.record_nav("switched to verify tab")
                            return True
                    except Exception:
                        continue
            except Exception:
                pass
            return False

        # 啟動即嘗試以 Cookie 直接登入（若已有 cookies.json）
        try:
            cookies_path = os.path.join(FILE_PATH, 'cookies.json')
            if os.path.exists(cookies_path):
                drivermanager.record_nav("try load cookies")
                drivermanager.load_cookies(cookies_path)
                if is_chat_ready(timeout=8):
                    drivermanager.record_nav("cookie login success")
                    LOGIN_STATE = 1
                    login_status("登入成功（Cookie）")
                    return
        except Exception:
            pass

        selectors_account = [
            (By.NAME, "loginKey"),
            (By.CSS_SELECTOR, "input[name='loginKey']"),
            (By.XPATH, "//input[@name='loginKey' or @type='text' or contains(@placeholder,'Email') or contains(@placeholder,'電話') or contains(@placeholder,'帳號')]")
        ]
        selectors_password = [
            (By.NAME, "password"),
            (By.CSS_SELECTOR, "input[name='password']"),
            (By.XPATH, "//input[@name='password' or @type='password']")
        ]
        login_btn_selectors = [
            (By.CSS_SELECTOR, "form button[type='submit']"),
            (By.XPATH, "//button[contains(., '登入') or contains(., 'Login')]")
        ]

        start = time.time()
        max_wait = 600  # 放寬，含人工驗證等待

        # 狀態變數
        has_submitted = False
        in_verify = False
        verify_lock = False  # 一旦偵測到驗證頁，鎖定人工等待，不再自動送單
        verify_start_ts = None
        verify_pin_end_ts = None  # 驗證鎖定後的前置釘住時間
        verify_forceback_count = 0
        login_loop_count = 0
        last_state_change_ts = time.time()

        def wait_manual_until_seller(max_seconds=600):
            login_status("請在瀏覽器完成驗證/二步驟，完成後停留在賣家聊聊頁面…")
            drivermanager.record_nav("manual mode: waiting user to complete verify")
            end_ts = time.time() + max_seconds
            while time.time() < end_ts:
                url_now = drivermanager.driver.current_url
                if "seller.shopee.tw" in url_now:
                    drivermanager.record_nav("manual mode success")
                    return True
                human_sleep(0.8, 1.5)
            drivermanager.record_nav("manual mode timeout")
            return False

        while True:
            url = drivermanager.driver.current_url
            drivermanager.record_nav()

            # 驗證鎖定期間，若被切到新分頁或非驗證頁，嘗試切回或導向驗證入口
            if verify_lock:
                if url.startswith("chrome://") or url.startswith("about:"):
                    if not switch_to_verify_tab():
                        drivermanager.driver.get(verify_entry)
                        human_sleep(0.6, 1.2)
                        continue
                # 若不在 verify/challenge/captcha 也不在聊天頁，就嘗試切回驗證分頁
                if ("/verify" not in url) and ("/challenge" not in url) and ("captcha" not in url) and ("/portal/sgw" not in url) and ("seller.shopee.tw" not in url):
                    if not switch_to_verify_tab():
                        drivermanager.driver.get(verify_entry)
                        human_sleep(0.6, 1.2)
                        continue
                # 驗證鎖定後的前 10 秒，強制釘住在驗證頁，避免被帶走
                if verify_pin_end_ts and time.time() < verify_pin_end_ts:
                    if ("/verify" not in url) and ("/challenge" not in url) and ("captcha" not in url) and ("/portal/sgw" not in url):
                        verify_forceback_count += 1
                        drivermanager.record_nav(f"pin-back-to-verify #{verify_forceback_count}")
                        drivermanager.driver.get(verify_entry)
                        human_sleep(0.8, 1.0)
                        continue

            if "accounts.shopee.tw" in url:
                # 若驗證鎖定中卻回到登入，強制導回驗證入口以便人工完成
                if verify_lock:
                    try:
                        drivermanager.record_nav("verify locked but back to accounts -> force go verify entry")
                        drivermanager.driver.get(verify_entry)
                        time.sleep(2)
                    except Exception:
                        pass
                # 若剛從 verify 回到登入，視為循環一次
                if in_verify:
                    login_loop_count += 1
                    drivermanager.record_nav(f"loop back to accounts (count={login_loop_count})")
                    in_verify = False
                    last_state_change_ts = time.time()

                # 僅送出一次表單，避免重複觸發風控
                if not has_submitted and not verify_lock:
                    try:
                        drivermanager.record_nav("on accounts: fill and submit")
                        acc = wait_first_available(selectors_account, timeout=20)
                        acc.clear(); human_sleep(0.6, 1.2); acc.send_keys(account_number)
                        pwd = wait_first_available(selectors_password, timeout=20)
                        pwd.clear(); human_sleep(0.6, 1.2); pwd.send_keys(account_password)
                        human_sleep(0.6, 1.2)
                        try:
                            btn = wait_first_available(login_btn_selectors, timeout=10)
                            btn.click()
                        except Exception:
                            try: pwd.submit()
                            except Exception: pass
                        has_submitted = True
                        human_sleep(2.5, 5.5)
                    except Exception:
                        pass

                # 循環過多次則切手動模式
                if login_loop_count >= 2 and (time.time() - last_state_change_ts) < 120:
                    if wait_manual_until_seller(600):
                        url = drivermanager.driver.current_url
                        if not url.startswith(target):
                            drivermanager.driver.get(target)
                        LOGIN_STATE = 1
                        login_status("登入成功")
                        return
                    else:
                        LOGIN_STATE = 0
                        login_status("登入失敗：驗證未通過或逾時")
                        return

            if "seller.shopee.tw" in url:
                # 只有在聊天頁元素就緒時才視為成功
                if not any(url.startswith(t) for t in targets):
                    drivermanager.driver.get(target)
                if is_chat_ready(timeout=6):
                    drivermanager.record_nav("arrived seller chats - ready")
                    # 先切到『全部聊聊』
                    try:
                        if drivermanager.click_all_conversations(timeout=6):
                            drivermanager.record_nav("clicked all conversations")
                    except Exception:
                        pass
                    try:
                        drivermanager.dump_cookies(os.path.join(FILE_PATH, 'cookies.json'))
                    except Exception:
                        pass
                    LOGIN_STATE = 1
                    login_status("登入成功")
                    return
                else:
                    drivermanager.record_nav("arrived seller chats - not ready")
                    if verify_lock:
                        # 回到驗證入口，等待人工完成
                        drivermanager.driver.get(verify_entry)
                        human_sleep(1.0, 2.0)
                        continue
                    # 若未鎖定驗證，先稍等再重試
                    human_sleep(1.0, 2.0)
                    continue

            # 擴充驗證偵測：含 IVS、Portal、Challenge、Captcha 等中間頁
            if ("/verify" in url) or ("shopee.tw/verify" in url) or ("/portal/sgw" in url) or ("/challenge" in url) or ("captcha" in url):
                LOGIN_STATE = 0
                login_status("請在瀏覽器完成驗證，完成後將自動繼續…")
                if not verify_lock:
                    verify_lock = True
                    verify_start_ts = time.time()
                    verify_pin_end_ts = verify_start_ts + 10
                    drivermanager.record_nav("verify detected - manual wait lock")
                in_verify = True
                last_state_change_ts = time.time()
                # 進入人工等待，最長 15 分鐘；成功條件為聊天頁可用
                if wait_manual_until_seller(900):
                    url = drivermanager.driver.current_url
                    if not any(url.startswith(t) for t in targets):
                        drivermanager.driver.get(target)
                    if is_chat_ready(timeout=8):
                        try:
                            if drivermanager.click_all_conversations(timeout=6):
                                drivermanager.record_nav("clicked all conversations")
                        except Exception:
                            pass
                        try:
                            drivermanager.dump_cookies(os.path.join(FILE_PATH, 'cookies.json'))
                        except Exception:
                            pass
                        LOGIN_STATE = 1
                        login_status("登入成功")
                        return
                    else:
                        # 仍未就緒則繼續等待/循環
                        if not switch_to_verify_tab():
                            drivermanager.driver.get(verify_entry)
                        continue
                else:
                    LOGIN_STATE = 0
                    login_status("驗證逾時，請重試")
                    drivermanager.record_nav("verify timeout")
                    return

            if time.time() - start > max_wait:
                LOGIN_STATE = 0
                login_status("驗證超時")
                drivermanager.record_nav("login timeout")
                return

            human_sleep(0.8, 1.5)
    except Exception:
        try:
            current_url = drivermanager.driver.current_url
            print(f"登入失敗，當前網址: {current_url}")
            screenshot_path = os.path.join(FILE_PATH, 'login_error.png')
            drivermanager.driver.save_screenshot(screenshot_path)
            print(f"已儲存截圖: {screenshot_path}")
            try:
                drivermanager.record_nav(f"exception; screenshot={screenshot_path}")
            except Exception:
                pass
        except Exception:
            pass
        logging.exception('自動登入流程失敗')
        print(traceback.format_exc())
        raise

import re
import openai

#ChatGPT機器人
@log_decorator
def ChatGPT_Robot(api_key, assistant_id, previous_conversations):
    try:
        # 設置您的API密鑰
        client = openai.OpenAI(api_key=api_key)

        thread = client.beta.threads.create()

        # 將之前的對話紀錄和新的用戶輸入加入到對話線程中
        for message in previous_conversations:
            client.beta.threads.messages.create(
                thread_id=thread.id,
                role=message["role"],
                content=message["content"]
            )

        run = client.beta.threads.runs.create_and_poll(
            thread_id=thread.id,
            assistant_id=assistant_id,
            instructions="你現在是賣家，請協助客人解決問題"
        )

        if run.status == 'completed':
            messages = client.beta.threads.messages.list(
                thread_id=thread.id
            )

            # messages 是從 API 獲得的 SyncCursorPage[Message] 對象
            for message in messages.data:
                if message.role == 'assistant':
                    # 直接訪問 Message 物件的 content 屬性
                    for content_block in message.content:
                        # 檢查 content_block 是否有 'text' 屬性
                        if hasattr(content_block, 'text'):
                            output = content_block.text.value
                            clean_output = re.sub(r"【.*?†來源】", "", output)
                            clean_output = re.sub(r"【.*?†source】", "", clean_output)
                            return clean_output
        else:
            print("LLM處理狀態:" + run.status)
        return None
    except Exception:
        logging.exception('ChatGPT_Robot 執行失敗')
        print(traceback.format_exc())
        raise

from bs4 import BeautifulSoup

@log_decorator
def chatgpt_extract_conversations(html_content):
    try:
        # 使用BeautifulSoup解析HTML
        soup = BeautifulSoup(html_content, 'html.parser')

        # 創建一個空列表來存儲提取的對話
        conversations = []

        # 提取所有對話消息
        messages = soup.find_all('div', class_='DEwekPN7v2')
        for message in messages:
            # 判斷對話類型
            if message.find(attrs={"data-cy": True}).get('data-cy') == 'webchat-message-receive':
                convo_type = 'user'
            elif message.find(attrs={"data-cy": True}).get('data-cy') == 'webchat-message-send':
                convo_type = 'assistant'
            elif message.find(class_='RtO616EACf'):
                convo_type = 'system'
            else:
                convo_type = 'unknown'

            # 如果是系統或未知類型，則跳過
            if convo_type in ['system', 'unknown']:
                continue

            # 提取時間戳記
            time_tag = message.find(class_='hvckbUfzJ0')
            time_text = time_tag.get_text(strip=True) if time_tag else '時間未知'

            # 提取對話文本
            text_tag = message.find('pre')
            text = text_tag.get_text(strip=True) if text_tag else '內容未知'

            # 移除文本中的時間戳記
            if time_tag:
                text = text.replace(time_tag.get_text(strip=True), '').strip()

            # 將對話添加到列表
            conversations.append({"role": convo_type, "content": text})

        return conversations
    except Exception:
        logging.exception('解析對話 HTML 失敗')
        print(traceback.format_exc())
        raise

@log_decorator
def view_custormer_chat(drivermanager):
    try:
        # 新增的滾動和內容檢查邏輯
        countdown_timer = 10  # 設定倒計時時間（秒）
        found_content = set()  # 儲存已經找到的內容
        start_time = time.time()  # 開始倒計時

        while True:
            current_time = time.time()
            elapsed_time = current_time - start_time

            # 更新倒計時並檢查是否超時
            remaining_time = countdown_timer - elapsed_time
            if remaining_time <= 0:
                print("倒計時結束，結束程式。")
                break

            # 將元素的HTML存儲在一個列表中
            elements = drivermanager.wait_for_elements(drivermanager.driver, By.CLASS_NAME, 'DEwekPN7v2.undefined')
            time.sleep(1)

            html_list = []
            for element in elements:
                html_list.append(element.get_attribute('outerHTML'))

            scroll_element = drivermanager.wait_for_element(drivermanager.driver, By.ID, '\\#message-virtualized-list') #取得對話視窗滾動元素
            # 對指定元素進行向上滾動
            drivermanager.driver.execute_script('arguments[0].scrollTop = arguments[0].scrollTop - 100;', scroll_element)

            # 檢查新內容
            new_elements = drivermanager.wait_for_elements(drivermanager.driver, By.CLASS_NAME, 'DEwekPN7v2.undefined')
            new_html_list = [element.get_attribute('outerHTML') for element in new_elements]

            # 比對並更新HTML列表
            for html in new_html_list:
                if html not in html_list:
                    html_list.insert(0, html)
                    start_time = time.time()  # 開始重新倒計時

        return '\n'.join(html_list)
    except Exception:
        logging.exception('擷取客服對話視窗失敗')
        print(traceback.format_exc())
        raise

@log_decorator
def answer_buyer_check(customer_name):
    whitelist_path = 'reply_whitelist.csv' # 回覆名單的檔案路徑
    # 讀取回覆名單
    whitelist = {}
    try:
        with open(whitelist_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            whitelist = {rows[0]: rows[1] for rows in reader if rows}
    except FileNotFoundError:
        pass

    # 確認是否回覆客戶
    now = datetime.datetime.now()
    if customer_name in whitelist:
        try:
            last_reply_time = datetime.datetime.strptime(whitelist[customer_name], '%Y/%m/%d %H:%M')
        except ValueError:
            # 如果格式不匹配，則記錄錯誤並跳過此客戶
            print(f"日期時間格式錯誤: {whitelist[customer_name]}")
            return False
        # 檢查是否已經超過2小時
        if (now - last_reply_time).total_seconds() > 7200:
            return True
    else:
        return True
    return False
                
@log_decorator
def ChatGPT_reply_content(drivermanager):
    try:
        # 輸出對話格式
        html_content = view_custormer_chat(drivermanager) # 生成html格式內容
        chatgpt_content = chatgpt_extract_conversations(html_content) # html轉換chatgpt格式內容
        
        # 調用ChatGPT回覆
        chatgpt_api_key = read_database('database.db', 'chatgpt_api_key')
        chatgpt_assistant_id = read_database('database.db', 'chatgpt_assistant_id')
        
        ChatGPT_Robot_reply = ChatGPT_Robot(chatgpt_api_key, chatgpt_assistant_id, chatgpt_content) # 接入chatgpt回覆
        print(ChatGPT_Robot_reply)
        return ChatGPT_Robot_reply
    except Exception:
        logging.exception('產生 ChatGPT 回覆失敗')
        print(traceback.format_exc())
        raise

@log_decorator
def reply_task(drivermanager, reply_type):
    try:
        global UI_STATUS_DATA
        # 進入至特定頁面、狀態
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        print("已刷新頁面，更新時間:" + now)
        drivermanager.driver.get("https://seller.shopee.tw/new-webchat/conversations")

        # 進入頁面後，優先切到『全部聊聊』，每次刷新都執行一次以確保在正確分頁
        try:
            if drivermanager.click_all_conversations(timeout=8):
                drivermanager.record_nav("reply_task: clicked all conversations")
        except Exception:
            pass

        # 伺服器錯誤/頁面未就緒 → 重試與重新加載
        def ensure_chat_ready(max_retry=4):
            for i in range(max_retry):
                try:
                    # 檢查是否出現伺服器錯誤提示
                    error_hint = None
                    try:
                        error_hint = drivermanager.wait_for_element(
                            drivermanager.driver,
                            By.XPATH,
                            "//*[contains(text(),'伺服器錯誤') or contains(text(),'服务器错误') or contains(text(),'Server Error')]",
                            timeout=2
                        )
                    except Exception:
                        pass

                    if error_hint:
                        # 嘗試按下重新加載按鈕
                        try:
                            reload_btn = drivermanager.wait_for_element(
                                drivermanager.driver,
                                By.XPATH,
                                "//button[contains(.,'重新加載') or contains(.,'重新加载') or contains(.,'Refresh')]",
                                timeout=2
                            )
                            reload_btn.click()
                            time.sleep(2)
                        except Exception:
                            drivermanager.driver.refresh()
                            time.sleep(2)

                    # 有任一聊天篩選器就算就緒
                    try:
                        filt_any = drivermanager.wait_for_element(
                            drivermanager.driver,
                            By.CSS_SELECTOR,
                            "[data-cy^='webchat-conversation-filter-filter-']",
                            timeout=5
                        )
                        if filt_any:
                            return True
                    except Exception:
                        pass

                    # 不就緒則刷新再試
                    drivermanager.driver.refresh()
                    time.sleep(2)
                except Exception:
                    time.sleep(1)
            return False

        if not ensure_chat_ready():
            print("聊天頁面載入失敗，稍後重試")
            return

        # 點擊篩選器：優先未回覆，否則全部，再否則以文字定位
        try:
            filt = drivermanager.wait_for_element(
                drivermanager.driver,
                By.CSS_SELECTOR,
                "[data-cy='webchat-conversation-filter-filter-new']",
                timeout=6
            )
        except Exception:
            try:
                filt = drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.CSS_SELECTOR,
                    "[data-cy='webchat-conversation-filter-filter-all']",
                    timeout=6
                )
            except Exception:
                filt = drivermanager.wait_for_element(
                    drivermanager.driver,
                    By.XPATH,
                    "//div[contains(text(),'未回覆') or contains(text(),'未回复') or contains(text(),'全部') or contains(text(),'All')]",
                    timeout=8
                )
        filt.click()
        try:
            container = drivermanager.wait_for_element(drivermanager.driver, By.CLASS_NAME, "ReactVirtualized__List")
        except Exception as e:
            print("無待回覆對話,結束回覆任務")
            return  # 退出函數

        # 初始數據
        whitelist = read_whitelist()
        old_chat_list = []
        new_chat_list = []
        is_scrolled_to_bottom = False

        # 開始載入對話
        while True:
            time.sleep(3)  # 等待頁面加載
            replies = drivermanager.wait_for_elements(container, By.CSS_SELECTOR, "[data-cy='webchat-conversation-cell-root']") # 找到所有客人元素
            for reply in replies:
                page_category = drivermanager.wait_for_element(drivermanager.driver, By.CSS_SELECTOR, "[data-cy='webchat-conversation-filter-filter-new']").text
                if "未回覆" in page_category:
                    customer_info_element = drivermanager.wait_for_element(reply, By.CSS_SELECTOR, "[data-cy='webchat-conversation-cell-name']") # 取得客人名稱元素
                    customer_name = customer_info_element.get_attribute('title') # 取得客人名稱

                    new_chat_list.append(customer_name)
                    # 使用集合操作找出新列表中獨有的元素
                    new_chat_items = set(new_chat_list) - set(old_chat_list)
                    new_chat_list.pop(0)
                    
                    # 對每個新對話元素執行操作並將其加入舊對話列表
                    if new_chat_items:
                        for item in new_chat_items:
                            old_chat_list.append(item) # 將新元素加入舊列表
                            # 如果舊列表長度超過50，則將最舊的元素去除
                            if len(old_chat_list) >= 50:
                                old_chat_list.pop(0)
                        
                            reply_need = answer_buyer_check(customer_name) #是否需要回覆
                            print(f"載入客人: {customer_name} 的對話視窗")
                            reply.click()
                            # 如果需要回覆
                            if reply_need:
                                
                                # 回覆客人
                                def send_text_with_shift_enter(textarea_element, text):
                                    parts = text.split('\n')
                                    for i, part in enumerate(parts):
                                        textarea_element.send_keys(part)
                                        if i < len(parts) - 1:
                                            textarea_element.send_keys(Keys.SHIFT, Keys.ENTER)
                                            
                                try:
                                    textarea_element = drivermanager.wait_for_element(drivermanager.driver, By.CLASS_NAME, "E2MWg3w8y6")
                                except:
                                    continue

                                if reply_type == 1:
                                    if UI_STATUS_DATA['standard_reply_status']:
                                        send_text_with_shift_enter(textarea_element, UI_STATUS_DATA['lunchbreak_text'])
                                    else:
                                        send_text_with_shift_enter(textarea_element, ChatGPT_reply_content(drivermanager))
                                if reply_type ==2:
                                    if UI_STATUS_DATA['standard_reply_status']:
                                        send_text_with_shift_enter(textarea_element, UI_STATUS_DATA['getoff_text'])
                                    else:
                                        send_text_with_shift_enter(textarea_element, ChatGPT_reply_content(drivermanager))
                                #drivermanager.wait_for_element(drivermanager.driver, By.CSS_SELECTOR, "i.GHUxSkxNuJ.yHRqJXUiCY > svg.chat-icon > path").click()
                                print(f'已回覆客人: {customer_name}')
                                update_whitelist(customer_name) # 更新回覆名單

                # 滾動處理（無論哪個過濾器都需要滾動）
                # 對指定元素進行向下滾動
                drivermanager.driver.execute_script('arguments[0].scrollTop = arguments[0].scrollTop + 50;', container)
                
                # 獲取元素的滾動高度、元素的可視高度和當前滾動的位置
                scroll_height = drivermanager.driver.execute_script("return arguments[0].scrollHeight", container)
                client_height = drivermanager.driver.execute_script("return arguments[0].clientHeight", container)
                scroll_top = drivermanager.driver.execute_script("return arguments[0].scrollTop", container)
                
                # 檢查是否已經滾動到底部
                is_scrolled_to_bottom = scroll_top + client_height >= scroll_height
            else:
                print("不正確的頁面")
            if is_scrolled_to_bottom:
                print("瀏覽完畢")
                break
    except Exception:
        logging.exception('回覆任務執行失敗')
        print(traceback.format_exc())
        raise


@log_decorator
def Customer_Serivce(drivermanager, ui_status_signals):
    ui_status_signals.request_status.emit()

    global LOGIN_STATE, UI_STATUS_DATA
    while UI_STATUS_DATA == None:
        time.sleep(1)
    
    if LOGIN_STATE == 1:
        now = datetime.datetime.now()
        current_time = QTime(now.hour, now.minute, now.second)
        weekday = now.weekday()  # 0 是星期一, 6 是星期日
    
        # 自訂休息日列表，格式為 '年/月/日'
        custom_holidays = ['2024/06/10']  # 聖誕節和元旦
    
        # 檢查今天是否為自訂休息日
        is_custom_holiday = now.strftime('%Y/%m/%d') in custom_holidays

        # 檢查是否為工作日
        if UI_STATUS_DATA['workday_checkboxes'][weekday] and not is_custom_holiday:
            lunchbreak_start = UI_STATUS_DATA['lunchbreak_starttime']
            lunchbreak_end = UI_STATUS_DATA['lunchbreak_endtime']
            getoff_start = UI_STATUS_DATA['getoff_starttime']
            getoff_end = UI_STATUS_DATA['getoff_endtime']
    
            #reply_type=0為非上班時間回覆,1為智能回覆,2為智能不回覆轉接人工回覆
            if lunchbreak_start <= current_time < lunchbreak_end:
                # 午休時間
                print("午休時間，智能客服接入處理")
                reply_task(drivermanager, 1)
                
            elif current_time < lunchbreak_start or (current_time >= lunchbreak_end and current_time < getoff_start):
                # 工作時間
                print("人工客服接入處理")
                reply_task(drivermanager, 0)
            else:
                # 非工作時間或下班後
                print("非工作時間/下班後，智能客服接入處理")
                reply_task(drivermanager, 2)
        else:
            # 週末或自訂休息日
            print("周末或自訂休息日，智能客服接入處理")
            reply_task(drivermanager, 2)
        
    UI_STATUS_DATA = None

# 主執行緒類
class MainThread(Thread):
    def __init__(self, drivermanager, login_status, ui_status_signals):
        super().__init__()
        self.drivermanager = drivermanager
        self.login_status = login_status
        self.ui_status_signals = ui_status_signals
        
    def run(self):
        # 主程式入口
        try:
            global ACCOUNT_NUMBER, ACCOUNT_PASSWORD
            autologin(self.drivermanager, self.login_status, ACCOUNT_NUMBER, ACCOUNT_PASSWORD) #開始登入
            Customer_Serivce(self.drivermanager, self.ui_status_signals) #開始服務
            scheduler = BlockingScheduler()
            scheduler.add_job(Customer_Serivce, 'interval', minutes=1,args=[self.drivermanager, self.ui_status_signals])
            scheduler.start()
        
        except Exception as e:
            logging.exception('主執行緒流程發生錯誤')
            try:
                self.login_status("發生錯誤")
            except:
                pass
            print(traceback.format_exc())


# 函數來儲存設置
@log_decorator
def save_settings(user_account, user_password, file_password):
    settings = QSettings('MyCompany', 'MyApp')
    settings.setValue('user_account', user_account)
    settings.setValue('user_password', user_password)
    settings.setValue('file_password', file_password)

# 函數來讀取設置
@log_decorator
def load_settings():
    settings = QSettings('MyCompany', 'MyApp')
    user_account = settings.value('user_account')
    user_password = settings.value('user_password')
    file_password = settings.value('file_password')
    return user_account, user_password, file_password

# 定義一個信號類，用於在主執行緒中執行操作
class UISignals(QObject):
    request_status = pyqtSignal()
    status_updated = pyqtSignal()

class EmittingStream(QObject):
    text_written = pyqtSignal(str)  # 定義信號

    def write(self, text):
        self.text_written.emit(text)  # 發射信號

    def flush(self):
        pass  # 對於這個例子，flush 方法可以不做任何事情
        
    def __del__(self):
        # 處理對象被刪除時的情況
        try:
            if hasattr(self, 'text_written'):
                self.text_written.disconnect()
        except:
            pass

#主視窗
class MyMainWindow(QMainWindow):
    def __init__(self):
        super(MyMainWindow, self).__init__()
        uic.loadUi('Shopee_Customer_Service.ui', self)
        # 將按鈕的點擊事件連接到自定義的槽函數
        self.Shoppe_Login_Button.clicked.connect(self.Shoppe_Login_Button_on_click)

        self.ui_status_signals = UISignals() #初始一個訊號實例
        self.ui_status_signals.request_status.connect(self.get_status) #連接訊號觸發函數

        # 預設輸出到 GUI；若需要顯示在 CMD，請將環境變數 STD_TO_GUI 設為 0
        redirect_to_gui = os.environ.get("STD_TO_GUI", "1") == "1"
        if redirect_to_gui:
            # 重定向 print 到 QTextBrowser
            sys.stdout = EmittingStream()
            sys.stdout.text_written.connect(self.append_text)
            sys.stderr = EmittingStream()
            sys.stderr.text_written.connect(self.append_text)

        self.read_settings()  #恢復 UI 狀態

    def Shoppe_Login_Button_on_click(self):
        self.drivermanager = WebDriverManager() #實例一個瀏覽器
        global LOGIN_STATE
        LOGIN_STATE = 0
        self.login_status("登入中") # 更新標籤的文本為'登入中'
        # 從控件獲取帳號和密碼
        global ACCOUNT_NUMBER,ACCOUNT_PASSWORD
        ACCOUNT_NUMBER = self.Shoppe_account_Input.text()
        ACCOUNT_PASSWORD = self.Shoppe_Password_Input.text()
        mainthread = MainThread(self.drivermanager, self.login_status, self.ui_status_signals) # 創建執行緒對象，並傳遞帳號、密碼和更新狀態的函數
        mainthread.start() # 啟動執行緒
        

    def login_status(self, status):
        # 更新標籤的文本為登入狀態
        self.Shoppe_LoginState_Text2.setText(status)
        if LOGIN_STATE == 1:
            self.ReplySetting_Frame.setEnabled(True)
        else:
            self.ReplySetting_Frame.setEnabled(False)
            
    def get_status(self):
        # 在這裡讀取 UI 元件的狀態
        global UI_STATUS_DATA
        UI_STATUS_DATA = {"shoppe_account": self.Shoppe_account_Input.text(),
                        "shoppe_password": self.Shoppe_Password_Input.text(),
                        "workday_checkboxes": [self.Reply_Workday_checkBox_1.isChecked(),
                         self.Reply_Workday_checkBox_2.isChecked(),
                         self.Reply_Workday_checkBox_3.isChecked(),
                         self.Reply_Workday_checkBox_4.isChecked(),
                         self.Reply_Workday_checkBox_5.isChecked(),
                         self.Reply_Workday_checkBox_6.isChecked(),
                         self.Reply_Workday_checkBox_7.isChecked(),],
                        "standard_reply_status": self.Standard_Reply_radioButton.isChecked(),
                        "lunchbreak_text": self.Reply_Lunchbreak_Input.toPlainText(),
                        "lunchbreak_starttime": self.Reply_LunchbreakStart_timeEdit.time(),
                        "lunchbreak_endtime": self.Reply_LunchbreakEnd_timeEdit.time(),
                        "intelligent_reply_status": self.Intelligent_Reply_radioButton.isChecked(),
                        "getoff_text": self.Reply_Getoff_Input.toPlainText(),
                        "getoff_starttime": self.Reply_GetoffStart_timeEdit.time(),
                        "getoff_endtime": self.Reply_GetoffEnd_timeEdit.time(),
        }
        self.write_settings()  #保存 UI 狀態

    @pyqtSlot(str)
    def append_text(self, text):
        # 將文本添加到 QTextBrowser
        cursor = self.State_textBrowser.textCursor() # 獲取 QTextBrowser 的 QTextCursor 對象
        cursor.movePosition(cursor.End) # 移動光標到文本末尾
        cursor.insertText(text) # 插入文本
        self.State_textBrowser.setTextCursor(cursor) # 確保 QTextBrowser 更新到新的光標位置
        QApplication.processEvents()  # 處理所有待處理的事件並更新 GUI

    def write_settings(self):
        settings = QSettings("MyCompany", "MyApp")
        settings.beginGroup("MainWindow")
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("state", self.saveState())
        
        # 將 QTime 對象轉換為字符串
        ui_status_data = UI_STATUS_DATA.copy()
        ui_status_data["lunchbreak_starttime"] = ui_status_data["lunchbreak_starttime"].toString()
        ui_status_data["lunchbreak_endtime"] = ui_status_data["lunchbreak_endtime"].toString()
        ui_status_data["getoff_starttime"] = ui_status_data["getoff_starttime"].toString()
        ui_status_data["getoff_endtime"] = ui_status_data["getoff_endtime"].toString()
        
        settings.setValue("UI_STATUS_DATA", json.dumps(ui_status_data))
        settings.endGroup()


    def read_settings(self):
        settings = QSettings("MyCompany", "MyApp")
        settings.beginGroup("MainWindow")
        geometry = settings.value("geometry")
        if geometry is not None:
            self.restoreGeometry(geometry)
        state = settings.value("state")
        if state is not None:
            self.restoreState(state)
        ui_status_data_str = settings.value("UI_STATUS_DATA", "{}")
        ui_status_data = json.loads(ui_status_data_str)

        # 將字符串轉換回 QTime 對象
        if "lunchbreak_starttime" in ui_status_data:
            ui_status_data["lunchbreak_starttime"] = QTime.fromString(ui_status_data["lunchbreak_starttime"], "HH:mm")
        if "lunchbreak_endtime" in ui_status_data:
            ui_status_data["lunchbreak_endtime"] = QTime.fromString(ui_status_data["lunchbreak_endtime"], "HH:mm")
        if "getoff_starttime" in ui_status_data:
            ui_status_data["getoff_starttime"] = QTime.fromString(ui_status_data["getoff_starttime"], "HH:mm")
        if "getoff_endtime" in ui_status_data:
            ui_status_data["getoff_endtime"] = QTime.fromString(ui_status_data["getoff_endtime"], "HH:mm")

        self.apply_ui_status(ui_status_data)
        settings.endGroup()



    def apply_ui_status(self, ui_status_data):
        if ui_status_data:
            self.Shoppe_account_Input.setText(ui_status_data.get("shoppe_account", ""))
            self.Shoppe_Password_Input.setText(ui_status_data.get("shoppe_password", ""))
            
            workday_checkboxes = ui_status_data.get("workday_checkboxes", [])
            if len(workday_checkboxes) == 7:
                self.Reply_Workday_checkBox_1.setChecked(workday_checkboxes[0])
                self.Reply_Workday_checkBox_2.setChecked(workday_checkboxes[1])
                self.Reply_Workday_checkBox_3.setChecked(workday_checkboxes[2])
                self.Reply_Workday_checkBox_4.setChecked(workday_checkboxes[3])
                self.Reply_Workday_checkBox_5.setChecked(workday_checkboxes[4])
                self.Reply_Workday_checkBox_6.setChecked(workday_checkboxes[5])
                self.Reply_Workday_checkBox_7.setChecked(workday_checkboxes[6])

            self.Standard_Reply_radioButton.setChecked(ui_status_data.get("standard_reply_status", False))
            self.Reply_Lunchbreak_Input.setPlainText(ui_status_data.get("lunchbreak_text", ""))
            self.Reply_LunchbreakStart_timeEdit.setTime(ui_status_data.get("lunchbreak_starttime", "00:00"))
            self.Reply_LunchbreakEnd_timeEdit.setTime(ui_status_data.get("lunchbreak_endtime", "00:00"))
            self.Intelligent_Reply_radioButton.setChecked(ui_status_data.get("intelligent_reply_status", False))
            self.Reply_Getoff_Input.setPlainText(ui_status_data.get("getoff_text", ""))
            self.Reply_GetoffStart_timeEdit.setTime(ui_status_data.get("getoff_starttime", "00:00"))
            self.Reply_GetoffEnd_timeEdit.setTime(ui_status_data.get("getoff_endtime", "00:00"))
        
    def closeEvent(self, event):
        # 重置 sys.stdout 和 sys.stderr
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        self.get_status()
        self.write_settings()  #保存 UI 狀態
        event.accept()
        
if __name__ == "__main__":
    global FILE_PATH
    
    # 設定 Qt plugin 路徑
    try:
        # 優先支援 conda 環境：使用 platforms 目錄並清除 QT_PLUGIN_PATH 以避免衝突
        conda_prefix = os.environ.get("CONDA_PREFIX")
        if conda_prefix:
            qpa_path = os.path.join(conda_prefix, "Library", "plugins", "platforms")
            if os.path.exists(qpa_path):
                os.environ.pop("QT_PLUGIN_PATH", None)
                os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = qpa_path
                print(f"已設定 Qt 平台外掛路徑: {qpa_path}")
        else:
            # venv/一般 pip 安裝時的預設位置
            from pathlib import Path
            venv_path = Path(sys.prefix)
            qpa_path = venv_path / "Lib" / "site-packages" / "PyQt5" / "Qt5" / "plugins" / "platforms"
            if qpa_path.exists():
                os.environ.pop("QT_PLUGIN_PATH", None)
                os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = str(qpa_path)
                print(f"已設定 Qt 平台外掛路徑: {qpa_path}")
    except Exception as e:
        print(f"設定 Qt plugin 路徑時發生錯誤: {e}")
    
    documents_path = os.path.join(os.path.expanduser('~'), 'Documents')
    folder_name = '蝦皮聊聊智能客服'
    FILE_PATH = os.path.join(documents_path, folder_name)
    if not os.path.exists(FILE_PATH):
        os.makedirs(FILE_PATH)
    logging.basicConfig(filename=os.path.join(FILE_PATH, 'app.log'), level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    expected_serial = read_database('database.db', 'expected_serial')
    if verify_cpu_serial(expected_serial):
        print("硬體ID驗證成功，開始執行程式。")
        app = QApplication([])
        mainWindow = MyMainWindow()
        mainWindow.show()
        app.exec_()
    else:
        print("硬體ID驗證失敗，不允許執行程式。")