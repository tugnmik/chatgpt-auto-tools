"""
ChatGPT Auto Registration - MULTITHREADING VERSION
Tự động đăng ký nhiều tài khoản ChatGPT đồng thời
"""

import requests
import time
import re
import sys
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import random
import string
import json
import os
import threading
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError as FutureTimeoutError
from colorama import init, Fore, Style
import warnings
import traceback
import tempfile
import shutil
import pyotp
from openpyxl import Workbook, load_workbook
import customtkinter as ctk
from tkinter import messagebox, filedialog
from datetime import datetime

# Initialize colorama
init(autoreset=True)

# Locks for thread-safe operations
file_lock = threading.Lock()
print_lock = threading.Lock()
driver_init_lock = threading.Lock()

# Chrome major version installed locally (update when Chrome updates)
CHROME_VERSION_MAIN = 144

# Global flag for getting checkout link (set by menu)
GET_CHECKOUT_LINK = False
GET_CHECKOUT_TYPE = "Plus"  # Options: "Plus", "Business", "Both"
NETWORK_MODE = "Fast"  # Options: "Fast" (stable network), "VPN/Slow" (unstable network)

# Network mode settings
NETWORK_SETTINGS = {
    "Fast": {
        "page_load_timeout": 15,
        "element_timeout": 5,
        "max_retries": 1,
        "retry_delay": 1,
        "extra_wait": 0.5,
    },
    "VPN/Slow": {
        "page_load_timeout": 20,
        "element_timeout": 7,
        "max_retries": 2,
        "retry_delay": 1.5,
        "extra_wait": 0.8,
    }
}

# Default password for registration (editable via GUI)
DEFAULT_PASSWORD = "Matkhau123!@#"

# Define colors
class Colors:
    SUCCESS = Fore.GREEN
    ERROR = Fore.RED
    WARNING = Fore.YELLOW
    INFO = Fore.CYAN
    HEADER = Fore.MAGENTA
    RESET = Style.RESET_ALL


def safe_print(thread_id, message, color=Colors.INFO, emoji=""):
    """Thread-safe print with thread ID"""
    with print_lock:
        print(f"[T{thread_id}] {emoji}{color}{message}{Colors.RESET}")


class TempMailAPI:
    """API client cho tinyhost.shop"""
    
    def __init__(self):
        self.base_url = "https://tinyhost.shop"
    
    def get_random_domains(self, limit=10):
        """Lấy danh sách domain ngẫu nhiên"""
        try:
            url = f"{self.base_url}/api/random-domains/"
            params = {"limit": limit}
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            return None
    
    def get_emails(self, domain, user, page=1, limit=20):
        """Lấy danh sách email"""
        try:
            url = f"{self.base_url}/api/email/{domain}/{user}/"
            params = {"page": page, "limit": limit}
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            return None
    
    def get_email_detail(self, domain, user, email_id):
        """Lấy chi tiết một email"""
        try:
            url = f"{self.base_url}/api/email/{domain}/{user}/{email_id}"
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            return response.json()
        except Exception as e:
            return None
    
    def generate_email(self):
        """Tạo một email ngẫu nhiên"""
        domains_data = self.get_random_domains(20)
        if not domains_data or not domains_data.get('domains'):
            return None
        
        domain = random.choice(domains_data['domains'])
        username = ''.join(random.choices(string.ascii_lowercase + string.digits, k=10))
        email = f"{username}@{domain}"
        
        return {
            'email': email,
            'username': username,
            'domain': domain
        }


class DongVanOAuth2API:
    """API client cho DongVan OAuth2 (thay thế tinyhost.shop)"""
    
    API_MESSAGES_URL = "https://tools.dongvanfb.net/api/get_messages_oauth2"
    
    def __init__(self, email, password, refresh_token, client_id):
        self.email = email
        self.password = password
        self.refresh_token = refresh_token
        self.client_id = client_id
    
    def fetch_messages(self):
        """Lấy danh sách email từ OAuth2 API"""
        payload = {
            "email": self.email,
            "refresh_token": self.refresh_token,
            "client_id": self.client_id,
        }
        try:
            response = requests.post(self.API_MESSAGES_URL, json=payload, timeout=20)
            response.raise_for_status()
            return response.json()
        except requests.RequestException:
            return None
        except json.JSONDecodeError:
            return None
    
    def extract_code_from_messages(self, messages_payload):
        """Trích xuất code 6 số từ thư mới nhất"""
        if not messages_payload:
            return None
        
        messages = messages_payload.get("messages")
        if not isinstance(messages, list) or not messages:
            return None
        
        pattern = re.compile(r"\b(\d{6})\b")
        
        def parse_msg_datetime(raw):
            if not raw:
                return datetime.min
            try:
                return datetime.fromisoformat(raw.replace("Z", "+00:00"))
            except ValueError:
                pass
            for fmt in ("%H:%M - %d/%m/%Y", "%d/%m/%Y %H:%M:%S"):
                try:
                    return datetime.strptime(raw, fmt)
                except ValueError:
                    continue
            return datetime.min
        
        # Sắp xếp thư mới nhất lên đầu
        sorted_messages = sorted(
            messages,
            key=lambda msg: parse_msg_datetime(msg.get("date")),
            reverse=True,
        )
        
        if not sorted_messages:
            return None
        
        latest_msg = sorted_messages[0]
        
        # Ưu tiên 1: Trường 'code' nếu có
        code_field = latest_msg.get("code", "")
        if code_field and pattern.match(str(code_field)):
            return code_field
        
        # Ưu tiên 2: Extract từ subject line
        subject = latest_msg.get("subject") or ""
        subject_codes = pattern.findall(subject)
        if subject_codes:
            return subject_codes[0]
        
        # Ưu tiên 3: Extract từ content/message
        content = latest_msg.get("content") or latest_msg.get("message") or ""
        content_codes = pattern.findall(content)
        if content_codes:
            return content_codes[0]
        
        return None
    
    def get_email_info(self):
        """Trả về thông tin email theo format tương thích"""
        return {
            'email': self.email,
            'username': self.email.split('@')[0],
            'domain': self.email.split('@')[1] if '@' in self.email else ''
        }


# Global list để quản lý tài khoản OAuth2
oauth2_accounts = []
current_account_index = 0
account_lock = threading.Lock()


def load_oauth2_accounts_from_excel(file_path="oauth2.xlsx", skip_registered=True):
    """Load tài khoản OAuth2 từ file Excel oauth2.xlsx
    Cột A: email|password|refresh_token|client_id
    Cột B: Status (registered = đã đăng ký, trống = chưa đăng ký)
    Row 1 là header, dữ liệu từ row 2
    
    Args:
        file_path: Path to oauth2.xlsx
        skip_registered: If True, skip accounts with status='registered' (for Registration)
                        If False, load all accounts (for MFA/Checkout lookup)
    """
    accounts = []
    
    if not os.path.exists(file_path):
        return accounts
    
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        
        skipped_count = 0
        
        for row_num in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_num, column=1).value
            if not cell_value:
                continue
            
            line = str(cell_value).strip()
            if not line or line.startswith("#"):
                continue
            
            parts = [part.strip() for part in line.split("|")]
            if len(parts) != 4:
                continue
            
            # Kiểm tra cột B (Status) - skip nếu đã registered (chỉ khi skip_registered=True)
            if skip_registered:
                status_value = ws.cell(row=row_num, column=2).value
                if status_value and str(status_value).strip().lower() == "registered":
                    skipped_count += 1
                    continue
            
            email, password, refresh_token, client_id = parts
            accounts.append({
                "email": email,
                "password": password,
                "refresh_token": refresh_token,
                "client_id": client_id,
                "row_num": row_num,
                "used": False
            })
        
        wb.close()
        
        if skipped_count > 0 and skip_registered:
            print(f"⏭️ Skipped {skipped_count} registered accounts")
        
    except Exception as e:
        print(f"Error loading OAuth2 accounts: {e}")
        return []
    
    return accounts


def get_next_oauth2_account():
    """Lấy tài khoản OAuth2 tiếp theo chưa được sử dụng (thread-safe)"""
    global oauth2_accounts
    
    with account_lock:
        for i, account in enumerate(oauth2_accounts):
            if not account.get("used", False):
                oauth2_accounts[i]["used"] = True
                return account
    return None


def reset_oauth2_accounts():
    """Reset trạng thái used của tất cả oauth2 accounts"""
    global oauth2_accounts
    with account_lock:
        for i in range(len(oauth2_accounts)):
            oauth2_accounts[i]["used"] = False


def mark_oauth2_registered(row_num, file_path="oauth2.xlsx"):
    """Ghi 'registered' vào cột B của oauth2.xlsx sau khi đăng ký thành công
    Thread-safe using file_lock
    """
    with file_lock:
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            ws.cell(row=row_num, column=2, value="registered")
            wb.save(file_path)
            wb.close()
            return True
        except Exception as e:
            print(f"Error marking OAuth2 account as registered: {e}")
            return False


class ChatGPTAutoRegisterWorker:
    """Worker thread for ChatGPT registration"""
    
    def __init__(self, thread_id, num_threads=1, email_mode="TinyHost", oauth2_account=None):
        self.thread_id = thread_id
        self.num_threads = num_threads  # Total threads for position calculation
        self.email_mode = email_mode  # "TinyHost" or "OAuth2"
        self.oauth2_account = oauth2_account  # OAuth2 account data if mode is OAuth2
        
        # Initialize mail API based on mode
        if email_mode == "OAuth2" and oauth2_account:
            self.mail_api = DongVanOAuth2API(
                email=oauth2_account["email"],
                password=oauth2_account["password"],
                refresh_token=oauth2_account["refresh_token"],
                client_id=oauth2_account["client_id"]
            )
            self.email_info = self.mail_api.get_email_info()
            self.password = DEFAULT_PASSWORD  # Use user-configured password, not oauth2 email password
            self.oauth2_row_num = oauth2_account.get("row_num")
        else:
            self.mail_api = TempMailAPI()
            self.email_info = None
            self.password = DEFAULT_PASSWORD
            self.oauth2_row_num = None
        
        self.driver = None
        self.user_data_dir = None
        self.stop_event = None
        self.operation_timeout_detected = False
        self.max_timeout_retries = 2
        self.current_retry = 0
        
        # Network settings based on mode
        self.net = NETWORK_SETTINGS.get(NETWORK_MODE, NETWORK_SETTINGS["Fast"])
        
    def log(self, message, color=Colors.INFO, emoji=""):
        """Log with thread ID"""
        safe_print(self.thread_id, message, color, emoji)

    
    def cleanup_browser(self):
        """Close browser and remove temporary profile"""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
        if self.user_data_dir and os.path.exists(self.user_data_dir):
            try:
                shutil.rmtree(self.user_data_dir, ignore_errors=True)
            except Exception:
                pass
            self.user_data_dir = None
    
    def has_operation_timeout_error(self):
        """Check if the current page shows the OpenAI operation timeout error"""
        if not self.driver:
            return False
        try:
            body = self.driver.find_element(By.TAG_NAME, "body")
            text = body.text.lower()
            return (
                "operation timed out" in text
                or "oops, an error occurred" in text
                or "rất tiếc, đã xảy ra lỗi không xác định" in text
                or ("thử lại" in text and "operation timed out" in text)
                or "something went wrong" in text
                or ("an error occurred" in text and "try again" in text)
            )
        except Exception:
            return False
    
    def click_try_again_button(self):
        """Click 'Try again' button on error page if present"""
        if not self.driver:
            return False
        try:
            selectors = [
                "//button[contains(text(), 'Try again')]",
                "//button[contains(text(), 'Thử lại')]",
                "//button[contains(@class, 'btn') and contains(text(), 'Try')]",
                "//a[contains(text(), 'Try again')]",
            ]
            for selector in selectors:
                try:
                    btn = self.driver.find_element(By.XPATH, selector)
                    if btn and btn.is_displayed():
                        btn.click()
                        self.log("Clicked 'Try again' button", Colors.INFO, "🔄 ")
                        time.sleep(2)
                        return True
                except:
                    continue
            return False
        except:
            return False
    
    def redirected_to_chatgpt_home(self, timeout_seconds=5):
        """Check if browser navigated to chatgpt.com within a short timeout"""
        if not self.driver:
            return False
        try:
            wait = WebDriverWait(self.driver, timeout_seconds)
            wait.until(lambda d: "chatgpt.com" in (d.current_url or "").lower() and "/auth" not in (d.current_url or "").lower())
            return True
        except Exception:
            return False
    
    def check_and_handle_timeout(self):
        """Check for timeout error and return True if detected (caller should retry)"""
        if self.has_operation_timeout_error():
            self.operation_timeout_detected = True
            self.log("Detected 'Operation timed out' error!", Colors.WARNING, "⚠️ ")
            # Try to click "Try again" button
            self.click_try_again_button()
            return True
        return False
    
    def setup_driver(self, max_retries=3):
        """Initialize undetected Chrome driver with retry"""
        for attempt in range(max_retries):
            try:
                if self.stop_event and self.stop_event.is_set():
                    return False
                    
                self.log(f"Initializing ChromeDriver (attempt {attempt + 1}/{max_retries})...", Colors.INFO, "🔄 ")
                
                options = uc.ChromeOptions()
                
                # Create separate user-data-dir for each thread
                timestamp = int(time.time() * 1000)
                random_suffix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=6))
                self.user_data_dir = tempfile.mkdtemp(prefix=f"chrome_chatgpt_T{self.thread_id}_{timestamp}_{random_suffix}_")
                options.add_argument(f'--user-data-dir={self.user_data_dir}')
                
                # Basic options
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                options.add_argument('--window-size=900,700')
                options.add_argument('--disable-blink-features=AutomationControlled')
                options.add_argument('--lang=en-US')  # Set browser language to English (MM/DD/YYYY format)
            
                
                # Disable notifications
                prefs = {
                    "credentials_enable_service": False,
                    "profile.password_manager_enabled": False,
                    "profile.default_content_setting_values.notifications": 2
                }
                options.add_experimental_option("prefs", prefs)
                
                if attempt > 0:
                    delay = attempt * 2
                    self.log(f"Waiting {delay}s before retry...", Colors.WARNING, "⏳ ")
                    time.sleep(delay)
                
                # Lock when initializing
                with driver_init_lock:
                    try:
                        self.driver = uc.Chrome(options=options, version_main=CHROME_VERSION_MAIN)
                    except Exception as version_err:
                        self.log(f"Version-specific ChromeDriver failed ({version_err}); retrying with auto.", Colors.WARNING, "⚠️ ")
                        self.driver = uc.Chrome(options=options, version_main=None)
                
                # Set size and position based on thread_id (cascade windows)
                # Use modulo to reset position for each batch of threads
                window_width = 900
                window_height = 900
                position_index = (self.thread_id - 1) % self.num_threads if self.num_threads > 1 else 0
                x_offset = position_index * 100  # 100px offset per thread
                y_offset = position_index * 50   # 50px offset per thread
                
                self.driver.set_window_size(window_width, window_height)
                self.driver.set_window_position(x_offset, y_offset)
                
                self.log("ChromeDriver initialized!", Colors.SUCCESS, "✅ ")
                return True
                
            except Exception as e:
                self.log(f"Error initializing driver (attempt {attempt + 1}): {e}", Colors.WARNING, "⚠️ ")
                
                if self.driver:
                    try:
                        self.driver.quit()
                    except:
                        pass
                    self.driver = None
                
                if self.user_data_dir and os.path.exists(self.user_data_dir):
                    try:
                        shutil.rmtree(self.user_data_dir, ignore_errors=True)
                    except:
                        pass
                    self.user_data_dir = None
                
                if attempt < max_retries - 1:
                    if self.stop_event and self.stop_event.is_set():
                        return False
                    continue
                else:
                    self.log("Could not initialize ChromeDriver", Colors.ERROR, "❌ ")
                    return False
        
        return False
    
    def human_like_delay(self, min_seconds=0.3, max_seconds=0.8):
        """Random delay (optimized)"""
        time.sleep(random.uniform(min_seconds, max_seconds))
    
    def human_like_typing(self, element, text):
        """Type with human-like speed"""
        for char in text:
            element.send_keys(char)
            time.sleep(random.uniform(0.01, 0.04))
    
    def move_mouse_randomly(self):
        """Move mouse randomly (fast)"""
        try:
            action = ActionChains(self.driver)
            for _ in range(random.randint(1, 2)):
                x_offset = random.randint(-30, 30)
                y_offset = random.randint(-30, 30)
                action.move_by_offset(x_offset, y_offset).perform()
                time.sleep(random.uniform(0.05, 0.15))
        except:
            pass
    
    def navigate_to_chatgpt(self):
        """Navigate to ChatGPT with retry based on network mode"""
        max_retries = self.net["max_retries"]
        
        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    self.log(f"Retry accessing ChatGPT ({attempt + 1}/{max_retries})...", Colors.WARNING, "🔄 ")
                    time.sleep(self.net["retry_delay"])
                
                self.driver.get("https://chatgpt.com/")
                
                # Wait for page to fully load
                timeout = self.net["page_load_timeout"] + (attempt * 5)
                wait = WebDriverWait(self.driver, timeout)
                wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
                
                # Additional check: wait for body to have content
                try:
                    WebDriverWait(self.driver, self.net["element_timeout"]).until(
                        lambda d: len(d.find_elements(By.TAG_NAME, "button")) > 0
                    )
                except Exception:
                    pass
                
                self.log("Accessed chatgpt.com", Colors.SUCCESS, "✅ ")
                time.sleep(self.net["extra_wait"])
                return True
                
            except TimeoutException:
                self.log(f"Page load timeout (attempt {attempt + 1})", Colors.WARNING, "⏳ ")
                if attempt < max_retries - 1:
                    continue
                    
            except Exception as e:
                self.log(f"Error accessing ChatGPT: {e}", Colors.ERROR, "❌ ")
                if attempt < max_retries - 1:
                    continue
        
        self.log(f"Failed to access ChatGPT after {max_retries} attempts", Colors.ERROR, "❌ ")
        return False
    
    def click_login_button(self):
        """Click Log in button with retry based on network mode"""
        try:
            max_retries = self.net["max_retries"]
            
            selectors = [
                "//button[@data-testid='login-button']",
                "//button[contains(text(), 'Log in')]",
                "//button[contains(text(), 'Đăng nhập')]",
                "//a[contains(text(), 'Log in')]",
                "//a[contains(text(), 'Đăng nhập')]",
                "//button[contains(@class, 'login')]",
                "//*[@data-testid='login-button']",
            ]
            
            login_button = None
            
            for attempt in range(max_retries):
                timeout = self.net["element_timeout"] + (attempt * 3)
                wait = WebDriverWait(self.driver, timeout)
                
                if attempt > 0:
                    self.log(f"Retry finding Login button ({attempt + 1}/{max_retries})...", Colors.WARNING, "🔄 ")
                    time.sleep(self.net["retry_delay"])
                    try:
                        page_state = self.driver.execute_script("return document.readyState")
                        if page_state != "complete":
                            self.driver.refresh()
                            time.sleep(self.net["retry_delay"])
                    except Exception:
                        pass
                
                for selector in selectors:
                    try:
                        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        if login_button.is_displayed():
                            break
                        else:
                            login_button = None
                    except TimeoutException:
                        continue
                    except Exception:
                        continue
                
                if login_button:
                    break
            
            if not login_button:
                self.log("Could not find Log in button", Colors.ERROR, "❌ ")
                return False
            
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", login_button)
            self.human_like_delay(0.2, 0.4)
            
            # Try multiple click methods
            click_success = False
            try:
                self.move_mouse_randomly()
                ActionChains(self.driver).move_to_element(login_button).pause(0.3).click().perform()
                click_success = True
            except Exception:
                pass
            
            if not click_success:
                try:
                    login_button.click()
                    click_success = True
                except Exception:
                    pass
            
            if not click_success:
                try:
                    self.driver.execute_script("arguments[0].click();", login_button)
                    click_success = True
                except Exception:
                    pass
            
            if click_success:
                self.log("Clicked Log in", Colors.SUCCESS, "✅ ")
                time.sleep(self.net["extra_wait"])
                return True
            else:
                self.log("Failed to click Log in button", Colors.ERROR, "❌ ")
                return False
            
        except Exception as e:
            self.log(f"Error clicking Log in: {e}", Colors.ERROR, "❌ ")
            return False
    
    def enter_email_address(self):
        """Enter email with retry based on network mode"""
        try:
            # For OAuth2 mode, email_info is already set in __init__
            # For TinyHost mode, we need to generate email
            if not self.email_info:
                self.email_info = self.mail_api.generate_email()
                if not self.email_info:
                    return False
            
            email = self.email_info['email']
            self.log(f"Email: {email}", Colors.INFO, "📧 ")
            
            max_retries = self.net["max_retries"]
            
            selectors = [
                (By.ID, "email"),
                (By.XPATH, "//input[@id='email']"),
                (By.XPATH, "//input[@aria-label='Địa chỉ email']"),
                (By.XPATH, "//input[@aria-label='Email address']"),
                (By.XPATH, "//input[@type='email']"),
                (By.XPATH, "//input[@name='email']"),
                (By.XPATH, "//input[@placeholder='Email address']"),
                (By.XPATH, "//input[contains(@class, 'email')]"),
            ]
            
            email_input = None
            
            for attempt in range(max_retries):
                timeout = self.net["element_timeout"] + (attempt * 3)
                wait = WebDriverWait(self.driver, timeout)
                
                if attempt > 0:
                    self.log(f"Retry finding email input ({attempt + 1}/{max_retries})...", Colors.WARNING, "🔄 ")
                    time.sleep(self.net["retry_delay"])
                    
                    # Check if page needs refresh (stuck on loading)
                    try:
                        page_state = self.driver.execute_script("return document.readyState")
                        if page_state != "complete":
                            self.log("Page not fully loaded, waiting...", Colors.INFO, "⏳ ")
                            WebDriverWait(self.driver, self.net["page_load_timeout"]).until(
                                lambda d: d.execute_script("return document.readyState") == "complete"
                            )
                    except Exception:
                        pass
                
                # Try each selector
                for by, selector in selectors:
                    try:
                        email_input = wait.until(EC.presence_of_element_located((by, selector)))
                        if email_input.is_displayed():
                            break
                        else:
                            email_input = None
                    except TimeoutException:
                        continue
                    except Exception:
                        continue
                
                if email_input:
                    break
                    
                # If not found and not last attempt, try refreshing the login page
                if attempt < max_retries - 1:
                    self.log("Email input not found, refreshing page...", Colors.WARNING, "🔄 ")
                    try:
                        self.driver.get("https://auth.openai.com/authorize")
                        time.sleep(self.net["retry_delay"])
                    except Exception:
                        self.driver.get("https://chatgpt.com")
                        time.sleep(self.net["retry_delay"])
                        self.click_login_button()
                        time.sleep(self.net["extra_wait"])
            
            if not email_input:
                self.log(f"Could not find email input after {max_retries} attempts", Colors.ERROR, "❌ ")
                try:
                    page_title = self.driver.title
                    current_url = self.driver.current_url
                    self.log(f"Current page: {page_title} | URL: {current_url}", Colors.INFO, "ℹ️ ")
                except Exception:
                    pass
                return False
            
            # Scroll and interact
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", email_input)
            self.human_like_delay(0.2, 0.3)
            
            # Try multiple click methods
            try:
                self.move_mouse_randomly()
                ActionChains(self.driver).move_to_element(email_input).click().perform()
            except Exception:
                try:
                    email_input.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", email_input)
            
            self.human_like_delay(0.1, 0.2)
            
            # Clear and type
            try:
                email_input.clear()
            except Exception:
                self.driver.execute_script("arguments[0].value = '';", email_input)
            
            self.human_like_typing(email_input, email)
            self.log("Entered email", Colors.SUCCESS, "✅ ")
            
            self.human_like_delay(0.1, 0.2)
            return True
            
        except Exception as e:
            self.log(f"Error entering email: {e}", Colors.ERROR, "❌ ")
            return False
    
    def click_continue_button(self):
        """Click Continue button with retry based on network mode"""
        try:
            max_retries = self.net["max_retries"]
            
            selectors = [
                "//button[contains(text(), 'Tiếp tục')]",
                "//button[contains(text(), 'Continue')]",
                "//button[@type='submit']",
                "//button[contains(@class, 'btn-primary')]",
                "//button[contains(@class, 'continue')]",
                "//input[@type='submit']",
            ]
            
            continue_button = None
            
            for attempt in range(max_retries):
                timeout = self.net["element_timeout"] + (attempt * 3)
                wait = WebDriverWait(self.driver, timeout)
                
                if attempt > 0:
                    self.log(f"Retry finding Continue button ({attempt + 1}/{max_retries})...", Colors.WARNING, "🔄 ")
                    time.sleep(self.net["retry_delay"])
                
                for selector in selectors:
                    try:
                        continue_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        if continue_button.is_displayed():
                            break
                        else:
                            continue_button = None
                    except TimeoutException:
                        continue
                    except Exception:
                        continue
                
                if continue_button:
                    break
            
            if not continue_button:
                self.log("Could not find Continue button", Colors.ERROR, "❌ ")
                return False
            
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", continue_button)
            self.human_like_delay(0.1, 0.2)
            
            # Try multiple click methods
            click_success = False
            try:
                self.move_mouse_randomly()
                ActionChains(self.driver).move_to_element(continue_button).pause(0.2).click().perform()
                click_success = True
            except Exception:
                pass
            
            if not click_success:
                try:
                    continue_button.click()
                    click_success = True
                except Exception:
                    pass
            
            if not click_success:
                try:
                    self.driver.execute_script("arguments[0].click();", continue_button)
                    click_success = True
                except Exception:
                    pass
            
            if click_success:
                self.log("Clicked Continue", Colors.SUCCESS, "✅ ")
            else:
                self.log("Failed to click Continue button", Colors.ERROR, "❌ ")
                return False
            
            return True
            
        except Exception as e:
            self.log(f"Error clicking Continue: {e}", Colors.ERROR, "❌ ")
            return False
    
    def enter_password(self):
        """Enter password with retry based on network mode"""
        try:
            max_retries = self.net["max_retries"]
            
            selectors = [
                (By.ID, "password"),
                (By.XPATH, "//input[@id='password']"),
                (By.XPATH, "//input[@type='password']"),
                (By.XPATH, "//input[@name='password']"),
                (By.XPATH, "//input[@aria-label='Password']"),
                (By.XPATH, "//input[@aria-label='Mật khẩu']"),
                (By.XPATH, "//input[contains(@class, 'password')]"),
            ]
            
            password_input = None
            
            for attempt in range(max_retries):
                timeout = self.net["element_timeout"] + (attempt * 3)
                wait = WebDriverWait(self.driver, timeout)
                
                if attempt > 0:
                    self.log(f"Retry finding password input ({attempt + 1}/{max_retries})...", Colors.WARNING, "🔄 ")
                    time.sleep(self.net["retry_delay"])
                    
                    # Check page state
                    try:
                        page_state = self.driver.execute_script("return document.readyState")
                        if page_state != "complete":
                            self.log("Waiting for page to load...", Colors.INFO, "⏳ ")
                            WebDriverWait(self.driver, self.net["page_load_timeout"]).until(
                                lambda d: d.execute_script("return document.readyState") == "complete"
                            )
                    except Exception:
                        pass
                
                for by, selector in selectors:
                    try:
                        password_input = wait.until(EC.presence_of_element_located((by, selector)))
                        if password_input.is_displayed():
                            break
                        else:
                            password_input = None
                    except TimeoutException:
                        continue
                    except Exception:
                        continue
                
                if password_input:
                    break
            
            if not password_input:
                self.log(f"Could not find password input after {max_retries} attempts", Colors.ERROR, "❌ ")
                try:
                    current_url = self.driver.current_url
                    self.log(f"Current URL: {current_url}", Colors.INFO, "ℹ️ ")
                except Exception:
                    pass
                return False
            
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", password_input)
            self.human_like_delay(0.2, 0.3)
            
            # Try multiple click methods
            try:
                self.move_mouse_randomly()
                ActionChains(self.driver).move_to_element(password_input).click().perform()
            except Exception:
                try:
                    password_input.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", password_input)
            
            self.human_like_delay(0.1, 0.2)
            
            # Clear and type
            try:
                password_input.clear()
            except Exception:
                self.driver.execute_script("arguments[0].value = '';", password_input)
            
            self.human_like_typing(password_input, self.password)
            self.log("Entered password", Colors.SUCCESS, "✅ ")
            
            self.human_like_delay(0.1, 0.2)
            return True
            
        except Exception as e:
            self.log(f"Error entering password: {e}", Colors.ERROR, "❌ ")
            return False
    
    def extract_otp_from_email(self, email_text):
        """Extract 6-digit OTP"""
        patterns = [
            r'\b(\d{6})\b',
            r'code[:\s]+(\d{6})',
            r'OTP[:\s]+(\d{6})',
            r'verification[:\s]+(\d{6})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, email_text, re.IGNORECASE)
            if match:
                return match.group(1)
        return None
    
    def resend_otp_email(self):
        """Click Resend email button"""
        try:
            wait = WebDriverWait(self.driver, 1)
            
            selectors = [
                "//button[contains(text(), 'Gửi lại email')]",
                "//button[contains(text(), 'Resend email')]",
                "//button[@name='intent' and @value='resend-code']",
            ]
            
            resend_button = None
            for selector in selectors:
                try:
                    resend_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    break
                except:
                    continue
            
            if resend_button:
                self.move_mouse_randomly()
                ActionChains(self.driver).move_to_element(resend_button).pause(0.5).click().perform()
                self.log("Clicked Resend email", Colors.SUCCESS, "✅ ")
                return True
            else:
                return False
            
        except Exception as e:
            return False
    
    def wait_and_get_otp(self, max_attempts=30, wait_seconds=5, max_resend=2):
        """Wait and get OTP from email"""
        self.log("Waiting for OTP...", Colors.INFO, "📧 ")
        
        resend_count = 0
        
        for attempt in range(max_attempts):
            # Check for timeout error during OTP wait
            if self.has_operation_timeout_error():
                self.log("Timeout error detected while waiting for OTP!", Colors.WARNING, "⚠️ ")
                self.operation_timeout_detected = True
                return None
            
            # Different logic for OAuth2 vs TinyHost mode
            if self.email_mode == "OAuth2":
                # OAuth2 mode: use DongVanOAuth2API
                messages_data = self.mail_api.fetch_messages()
                if messages_data:
                    otp = self.mail_api.extract_code_from_messages(messages_data)
                    if otp:
                        self.log(f"Got OTP: {otp}", Colors.SUCCESS, "✅ ")
                        return otp
            else:
                # TinyHost mode: use TempMailAPI
                username = self.email_info['username']
                domain = self.email_info['domain']
                
                emails_data = self.mail_api.get_emails(domain, username, page=1, limit=5)
                
                if emails_data and emails_data.get('emails'):
                    for email_item in emails_data['emails']:
                        sender = email_item.get('sender', '').lower()
                        subject = email_item.get('subject', '').lower()
                        
                        if 'openai' in sender or 'verification' in subject or 'code' in subject:
                            email_id = email_item['id']
                            
                            email_detail = self.mail_api.get_email_detail(domain, username, email_id)
                            
                            if email_detail:
                                body = email_detail.get('body', '')
                                html_body = email_detail.get('html_body', '')
                                full_content = f"{body} {html_body}"
                                otp = self.extract_otp_from_email(full_content)
                                
                                if otp:
                                    self.log(f"Got OTP: {otp}", Colors.SUCCESS, "✅ ")
                                    return otp
            
            # Resend if needed (only for TinyHost mode)
            if self.email_mode != "OAuth2" and attempt == max_attempts // 3 and resend_count < max_resend:
                if self.resend_otp_email():
                    resend_count += 1
                    self.log(f"Resent email ({resend_count}/{max_resend})", Colors.INFO, "📧 ")
            
            # Check stop event and timeout during wait
            for _ in range(wait_seconds * 2):
                if self.stop_event and self.stop_event.is_set():
                    self.log("Process stopped by user", Colors.WARNING, "🛑 ")
                    return None
                # Check for timeout error periodically
                if self.has_operation_timeout_error():
                    self.log("Timeout error detected during OTP wait!", Colors.WARNING, "⚠️ ")
                    self.operation_timeout_detected = True
                    return None
                time.sleep(0.5)
        
        self.log("Timeout - Did not receive OTP", Colors.ERROR, "❌ ")
        return None
    
    def enter_otp_code(self, otp):
        """Enter OTP"""
        try:
            wait = WebDriverWait(self.driver, 1)
            time.sleep(0.5)
            
            selectors = [
                (By.ID, "code"),
                (By.XPATH, "//input[@id='code']"),
                (By.XPATH, "//input[@name='code']"),
                (By.XPATH, "//input[@inputmode='numeric']"),
            ]
            
            code_input = None
            for by, selector in selectors:
                try:
                    code_input = wait.until(EC.presence_of_element_located((by, selector)))
                    break
                except:
                    continue
            
            if not code_input:
                self.log("Could not find code input", Colors.ERROR, "❌ ")
                return False
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", code_input)
            self.human_like_delay(0.2, 0.4)
            
            self.move_mouse_randomly()
            ActionChains(self.driver).move_to_element(code_input).click().perform()
            self.human_like_delay(0.2, 0.4)
            
            code_input.clear()
            self.human_like_typing(code_input, otp)
            self.log("Entered OTP", Colors.SUCCESS, "✅ ")
            
            self.human_like_delay(0.2, 0.4)
            return True
            
        except Exception as e:
            self.log(f"Error entering OTP: {e}", Colors.ERROR, "❌ ")
            return False
    
    def enter_name_and_dob(self):
        """Enter name and DOB"""
        try:
            wait = WebDriverWait(self.driver, 1)
            
            name_selectors = [
                (By.ID, "name"),
                (By.XPATH, "//input[@id='name']"),
                (By.XPATH, "//input[@name='name']"),
                (By.XPATH, "//input[@autocomplete='name']"),
            ]
            
            name_input = None
            for by, selector in name_selectors:
                try:
                    name_input = wait.until(EC.presence_of_element_located((by, selector)))
                    break
                except:
                    continue
            
            if not name_input:
                self.log("Could not find name input", Colors.ERROR, "❌ ")
                return False
            
            self.driver.execute_script("arguments[0].scrollIntoView(true);", name_input)
            self.human_like_delay(0.2, 0.4)
            
            self.move_mouse_randomly()
            ActionChains(self.driver).move_to_element(name_input).click().perform()
            self.human_like_delay(0.2, 0.4)
            
            name_input.clear()
            self.human_like_typing(name_input, "GPT")
            self.log("Entered name: GPT", Colors.SUCCESS, "✅ ")
            
            self.human_like_delay(0.2, 0.4)
            
            # Press Tab to move to DOB
            name_input.send_keys(Keys.TAB)
            self.human_like_delay(0.3, 0.6)
            
            # Random DOB
            year = random.randint(1997, 2006)
            month = random.randint(1, 12)
            day = random.randint(1, 28)
            # Format MM/DD/YYYY (cả tháng và ngày đều có số 0 đầu)
            dob = f"{month:02d}/{day:02d}/{year}"
            
            # Type DOB
            actions = ActionChains(self.driver)
            for char in dob:
                actions.send_keys(char)
                actions.pause(random.uniform(0.1, 0.3))
            actions.perform()
            
            self.log(f"Entered DOB: {dob}", Colors.SUCCESS, "✅ ")
            
            # Check and click "I agree to all of the following" checkbox if present
            try:
                checkbox_selectors = [
                    (By.XPATH, "//input[@name='allCheckboxes']"),
                    (By.XPATH, "//input[contains(@id, 'allCheckboxes')]"),
                    (By.XPATH, "//label[contains(., 'I agree to all')]//input[@type='checkbox']"),
                ]
                
                checkbox = None
                for by, selector in checkbox_selectors:
                    try:
                        checkbox = WebDriverWait(self.driver, 1).until(
                            EC.presence_of_element_located((by, selector))
                        )
                        break
                    except:
                        continue
                
                if checkbox and not checkbox.is_selected():
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                    
                    # Use JavaScript click for checkbox (more reliable)
                    self.driver.execute_script("arguments[0].click();", checkbox)
                    self.log("Clicked 'I agree to all of the following' checkbox", Colors.SUCCESS, "✅ ")
            except Exception as checkbox_error:
                self.log(f"No 'I agree to all' checkbox found (optional): {checkbox_error}", Colors.INFO, "ℹ️ ")
            
            return True
            
        except Exception as e:
            self.log(f"Error entering name/DOB: {e}", Colors.ERROR, "❌ ")
            return False
    

    
    def get_cookies_json(self):
        """Get cookies in JSON format"""
        try:
            self.log("Getting cookies...", Colors.INFO, "🍪 ")
            time.sleep(4)
            
            all_cookies = self.driver.get_cookies()
            cookies_for_export = []
            
            for cookie in all_cookies:
                same_site = cookie.get("sameSite", "no_restriction")
                if same_site:
                    same_site = same_site.lower()
                    if same_site == "none":
                        same_site = "no_restriction"
                    elif same_site not in ["lax", "strict", "no_restriction", "unspecified"]:
                        same_site = "no_restriction"
                else:
                    same_site = "no_restriction"
                
                domain = cookie.get("domain", "")
                host_only = not domain.startswith(".")
                
                cookie_item = {
                    "domain": domain,
                    "expirationDate": cookie.get("expiry"),
                    "hostOnly": host_only,
                    "httpOnly": cookie.get("httpOnly", False),
                    "name": cookie.get("name", ""),
                    "path": cookie.get("path", "/"),
                    "sameSite": same_site,
                    "secure": cookie.get("secure", False),
                    "session": False if cookie.get("expiry") else True,
                    "storeId": None,
                    "value": cookie.get("value", "")
                }
                cookies_for_export.append(cookie_item)
            
            has_session = any(c["name"] == "__Secure-next-auth.session-token" for c in cookies_for_export)
            if has_session:
                self.log("Found session token", Colors.SUCCESS, "✅ ")
            else:
                self.log("No session token found", Colors.WARNING, "⚠️ ")
            
            return cookies_for_export
            
        except Exception as e:
            self.log(f"Error getting cookies: {e}", Colors.ERROR, "❌ ")
            return []
    
    def ensure_personal_tab_active(self):
        """Ensure Personal tab is active, not Business tab (for Plus checkout)"""
        try:
            # Check if Personal tab is active (data-state="on" means active)
            personal_active_xpath = "//button[@role='radio' and contains(., 'Personal') and @data-state='on']"
            personal_inactive_xpath = "//button[@role='radio' and contains(., 'Personal') and @data-state='off']"
            
            # Check if Personal is already active
            active_tabs = self.driver.find_elements(By.XPATH, personal_active_xpath)
            if active_tabs:
                return True  # Already on Personal tab
            
            # Personal is not active, try to click it
            inactive_tabs = self.driver.find_elements(By.XPATH, personal_inactive_xpath)
            if inactive_tabs:
                self.log("Switching to Personal tab...", Colors.INFO, "🔄 ")
                inactive_tabs[0].click()
                time.sleep(1)
                return True
            
            return False
        except Exception:
            return False
    
    def get_checkout_link(self):
        """Navigate to pricing section and capture checkout URL"""
        try:
            if not self.driver:
                return None
            self.log("Navigating to pricing to capture checkout link...", Colors.INFO, "💳 ")
            time.sleep(1.5)  # Wait after getting session cookie
            try:
                self.driver.execute_script("location.hash = '#pricing';")
            except Exception:
                pass
            time.sleep(1.5)
            wait = WebDriverWait(self.driver, 2)
            
            # Ensure Personal tab is active (not Business)
            self.ensure_personal_tab_active()
            
            button_xpath = (
                "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'claim free offer')]"
                " | //a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'claim free offer')]"
                " | //button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'get plus')]"
                " | //a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'get plus')]"
                " | //button[contains(normalize-space(.), 'Dùng bản Plus')]"
                " | //a[contains(normalize-space(.), 'Dùng bản Plus')]"
            )

            def locate_plus_button(reload_if_needed):
                attempts = 2 if reload_if_needed else 1
                for attempt in range(attempts):
                    try:
                        return wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                    except TimeoutException:
                        if reload_if_needed and attempt == 0:
                            self.log("Get Plus button not found, ensuring Personal tab...", Colors.WARNING, "ℹ️ ")
                            self.driver.get("https://chatgpt.com/#pricing")
                            time.sleep(2)
                            # Ensure Personal tab is active
                            self.ensure_personal_tab_active()
                        else:
                            break
                return None

            def check_payment_error():
                """Check if payment error message is displayed using proper element"""
                try:
                    # Check for error alert element with orange background
                    error_selectors = [
                        "//div[contains(@class, 'bg-orange-500') and @role='alert']//div[contains(@class, 'text-start')]",
                        "//div[@role='alert' and contains(@class, 'border-orange-500')]//div[contains(@class, 'whitespace-pre-wrap')]",
                        "//div[contains(@class, 'bg-orange-500')]//div[contains(., 'payments page encountered an error')]",
                    ]
                    for selector in error_selectors:
                        try:
                            error_elements = self.driver.find_elements(By.XPATH, selector)
                            for el in error_elements:
                                if "payments page encountered an error" in el.text.lower():
                                    return True
                        except:
                            continue
                    return False
                except:
                    return False

            def poll_for_checkout_or_error(timeout=30, poll_interval=0.5):
                """Poll continuously for checkout URL or payment error"""
                start_time = time.time()
                while time.time() - start_time < timeout:
                    try:
                        # Check for payment error first
                        if check_payment_error():
                            return "PAYMENT_ERROR"
                        
                        # Check for checkout URL
                        current_url = self.driver.current_url or ""
                        if "/checkout/" in current_url.lower() or "pay.openai.com" in current_url.lower():
                            # Check for invalid verify URL (e.g. /checkout/verify?stripe_session_id=)
                            if "verify?stripe_session_id" in current_url.lower() or "/checkout/verify?" in current_url.lower():
                                self.log("Invalid verify URL detected, need to retry...", Colors.WARNING, "⚠️ ")
                                return "VERIFY_ERROR"
                            return current_url
                    except:
                        pass
                    
                    time.sleep(poll_interval)
                
                return None  # Timeout

            def click_plus_and_wait(plus_button, attempt=1):
                handles_before = list(self.driver.window_handles)
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", plus_button)
                time.sleep(0.5)
                plus_button.click()
                self.log(f"Clicked Get Plus (attempt {attempt})", Colors.SUCCESS, "✅ ")
                time.sleep(1)
                handles_after = self.driver.window_handles
                new_handles = [handle for handle in handles_after if handle not in handles_before]
                if new_handles:
                    self.driver.switch_to.window(new_handles[-1])
                    self.log("Switched to checkout tab", Colors.INFO, "🪟 ")
                
                # Use polling to check for checkout URL or error
                self.log("Waiting for checkout page...", Colors.INFO, "⏳ ")
                result = poll_for_checkout_or_error(timeout=30, poll_interval=0.5)
                
                if result == "PAYMENT_ERROR":
                    self.log("Payment page error detected!", Colors.WARNING, "⚠️ ")
                    return "PAYMENT_ERROR"
                elif result == "VERIFY_ERROR":
                    self.log("Invalid verify URL detected!", Colors.WARNING, "⚠️ ")
                    return "VERIFY_ERROR"
                elif result:
                    return result
                else:
                    # Timeout - switch back to main window
                    if self.driver.window_handles:
                        try:
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        except Exception:
                            pass
                    return None

            # Retry loop for Plus checkout (max 2 attempts for payment errors)
            checkout_url = None
            for attempt in range(1, 3):
                button = locate_plus_button(reload_if_needed=True)
                if not button:
                    self.log("Get Plus button not found", Colors.WARNING, "ℹ️ ")
                    return None
                
                checkout_url = click_plus_and_wait(button, attempt)
                
                if checkout_url == "PAYMENT_ERROR" or checkout_url == "VERIFY_ERROR":
                    error_type = "payment error" if checkout_url == "PAYMENT_ERROR" else "verify URL error"
                    if attempt < 2:
                        self.log(f"Retrying due to {error_type}...", Colors.WARNING, "🔄 ")
                        # Close error tab and go back
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        # Navigate back to pricing and ensure Personal tab
                        self.log("Refreshing page to clear error state...", Colors.INFO, "🔄 ")
                        self.driver.refresh()
                        time.sleep(3)
                        
                        self.log("Navigating back to pricing...", Colors.INFO, "🔄 ")
                        self.driver.get("https://chatgpt.com/#pricing")
                        time.sleep(2)
                        self.ensure_personal_tab_active()
                        time.sleep(1)
                        continue
                    else:
                        self.log(f"{error_type.capitalize()} persists, cannot capture Plus link", Colors.ERROR, "❌ ")
                        return None
                elif checkout_url:
                    break
                else:
                    if attempt < 2:
                        self.log("Checkout page did not load, retrying...", Colors.WARNING, "🔄 ")
                        self.driver.get("https://chatgpt.com/#pricing")
                        time.sleep(2)
                        self.ensure_personal_tab_active()
                    else:
                        self.log("Failed to open checkout after retries", Colors.WARNING, "ℹ️ ")
                        return None
            if checkout_url:
                self.log(f"Plus Checkout URL: {checkout_url}", Colors.SUCCESS, "✅ ")
            return checkout_url
        except Exception as e:
            self.log(f"Could not capture checkout link: {e}", Colors.WARNING, "ℹ️ ")
            return None
    
    def get_business_checkout_link(self):
        """Navigate to Business pricing and capture checkout URL"""
        try:
            if not self.driver:
                return None
            self.log("Navigating to Business pricing...", Colors.INFO, "💼 ")
            
            # Navigate to Business team pricing page
            business_url = "https://chatgpt.com/?numSeats=5&selectedPlan=month&referrer=#team-pricing-seat-selection"
            self.driver.get(business_url)
            time.sleep(3)
            
            wait = WebDriverWait(self.driver, 7)
            
            # Find and click "Continue to billing" button
            continue_billing_xpath = (
                "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'continue to billing')]"
                " | //button[contains(@class, 'btn-green') and contains(., 'Continue to billing')]"
                " | //button[contains(@class, 'btn-green') and contains(., 'billing')]"
                " | //button[contains(., 'Continue to billing')]"
            )
            
            def locate_continue_button():
                try:
                    return wait.until(EC.element_to_be_clickable((By.XPATH, continue_billing_xpath)))
                except TimeoutException:
                    self.log("Continue to billing button not found, retrying...", Colors.WARNING, "⚠️ ")
                    # Retry with page reload
                    self.driver.get(business_url)
                    time.sleep(3)
                    try:
                        return wait.until(EC.element_to_be_clickable((By.XPATH, continue_billing_xpath)))
                    except TimeoutException:
                        return None
            
            def check_payment_error():
                """Check if payment error message is displayed using proper element"""
                try:
                    # Check for error alert element with orange background
                    error_selectors = [
                        "//div[contains(@class, 'bg-orange-500') and @role='alert']//div[contains(@class, 'text-start')]",
                        "//div[@role='alert' and contains(@class, 'border-orange-500')]//div[contains(@class, 'whitespace-pre-wrap')]",
                        "//div[contains(@class, 'bg-orange-500')]//div[contains(., 'payments page encountered an error')]",
                    ]
                    for selector in error_selectors:
                        try:
                            error_elements = self.driver.find_elements(By.XPATH, selector)
                            for el in error_elements:
                                if "payments page encountered an error" in el.text.lower():
                                    return True
                        except:
                            continue
                    return False
                except:
                    return False

            def poll_for_checkout_or_error(timeout=30, poll_interval=0.5):
                """Poll continuously for checkout URL or payment error"""
                start_time = time.time()
                while time.time() - start_time < timeout:
                    try:
                        # Check for payment error first
                        if check_payment_error():
                            return "PAYMENT_ERROR"
                        
                        # Check for checkout URL
                        current_url = self.driver.current_url or ""
                        if "/checkout/" in current_url.lower() or "pay.openai.com" in current_url.lower():
                            # Check for invalid verify URL (e.g. /checkout/verify?stripe_session_id=)
                            if "verify?stripe_session_id" in current_url.lower() or "/checkout/verify?" in current_url.lower():
                                self.log("Invalid verify URL detected, need to retry...", Colors.WARNING, "⚠️ ")
                                return "VERIFY_ERROR"
                            return current_url
                    except:
                        pass
                    
                    time.sleep(poll_interval)
                
                return None  # Timeout

            def click_continue_and_wait(button, attempt=1):
                handles_before = list(self.driver.window_handles)
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
                time.sleep(0.5)
                
                # Try multiple click methods
                try:
                    button.click()
                except Exception:
                    try:
                        self.driver.execute_script("arguments[0].click();", button)
                    except Exception:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(self.driver).move_to_element(button).click().perform()
                
                self.log(f"Clicked Continue to billing (attempt {attempt})", Colors.SUCCESS, "✅ ")
                time.sleep(1)
                
                # Check for new tab
                handles_after = self.driver.window_handles
                new_handles = [h for h in handles_after if h not in handles_before]
                if new_handles:
                    self.driver.switch_to.window(new_handles[-1])
                    self.log("Switched to checkout tab", Colors.INFO, "🪟 ")
                
                # Use polling to check for checkout URL or error
                self.log("Waiting for checkout page...", Colors.INFO, "⏳ ")
                result = poll_for_checkout_or_error(timeout=30, poll_interval=0.5)
                
                if result == "PAYMENT_ERROR":
                    self.log("Payment page error detected!", Colors.WARNING, "⚠️ ")
                    return "PAYMENT_ERROR"
                elif result == "VERIFY_ERROR":
                    self.log("Invalid verify URL detected!", Colors.WARNING, "⚠️ ")
                    return "VERIFY_ERROR"
                elif result:
                    return result
                else:
                    self.log(f"Timeout waiting for checkout (attempt {attempt})", Colors.WARNING, "⏱️ ")
                    # Switch back to main window if checkout didn't load
                    if self.driver.window_handles:
                        try:
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        except Exception:
                            pass
                    return None
            
            # Retry loop for clicking Continue to billing (max 3 attempts)
            checkout_url = None
            max_click_attempts = 3
            
            for attempt in range(1, max_click_attempts + 1):
                button = locate_continue_button()
                if not button:
                    if attempt < max_click_attempts:
                        self.log(f"Button not found, retrying... ({attempt}/{max_click_attempts})", Colors.WARNING, "🔄 ")
                        self.driver.get(business_url)
                        time.sleep(3)
                        continue
                    else:
                        self.log("Continue to billing button not found after all attempts", Colors.WARNING, "ℹ️ ")
                        return None
                
                checkout_url = click_continue_and_wait(button, attempt)
                
                if checkout_url == "PAYMENT_ERROR" or checkout_url == "VERIFY_ERROR":
                    error_type = "payment error" if checkout_url == "PAYMENT_ERROR" else "verify URL error"
                    if attempt < max_click_attempts:
                        self.log(f"Retrying due to {error_type}... ({attempt + 1}/{max_click_attempts})", Colors.WARNING, "🔄 ")
                        # Close error tab and go back
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        self.driver.get(business_url)
                        time.sleep(3)
                        continue
                    else:
                        self.log(f"{error_type.capitalize()} persists, cannot capture Business link", Colors.ERROR, "❌ ")
                        return None
                elif checkout_url:
                    break  # Success, exit loop
                elif attempt < max_click_attempts:
                    self.log(f"Retrying click... ({attempt + 1}/{max_click_attempts})", Colors.WARNING, "🔄 ")
                    # Reload page for next attempt
                    try:
                        self.driver.get(business_url)
                        time.sleep(3)
                    except Exception:
                        pass
            
            if checkout_url:
                self.log(f"Business Checkout URL: {checkout_url}", Colors.SUCCESS, "✅ ")
            return checkout_url
            
        except Exception as e:
            self.log(f"Could not capture Business checkout link: {e}", Colors.WARNING, "ℹ️ ")
            return None
    
    def save_account_info(self, cookies, checkout_url=None, business_checkout_url=None):
        """Save account info to Excel (thread-safe) - only if valid session cookie"""
        try:
            if not self.email_info:
                return False
            
            if not cookies:
                self.log("No cookies to save", Colors.WARNING, "⚠️ ")
                return False
            
            # Check if cookies start with valid .chatgpt.com domain
            cookies_json = json.dumps(cookies, separators=(',', ':'), ensure_ascii=False)
            
            if cookies_json.startswith('[{"domain":".chatgpt.com"'):
                self.log("Valid session cookie (.chatgpt.com)", Colors.SUCCESS, "✅ ")
            elif cookies_json.startswith('[{"domain":".auth.openai'):
                self.log("Invalid session cookie (.auth.openai) - NOT saving", Colors.ERROR, "❌ ")
                return False
            else:
                # Check if any cookie has .chatgpt.com domain
                has_chatgpt_cookie = any(c.get('domain') == '.chatgpt.com' for c in cookies)
                if not has_chatgpt_cookie:
                    self.log("No .chatgpt.com cookie found - NOT saving", Colors.ERROR, "❌ ")
                    return False
            
            filename = "chatgpt.xlsx"
            account = f"{self.email_info['email']}:{self.password}"
            
            with file_lock:
                if os.path.exists(filename):
                    # Load existing workbook
                    wb = load_workbook(filename)
                    ws = wb.active
                else:
                    # Create new workbook with headers
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Accounts"
                    ws['A1'] = "Account"
                    ws['B1'] = "Cookie"
                    ws['C1'] = "Plus Checkout URL"
                    ws['D1'] = "Business Checkout URL"
                    ws['E1'] = "2FA Secret"
                
                # Ensure headers exist for older files
                if ws['D1'].value is None:
                    ws['D1'] = "Business Checkout URL"
                if ws['E1'].value is None:
                    ws['E1'] = "2FA Secret"
                
                # Find next row
                next_row = ws.max_row + 1
                
                # Write data
                ws[f'A{next_row}'] = account
                ws[f'B{next_row}'] = cookies_json
                ws[f'C{next_row}'] = checkout_url if checkout_url else ""
                ws[f'D{next_row}'] = business_checkout_url if business_checkout_url else ""
                
                # Save workbook
                wb.save(filename)
                wb.close()
            
            self.log("Info saved to chatgpt.xlsx", Colors.SUCCESS, "✅ ")
            if checkout_url:
                self.log("Plus Checkout URL saved!", Colors.SUCCESS, "✅ ")
            if business_checkout_url:
                self.log("Business Checkout URL saved!", Colors.SUCCESS, "💼 ")
            return True
            
        except Exception as e:
            self.log(f"Error saving info: {e}", Colors.ERROR, "❌ ")
            return False
    
    def run(self):
        """Run the entire registration flow with timeout retry support"""
        # Retry loop for operation timeout errors
        while self.current_retry <= self.max_timeout_retries:
            try:
                if self.current_retry > 0:
                    self.log(f"Retry attempt {self.current_retry}/{self.max_timeout_retries} after timeout...", Colors.WARNING, "🔄 ")
                    time.sleep(3)
                
                self.operation_timeout_detected = False
                
                self.log("=" * 60, Colors.HEADER)
                self.log("STARTING CHATGPT REGISTRATION", Colors.HEADER, "🚀 ")
                self.log("=" * 60, Colors.HEADER)
                
                if not self.setup_driver(max_retries=3):
                    return (False, None)
                
                if not self.navigate_to_chatgpt():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.click_login_button():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.enter_email_address():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.click_continue_button():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.enter_password():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.click_continue_button():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                # Check for timeout after password submit (common failure point)
                time.sleep(1)
                if self.check_and_handle_timeout():
                    self.log("Timeout detected after password. Restarting...", Colors.WARNING, "🔁 ")
                    self.cleanup_browser()
                    self.current_retry += 1
                    continue
                
                otp = self.wait_and_get_otp(max_attempts=30, wait_seconds=3)
                if not otp:
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.enter_otp_code(otp):
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.click_continue_button():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                # Check for timeout after OTP submit
                time.sleep(1)
                if self.check_and_handle_timeout():
                    self.log("Timeout detected after OTP. Restarting...", Colors.WARNING, "🔁 ")
                    self.cleanup_browser()
                    self.current_retry += 1
                    continue
                
                if not self.enter_name_and_dob():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                if not self.click_continue_button():
                    if self.check_and_handle_timeout():
                        self.cleanup_browser()
                        self.current_retry += 1
                        continue
                    self.cleanup_browser()
                    return (False, None)
                
                # Final timeout check
                time.sleep(1)
                if self.check_and_handle_timeout():
                    self.log("Timeout detected after DOB. Restarting...", Colors.WARNING, "🔁 ")
                    self.cleanup_browser()
                    self.current_retry += 1
                    continue
                
                # Get cookies directly after DOB (skip Skip and final Continue buttons)
                self.log("Registration complete! Getting cookies...", Colors.SUCCESS, "✅ ")
                
                cookies = self.get_cookies_json()
                
                # Get checkout link if mode is enabled and cookies are valid
                checkout_url = None
                business_checkout_url = None
                if GET_CHECKOUT_LINK:
                    cookies_json = json.dumps(cookies, separators=(',', ':'), ensure_ascii=False)
                    if cookies_json.startswith('[{"domain":".chatgpt.com"') or any(c.get('domain') == '.chatgpt.com' for c in cookies):
                        checkout_type = GET_CHECKOUT_TYPE
                        
                        if checkout_type == "Plus":
                            checkout_url = self.get_checkout_link()
                        elif checkout_type == "Business":
                            business_checkout_url = self.get_business_checkout_link()
                        elif checkout_type == "Both":
                            checkout_url = self.get_checkout_link()
                            # Close any new tabs and go back for business link
                            try:
                                while len(self.driver.window_handles) > 1:
                                    self.driver.switch_to.window(self.driver.window_handles[-1])
                                    self.driver.close()
                                self.driver.switch_to.window(self.driver.window_handles[0])
                                time.sleep(1)
                            except Exception:
                                pass
                            business_checkout_url = self.get_business_checkout_link()
                    else:
                        self.log("Skipping checkout link - invalid session cookie", Colors.WARNING, "⚠️ ")
                
                saved = self.save_account_info(cookies, checkout_url, business_checkout_url)
                
                if not saved:
                    self.log("Account not saved due to invalid cookies", Colors.ERROR, "❌ ")
                    self.cleanup_browser()
                    return (False, None)
                
                self.log("=" * 60, Colors.SUCCESS)
                self.log("COMPLETED!", Colors.SUCCESS, "🎉 ")
                self.log(f"Account: {self.email_info['email']}:{self.password}", Colors.INFO)
                self.log("=" * 60, Colors.SUCCESS)
                
                time.sleep(3)
                
                result = {
                    'email': self.email_info['email'],
                    'password': self.password
                }
                
                # Cleanup browser after success
                self.cleanup_browser()
                
                return (True, result)
                
            except Exception as e:
                self.log(f"Error: {e}", Colors.ERROR, "❌ ")
                traceback.print_exc()
                
                # Check if it's a timeout error that warrants retry
                if self.check_and_handle_timeout() and self.current_retry < self.max_timeout_retries:
                    self.cleanup_browser()
                    self.current_retry += 1
                    continue
                
                self.cleanup_browser()
                return (False, None)
        
        # Exhausted all retries
        self.log(f"Failed after {self.max_timeout_retries} timeout retries", Colors.ERROR, "❌ ")
        self.cleanup_browser()
        return (False, None)



def run_worker(thread_id, stop_event=None, thread_delay=2, num_threads=1, email_mode="TinyHost", oauth2_account=None):
    """Worker function for registration with staggered start and retry logic"""
    # Apply delay based on position within current batch (reset for each batch)
    position_in_batch = (thread_id - 1) % num_threads if num_threads > 1 else 0
    delay = position_in_batch * thread_delay
    if delay > 0:
        safe_print(thread_id, f"Waiting {delay}s before starting...", Colors.INFO, "⏳ ")
        # Sleep in intervals to check stop_event
        for _ in range(int(delay * 2)):
            if stop_event and stop_event.is_set():
                return (False, None)
            time.sleep(0.5)
    
    worker = ChatGPTAutoRegisterWorker(
        thread_id, 
        num_threads=num_threads,
        email_mode=email_mode,
        oauth2_account=oauth2_account
    )
    worker.stop_event = stop_event
    return worker.run()



# ============================================================================
# MODULE 3: MFA AUTOMATION
# ============================================================================

class MFAWorker:
    """Worker for MFA enrollment"""
    
    def __init__(self, thread_id, row_index, email, password, cookie_json, excel_file, oauth2_account=None):
        self.thread_id = thread_id
        self.row_index = row_index
        self.email = email
        self.password = password
        self.cookie_json = cookie_json
        self.excel_file = excel_file
        self.oauth2_account = oauth2_account
        self.driver = None
        self.user_data_dir = None
        self.stop_event = None
        
        # Popup polling
        self._popup_polling_active = False
        self._popup_polling_thread = None
        self._popup_dismissed = False

        # Parse email info for TempMailAPI
        parts = self.email.split('@')
        if len(parts) == 2:
            self.email_info = {
                'username': parts[0],
                'domain': parts[1],
                'email': self.email
            }
        else:
            self.email_info = None
            
        if self.oauth2_account:
            self.mail_api = DongVanOAuth2API(
                email=self.oauth2_account.get("email", ""),
                password=self.oauth2_account.get("password", ""),
                refresh_token=self.oauth2_account.get("refresh_token", ""),
                client_id=self.oauth2_account.get("client_id", "")
            )
        else:
            self.mail_api = TempMailAPI()

        
    def log(self, message, color=Colors.INFO, emoji=""):
        safe_print(self.thread_id, message, color, emoji)
    
    def setup_driver(self):
        """Initialize Chrome driver"""
        try:
            self.log("Initializing ChromeDriver...", Colors.INFO, "🔄 ")
            
            options = uc.ChromeOptions()
            
            timestamp = int(time.time() * 1000)
            random_suffix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=6))
            self.user_data_dir = tempfile.mkdtemp(prefix=f"chrome_mfa_T{self.thread_id}_{timestamp}_{random_suffix}_")
            options.add_argument(f'--user-data-dir={self.user_data_dir}')
            
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--window-size=900,700')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--lang=en-US')
            
            prefs = {
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
                "profile.default_content_setting_values.notifications": 2
            }
            options.add_experimental_option("prefs", prefs)
            
            with driver_init_lock:
                try:
                    self.driver = uc.Chrome(options=options, version_main=CHROME_VERSION_MAIN)
                except Exception:
                    self.driver = uc.Chrome(options=options, version_main=None)
            
            window_width = 900
            window_height = 700
            x_offset = (self.thread_id - 1) * 80
            y_offset = (self.thread_id - 1) * 40
            
            self.driver.set_window_size(window_width, window_height)
            self.driver.set_window_position(x_offset, y_offset)
            
            self.log("ChromeDriver initialized!", Colors.SUCCESS, "✅ ")
            return True
            
        except Exception as e:
            self.log(f"Error initializing driver: {e}", Colors.ERROR, "❌ ")
            return False
    
    def import_cookies(self):
        """Import cookies into browser"""
        try:
            self.log("Importing cookies...", Colors.INFO, "🍪 ")
            
            self.driver.get("https://chatgpt.com/")
            time.sleep(2)
            
            cookies = json.loads(self.cookie_json) if isinstance(self.cookie_json, str) else self.cookie_json
            
            for cookie in cookies:
                try:
                    cookie_dict = {
                        "name": cookie.get("name"),
                        "value": cookie.get("value"),
                        "domain": cookie.get("domain", ".chatgpt.com"),
                        "path": cookie.get("path", "/"),
                        "secure": cookie.get("secure", True),
                        "httpOnly": cookie.get("httpOnly", False)
                    }
                    
                    if cookie.get("expirationDate"):
                        cookie_dict["expiry"] = int(cookie.get("expirationDate"))
                    
                    self.driver.add_cookie(cookie_dict)
                except Exception:
                    continue
            
            self.log("Cookies imported!", Colors.SUCCESS, "✅ ")
            self.driver.refresh()
            time.sleep(3)
            
            # Dismiss onboarding popup if appears
            self.start_popup_polling()

            
            return True
            
        except Exception as e:
            self.log(f"Error importing cookies: {e}", Colors.ERROR, "❌ ")
            return False
    
    def start_popup_polling(self):
        """Start background thread to continuously check and dismiss onboarding popup"""
        if self._popup_polling_active or self._popup_dismissed:
            return
        
        self._popup_polling_active = True
        self._popup_polling_thread = threading.Thread(target=self._popup_polling_loop, daemon=True)
        self._popup_polling_thread.start()
    
    def stop_popup_polling(self):
        """Stop the popup polling thread"""
        self._popup_polling_active = False
        if self._popup_polling_thread and self._popup_polling_thread.is_alive():
            self._popup_polling_thread.join(timeout=2)
        self._popup_polling_thread = None
    
    def _popup_polling_loop(self):
        """Background loop to check for onboarding popup every 2 seconds"""
        while self._popup_polling_active and not self._popup_dismissed:
            try:
                if not self.driver:
                    break
                
                button_selectors = [
                    "//button[.//div[contains(text(), \"Okay, let's go\")]]",
                    "//button[contains(., \"Okay, let's go\")]",
                    "//div[@role='dialog']//button[contains(., 'Okay')]",
                ]
                
                for selector in button_selectors:
                    try:
                        buttons = self.driver.find_elements(By.XPATH, selector)
                        for btn in buttons:
                            if btn.is_displayed():
                                self.driver.execute_script("arguments[0].click();", btn)
                                self.log("Dismissed onboarding popup", Colors.SUCCESS, "✅ ")
                                self._popup_dismissed = True
                                self._popup_polling_active = False
                                return
                    except:
                        continue
            except:
                pass
            
            # Poll every 2 seconds
            time.sleep(2)

    
    def navigate_to_security_settings(self):
        """Navigate to Settings > Security"""
        try:
            self.log("Navigating to Security settings...", Colors.INFO, "⚙️ ")
            self.driver.get("https://chatgpt.com/#settings/Security")
            time.sleep(3)
            return True
        except Exception as e:
            self.log(f"Error navigating to settings: {e}", Colors.ERROR, "❌ ")
            return False
    
    def click_mfa_toggle(self):
        """Click the MFA toggle button"""
        try:
            self.log("Clicking MFA toggle...", Colors.INFO, "🔐 ")
            time.sleep(2)
            
            toggle_selectors = [
                "//button[@data-state='unchecked' and contains(@class, 'radix')]",
                "//button[@data-state='unchecked']",
                "//div[contains(text(), 'Authenticator app')]/following::button[@data-state][1]",
                "//span[contains(text(), 'Authenticator app')]/ancestor::div//button[@data-state]",
                "//button[@role='switch']",
                "//button[contains(@class, 'rounded-full') and contains(@class, 'bg-')]",
            ]
            
            toggle = None
            for selector in toggle_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        state = elem.get_attribute("data-state")
                        if state == "checked":
                            continue
                        if elem.is_displayed():
                            toggle = elem
                            break
                    if toggle:
                        break
                except:
                    continue
            
            if toggle:
                self.driver.execute_script("arguments[0].click();", toggle)
                self.log("Clicked MFA toggle", Colors.SUCCESS, "✅ ")
                time.sleep(3)
                return True
            else:
                self.log("Could not find MFA toggle", Colors.ERROR, "❌ ")
                return False
                
        except Exception as e:
            self.log(f"Error clicking toggle: {e}", Colors.ERROR, "❌ ")
            return False
    
    def extract_otp_from_email(self, email_text):
        patterns = [r'\b(\d{6})\b', r'code[:\s]+(\d{6})', r'verification[:\s]+(\d{6})']
        for pattern in patterns:
            match = re.search(pattern, email_text, re.IGNORECASE)
            if match: return match.group(1)
        return None

    def wait_and_get_otp(self, max_attempts=30, wait_seconds=5):
        self.log("Waiting for OTP from email...", Colors.INFO, "📧 ")
        
        # Check if email is Outlook/Hotmail
        is_outlook = self.email.lower().endswith(('@outlook.com', '@hotmail.com'))
        
        # HANDLE OAUTH2 MODE (Using DongVanOAuth2API)
        if self.oauth2_account:
            refresh_token = self.oauth2_account.get("refresh_token")
            client_id = self.oauth2_account.get("client_id")
            
            if not refresh_token or not client_id:
                self.log("Missing refresh_token or client_id for OAuth2", Colors.ERROR, "❌ ")
                return None
                
            for _ in range(max_attempts):
                if self.stop_event and self.stop_event.is_set():
                    self.log("Process stopped by user", Colors.WARNING, "🛑 ")
                    return None
                
                try:
                    # DongVanOAuth2API uses instance vars for auth, no params needed
                    messages = self.mail_api.fetch_messages()
                    if messages:
                        code = self.mail_api.extract_code_from_messages(messages)
                        if code:
                            self.log(f"Got OTP (OAuth2): {code}", Colors.SUCCESS, "✅ ")
                            return code
                except Exception as e:
                    pass  # Ignore errors during polling
                
                # Sleep in intervals
                for _ in range(wait_seconds * 2):
                    if self.stop_event and self.stop_event.is_set():
                        return None
                    time.sleep(0.5)
            return None
            
        # HANDLE OUTLOOK/HOTMAIL WITHOUT OAUTH2 - SKIP
        elif is_outlook:
            self.log("Outlook/Hotmail email requires OAuth2 credentials in oauth2.xlsx", Colors.ERROR, "❌ ")
            return None
            
        # HANDLE TINYHOST MODE (Using TempMailAPI)
        else:
            if not self.email_info: return None
            user, domain = self.email_info['username'], self.email_info['domain']
            
            for _ in range(max_attempts):
                if self.stop_event and self.stop_event.is_set():
                    self.log("Process stopped by user", Colors.WARNING, "🛑 ")
                    return None
                    
                data = self.mail_api.get_emails(domain, user, limit=5)
                if data and data.get('emails'):
                    for item in data['emails']:
                        if 'openai' in item.get('sender', '').lower() or 'code' in item.get('subject', '').lower():
                            detail = self.mail_api.get_email_detail(domain, user, item['id'])
                            if detail:
                                full = f"{detail.get('body','')} {detail.get('html_body','')}"
                                otp = self.extract_otp_from_email(full)
                                if otp:
                                    self.log(f"Got OTP: {otp}", Colors.SUCCESS, "✅ ")
                                    return otp
                
                # Sleep in intervals to check stop_event
                for _ in range(wait_seconds * 2):
                    if self.stop_event and self.stop_event.is_set():
                        return None
                    time.sleep(0.5)
                    
            return None

    def enter_otp_code(self, otp):
        try:
            wait = WebDriverWait(self.driver, 3)
            inp = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Code' or @name='code']")))
            inp.clear()
            for c in otp:
                inp.send_keys(c)
                time.sleep(0.05)
            self.log("Entered email OTP", Colors.SUCCESS, "✅ ")
            
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]")))
            btn.click()
            self.log("Clicked Continue (OTP)", Colors.SUCCESS, "✅ ")
            time.sleep(3)
            return True
        except Exception as e:
            self.log(f"Error entering OTP: {e}", Colors.ERROR, "❌ ")
            return False

    def handle_password_verification(self):
        """Handle password OR email verification, or password -> OTP flow"""
        try:
            self.log("Waiting for verification...", Colors.INFO, "🔑 ")
            time.sleep(3)
            
            url = self.driver.current_url
            page_source = self.driver.page_source.lower()
            
            # Case 1: Email OTP first
            if "check your inbox" in page_source or "enter the verification code" in page_source or "email-verification" in url:
                self.log("Detected 'Check your inbox' screen", Colors.WARNING, "📧 ")
                if not self._handle_email_otp_verification():
                    return False

            # Case 2: Password page
            elif "auth.openai.com" in url or "log-in/password" in url or "enter your password" in page_source:
                self.log("On password verification page", Colors.INFO, "ℹ️ ")
                if not self._enter_password_and_click_continue():
                    return False
                
                # After entering password, check if OTP page appears
                time.sleep(3)
                new_url = self.driver.current_url
                new_page_source = self.driver.page_source.lower()
                if "check your inbox" in new_page_source or "enter the verification code" in new_page_source or "email-verification" in new_url:
                    self.log("OTP page appeared after password", Colors.INFO, "📧 ")
                    if not self._handle_email_otp_verification():
                        return False
            
            else:
                self.log("Not on known verification page, continuing...", Colors.INFO, "ℹ️ ")

            # Check redirect and return to settings if needed
            if "chatgpt.com" in self.driver.current_url:
                self.log("Redirecting back to Security settings...", Colors.INFO, "⚙️ ")
                self.driver.get("https://chatgpt.com/#settings/Security")
                time.sleep(3)
                try:
                    toggle = self.driver.find_element(By.XPATH, "//button[@data-state='unchecked']")
                    if toggle.is_displayed():
                        self.driver.execute_script("arguments[0].click();", toggle)
                        self.log("Clicked MFA toggle again", Colors.SUCCESS, "✅ ")
                        time.sleep(3)
                except: pass
            
            return True
                
        except Exception as e:
            self.log(f"Error with verification: {e}", Colors.ERROR, "❌ ")
            return False
    
    def _enter_password_and_click_continue(self):
        """Enter password and click Continue button"""
        try:
            wait = WebDriverWait(self.driver, 5)
            password_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
            password_input.clear()
            for char in self.password:
                password_input.send_keys(char)
                time.sleep(random.uniform(0.05, 0.1))
            self.log("Entered password", Colors.SUCCESS, "✅ ")
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Continue')]"))).click()
            self.log("Clicked Continue", Colors.SUCCESS, "✅ ")
            time.sleep(5)
            return True
        except Exception as e:
            self.log(f"Error entering password: {e}", Colors.ERROR, "❌ ")
            return False
    
    def _handle_email_otp_verification(self):
        """Handle email OTP verification"""
        try:
            otp = self.wait_and_get_otp()
            if otp:
                if self.enter_otp_code(otp):
                    time.sleep(5)  # Wait for redirect
                    return True
            self.log("Failed to get OTP from email", Colors.ERROR, "❌ ")
            return False
        except Exception as e:
            self.log(f"Error handling email OTP: {e}", Colors.ERROR, "❌ ")
            return False
                

                

    def fix_base32_padding(self, secret):
        """Fix base32 padding for pyotp"""
        secret = secret.rstrip("=")
        padding_needed = (8 - len(secret) % 8) % 8
        return secret + "=" * padding_needed
    
    def extract_totp_secret(self):
        """Extract TOTP secret from the QR code dialog"""
        try:
            self.log("Extracting TOTP secret...", Colors.INFO, "🔍 ")
            
            wait = WebDriverWait(self.driver, 2)
            
            trouble_link_selectors = [
                "//a[contains(text(), 'Trouble scanning')]",
                "//button[contains(text(), 'Trouble scanning')]",
                "//*[contains(text(), 'Trouble') and contains(text(), 'scanning')]",
            ]
            
            for selector in trouble_link_selectors:
                try:
                    link = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    self.driver.execute_script("arguments[0].click();", link)
                    self.log("Clicked 'Trouble scanning?' link", Colors.SUCCESS, "✅ ")
                    time.sleep(1)
                    break
                except:
                    continue
            
            secret = None
            secret_selectors = [
                "//code", "//pre", "//input[@readonly]", "//input[@value]",
                "//*[contains(@class, 'font-mono')]",
                "//p[string-length(normalize-space(text())) >= 16]",
                "//span[string-length(normalize-space(text())) >= 16]",
            ]
            
            for selector in secret_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        text = elem.text.strip().replace(" ", "").replace("-", "")
                        if not text:
                            text = elem.get_attribute("value") or ""
                            text = text.strip().replace(" ", "").replace("-", "")
                        
                        if text and len(text) >= 16 and len(text) <= 64:
                            clean_text = text.upper()
                            valid_chars = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ234567")
                            if all(c in valid_chars for c in clean_text):
                                secret = clean_text
                                break
                except:
                    continue
                if secret:
                    break
            
            if not secret:
                try:
                    page_source = self.driver.page_source
                    match = re.search(r'secret=([A-Z2-7]+)', page_source, re.IGNORECASE)
                    if match:
                        secret = match.group(1).upper()
                except:
                    pass
            
            if secret:
                secret = self.fix_base32_padding(secret)
                self.log(f"Got TOTP secret: {secret[:8]}...", Colors.SUCCESS, "✅ ")
                return secret
            else:
                self.log("Could not extract TOTP secret", Colors.ERROR, "❌ ")
                return None
                
        except Exception as e:
            self.log(f"Error extracting secret: {e}", Colors.ERROR, "❌ ")
            return None
    
    def enter_otp_and_verify(self, secret):
        """Generate OTP and enter it to verify"""
        try:
            self.log("Generating and entering OTP...", Colors.INFO, "🔢 ")
            
            totp = pyotp.TOTP(secret)
            code = totp.now()
            self.log(f"Generated OTP: {code}", Colors.INFO, "🔢 ")
            
            wait = WebDriverWait(self.driver, 5)
            
            otp_input = wait.until(EC.presence_of_element_located((
                By.XPATH, 
                "//input[contains(@placeholder, '6-digit') or contains(@placeholder, 'code') or @type='text']"
            )))
            
            otp_input.clear()
            for char in code:
                otp_input.send_keys(char)
                time.sleep(0.1)
            
            self.log("Entered OTP code", Colors.SUCCESS, "✅ ")
            time.sleep(1)
            
            verify_selectors = [
                "//button[contains(@class, 'btn-primary') and .//div[contains(text(), 'Verify')]]",
                "//button[contains(@class, 'btn-primary') and contains(., 'Verify')]",
                "//button[.//div[text()='Verify']]",
                "//button[contains(@class, 'btn-primary')]",
            ]
            
            verify_btn = None
            for selector in verify_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        if elem.is_displayed() and 'Verify' in elem.text:
                            verify_btn = elem
                            break
                    if verify_btn:
                        break
                except:
                    continue
            
            if verify_btn:
                self.driver.execute_script("arguments[0].click();", verify_btn)
                self.log("Clicked Verify", Colors.SUCCESS, "✅ ")
            else:
                self.log("Could not find Verify button", Colors.ERROR, "❌ ")
                return False
            
            time.sleep(2)
            
            success_selectors = [
                "//*[@data-testid='totp-mfa-activation-success']",
                "//*[@role='alert' and contains(., 'Authenticator app enabled')]",
                "//*[contains(text(), 'Authenticator app enabled')]"
            ]
            
            success_found = False
            for selector in success_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    if elements:
                        for elem in elements:
                            if elem.is_displayed():
                                success_found = True
                                break
                except:
                    continue
                if success_found:
                    break
            
            if success_found:
                self.log("MFA activated successfully!", Colors.SUCCESS, "🎉 ")
                return True
            else:
                page_text = self.driver.page_source.lower()
                if "enabled" in page_text or "success" in page_text:
                    self.log("MFA likely activated", Colors.SUCCESS, "✅ ")
                    return True
                else:
                    self.log("Could not confirm MFA activation", Colors.WARNING, "⚠️ ")
                    return False
                
        except Exception as e:
            self.log(f"Error verifying OTP: {e}", Colors.ERROR, "❌ ")
            return False
    
    def enroll_mfa(self):
        """Main MFA enrollment flow"""
        try:
            if not self.navigate_to_security_settings():
                return None
            if not self.click_mfa_toggle():
                return None
            if not self.handle_password_verification():
                return None
            secret = self.extract_totp_secret()
            if not secret:
                return None
            return {"secret": secret}
        except Exception as e:
            self.log(f"Error in MFA enrollment: {e}", Colors.ERROR, "❌ ")
            return None
    
    def activate_mfa(self, enroll_data):
        """Activate MFA by entering OTP"""
        try:
            secret = enroll_data.get("secret")
            if not secret:
                return None
            if self.enter_otp_and_verify(secret):
                return secret
            else:
                return None
        except Exception as e:
            self.log(f"Error activating MFA: {e}", Colors.ERROR, "❌ ")
            return None
    
    def save_to_excel(self, totp_secret):
        """Save TOTP secret to Excel"""
        try:
            with file_lock:
                wb = load_workbook(self.excel_file)
                ws = wb.active
                ws.cell(row=self.row_index, column=5, value=totp_secret)  # Column E = 2FA Secret
                wb.save(self.excel_file)
                self.log(f"Saved TOTP secret to Excel row {self.row_index}", Colors.SUCCESS, "💾 ")
                return True
        except Exception as e:
            self.log(f"Error saving to Excel: {e}", Colors.ERROR, "❌ ")
            return False
    
    def cleanup(self):
        """Cleanup driver and temp files"""
        # Stop popup polling thread first
        self.stop_popup_polling()
        
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

        
        if self.user_data_dir and os.path.exists(self.user_data_dir):
            try:
                shutil.rmtree(self.user_data_dir, ignore_errors=True)
            except:
                pass
            self.user_data_dir = None
    
    def run(self):
        """Main worker flow"""
        try:
            self.log(f"Processing: {self.email}", Colors.HEADER, "🚀 ")
            
            if not self.setup_driver():
                return False
            if not self.import_cookies():
                return False
            
            enroll_data = self.enroll_mfa()
            if not enroll_data:
                return False
            
            totp_secret = self.activate_mfa(enroll_data)
            if not totp_secret:
                return False
            
            self.save_to_excel(totp_secret)
            
            self.log(f"✅ MFA enabled for {self.email}", Colors.SUCCESS, "🎉 ")
            return True
            
        except Exception as e:
            self.log(f"Error in worker: {e}", Colors.ERROR, "❌ ")
            return False
        finally:
            self.cleanup()


def load_mfa_accounts(excel_file):
    """Load accounts for MFA from Excel"""
    wb = load_workbook(excel_file)
    ws = wb.active
    
    accounts = []
    skipped_count = 0
    
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if len(row) < 2:
            continue
        
        account = row[0] if row[0] else ""
        cookie = row[1] if len(row) > 1 and row[1] else ""
        two_fa = row[4] if len(row) > 4 and row[4] else ""  # Column E = 2FA Secret
        
        if not account or not cookie:
            continue
        
        if two_fa:
            skipped_count += 1
            continue
        
        if ":" in str(account):
            parts = str(account).split(":", 1)
            email = parts[0]
            password = parts[1] if len(parts) > 1 else ""
        else:
            email = str(account)
            password = ""
        
        accounts.append({
            "row_index": row_idx,
            "email": email,
            "password": password,
            "cookie": cookie
        })
    
    if skipped_count > 0:
        print(f"{Colors.WARNING}Skipped {skipped_count} accounts (already have 2fa){Colors.RESET}")
    
    return accounts


def run_mfa_worker(thread_id, account, excel_file, stop_event=None, thread_delay=2, slot_index=0, is_first_batch=True, oauth2_account=None):
    """Run MFA worker thread with retries and staggered start
    
    Delay logic:
    - First batch: Thread 1 = 0s, Thread 2 = delay, Thread 3 = delay × 2
    - Subsequent batches: Thread 1 = 5s, Thread 2 = 5s + delay, Thread 3 = 5s + delay × 2
    """
    base_delay = 0 if is_first_batch else 5  # 5s base delay after first batch
    slot_delay = slot_index * thread_delay    # slot_index: 0, 1, 2 for 3 threads
    total_delay = base_delay + slot_delay
    
    if total_delay > 0:
        safe_print(thread_id, f"Waiting {total_delay:.1f}s before starting...", Colors.INFO, "⏳ ")
        # Sleep in intervals to check stop_event
        for _ in range(int(total_delay * 2)):
            if stop_event and stop_event.is_set():
                return False, account["email"]
            time.sleep(0.5)
    
    max_retries = 3
    for attempt in range(max_retries):
        if stop_event and stop_event.is_set():
            return False, account["email"]
            
        worker = MFAWorker(
            thread_id=thread_id,
            row_index=account["row_index"],
            email=account["email"],
            password=account["password"],
            cookie_json=account["cookie"],
            excel_file=excel_file,
            oauth2_account=oauth2_account
        )
        worker.stop_event = stop_event
        
        result = worker.run()
        if result:
            return True, account["email"]
            
        if attempt < max_retries - 1:
            if stop_event and stop_event.is_set():
                return False, account["email"]
            print(f"Retry {attempt+1}/{max_retries} for {account['email']}...")
            time.sleep(3)
            
    return False, account["email"]


# ============================================================================
# MODULE 4: CHECKOUT CAPTURE
# ============================================================================

class CheckoutCaptureWorker:
    """Worker for capturing checkout links from existing accounts"""
    
    def __init__(self, thread_id, email, cookie_json, excel_file, row_index, checkout_type="Plus"):
        self.thread_id = thread_id
        self.email = email
        self.cookie_json = cookie_json
        self.excel_file = excel_file
        self.row_index = row_index
        self.checkout_type = checkout_type  # "Plus", "Business", or "Both"
        self.driver = None
        self.user_data_dir = None
        self.stop_event = None
        self.net = NETWORK_SETTINGS.get(NETWORK_MODE, NETWORK_SETTINGS["Fast"])
        
        # Popup polling
        self._popup_polling_active = False
        self._popup_polling_thread = None
        self._popup_dismissed = False
        
    def log(self, message, color=Colors.INFO, emoji=""):
        safe_print(self.thread_id, message, color, emoji)
    
    def setup_driver(self):
        """Initialize Chrome driver with lock to prevent race condition"""
        try:
            self.log("Initializing browser...", Colors.INFO, "🔄 ")
            
            timestamp = int(time.time())
            random_suffix = random.randint(1000, 9999)
            self.user_data_dir = tempfile.mkdtemp(prefix=f"chrome_checkout_T{self.thread_id}_{timestamp}_{random_suffix}_")
            
            options = uc.ChromeOptions()
            options.add_argument(f"--user-data-dir={self.user_data_dir}")
            options.add_argument("--no-first-run")
            options.add_argument("--no-default-browser-check")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--lang=en-US")  # Force English UI
            
            # Use lock to prevent race condition when multiple threads initialize ChromeDriver
            with driver_init_lock:
                self.driver = uc.Chrome(options=options, version_main=CHROME_VERSION_MAIN)
            
            # Window size same as MFA module
            window_width = 900
            window_height = 700
            x_offset = (self.thread_id - 1) * 80
            y_offset = (self.thread_id - 1) * 40
            self.driver.set_window_size(window_width, window_height)
            self.driver.set_window_position(x_offset, y_offset)
            
            self.log("Browser ready!", Colors.SUCCESS, "✅ ")
            return True
            
        except Exception as e:
            self.log(f"Failed to start browser: {e}", Colors.ERROR, "❌ ")
            return False
    
    def cleanup_browser(self):
        """Close browser and clean up"""
        # Stop popup polling thread first
        self.stop_popup_polling()
        
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

        if self.user_data_dir and os.path.exists(self.user_data_dir):
            try:
                shutil.rmtree(self.user_data_dir, ignore_errors=True)
            except:
                pass
    
    def import_cookies(self):
        """Import cookies into browser"""
        try:
            self.log("Importing cookies...", Colors.INFO, "🍪 ")
            
            # First navigate to the domain
            self.driver.get("https://chatgpt.com")
            time.sleep(2)
            
            # Parse and import cookies
            cookies = json.loads(self.cookie_json)
            for cookie in cookies:
                try:
                    # Clean up cookie for selenium
                    cookie_dict = {
                        'name': cookie.get('name'),
                        'value': cookie.get('value'),
                        'domain': cookie.get('domain', '.chatgpt.com'),
                        'path': cookie.get('path', '/'),
                    }
                    if cookie.get('secure'):
                        cookie_dict['secure'] = True
                    if cookie.get('httpOnly'):
                        cookie_dict['httpOnly'] = True
                        
                    self.driver.add_cookie(cookie_dict)
                except Exception:
                    continue
            
            # Refresh to apply cookies
            self.driver.refresh()
            time.sleep(self.net["extra_wait"])
            
            self.log("Cookies imported!", Colors.SUCCESS, "✅ ")
            
            # Start popup polling thread
            self.start_popup_polling()
            
            return True
            
        except Exception as e:
            self.log(f"Failed to import cookies: {e}", Colors.ERROR, "❌ ")
            return False
    
    def start_popup_polling(self):
        """Start background thread to continuously check and dismiss onboarding popup"""
        if self._popup_polling_active or self._popup_dismissed:
            return
        
        self._popup_polling_active = True
        self._popup_polling_thread = threading.Thread(target=self._popup_polling_loop, daemon=True)
        self._popup_polling_thread.start()
    
    def stop_popup_polling(self):
        """Stop the popup polling thread"""
        self._popup_polling_active = False
        if self._popup_polling_thread and self._popup_polling_thread.is_alive():
            self._popup_polling_thread.join(timeout=2)
        self._popup_polling_thread = None
    
    def _popup_polling_loop(self):
        """Background loop to check for onboarding popup every 2 seconds"""
        while self._popup_polling_active and not self._popup_dismissed:
            try:
                if not self.driver:
                    break
                
                button_selectors = [
                    "//button[.//div[contains(text(), \"Okay, let's go\")]]",
                    "//button[contains(., \"Okay, let's go\")]",
                    "//div[@role='dialog']//button[contains(., 'Okay')]",
                ]
                
                for selector in button_selectors:
                    try:
                        buttons = self.driver.find_elements(By.XPATH, selector)
                        for btn in buttons:
                            if btn.is_displayed():
                                self.driver.execute_script("arguments[0].click();", btn)
                                self.log("Dismissed onboarding popup", Colors.SUCCESS, "✅ ")
                                self._popup_dismissed = True
                                self._popup_polling_active = False
                                return
                    except:
                        continue
            except:
                pass
            
            # Poll every 2 seconds
            time.sleep(2)

    
    def check_payment_error(self):
        """Check if payment error message is displayed using proper element"""
        try:
            error_selectors = [
                "//div[contains(@class, 'bg-orange-500') and @role='alert']//div[contains(@class, 'text-start')]",
                "//div[@role='alert' and contains(@class, 'border-orange-500')]//div[contains(@class, 'whitespace-pre-wrap')]",
                "//div[contains(@class, 'bg-orange-500')]//div[contains(., 'payments page encountered an error')]",
            ]
            for selector in error_selectors:
                try:
                    error_elements = self.driver.find_elements(By.XPATH, selector)
                    for el in error_elements:
                        if "payments page encountered an error" in el.text.lower():
                            return True
                except:
                    continue
            return False
        except:
            return False

    def poll_for_checkout_or_error(self, timeout=30, poll_interval=0.5):
        """Poll continuously for checkout URL or payment error"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Check for payment error first
                if self.check_payment_error():
                    return "PAYMENT_ERROR"
                
                # Check for checkout URL
                current_url = self.driver.current_url or ""
                if "/checkout/" in current_url.lower() or "pay.openai.com" in current_url.lower():
                    # Check for invalid verify URL (e.g. /checkout/verify?stripe_session_id=)
                    if "verify?stripe_session_id" in current_url.lower() or "/checkout/verify?" in current_url.lower():
                        self.log("Invalid verify URL detected, need to retry...", Colors.WARNING, "⚠️ ")
                        return "VERIFY_ERROR"
                    return current_url
            except:
                pass
            
            time.sleep(poll_interval)
        
        return None  # Timeout

    def ensure_personal_tab_active(self):
        """Ensure Personal tab is active, not Business tab"""
        try:
            # Check if Personal tab is active (data-state="on" means active)
            personal_active_xpath = "//button[@role='radio' and contains(., 'Personal') and @data-state='on']"
            personal_inactive_xpath = "//button[@role='radio' and contains(., 'Personal') and @data-state='off']"
            
            # Check if Personal is already active
            active_tabs = self.driver.find_elements(By.XPATH, personal_active_xpath)
            if active_tabs:
                return True  # Already on Personal tab
            
            # Personal is not active, try to click it
            inactive_tabs = self.driver.find_elements(By.XPATH, personal_inactive_xpath)
            if inactive_tabs:
                self.log("Switching to Personal tab...", Colors.INFO, "🔄 ")
                inactive_tabs[0].click()
                time.sleep(1)
                return True
            
            return False
        except Exception as e:
            return False

    def get_plus_checkout(self):
        """Get Plus checkout link with retry logic"""
        plus_url = "https://chatgpt.com/#pricing"
        max_attempts = 3
        
        for attempt in range(1, max_attempts + 1):
            try:
                self.log(f"Getting Plus checkout link (attempt {attempt}/{max_attempts})...", Colors.INFO, "💳 ")
                
                # Navigate to pricing
                self.driver.get(plus_url)
                time.sleep(2)
                
                wait = WebDriverWait(self.driver, self.net["element_timeout"])
                
                # Ensure Personal tab is active (not Business)
                self.ensure_personal_tab_active()
                time.sleep(1)
                
                # Check if Plus is a free offer (price = $0)
                # Only check on first attempt to avoid redundant checks
                if attempt == 1:
                    try:
                        # Look for the price element with text-5xl class but NOT line-through
                        # line-through = original price (crossed out)
                        # text-5xl without line-through = actual discounted price
                        price_selectors = [
                            "//div[@data-testid='plus-pricing-column-cost']//div[contains(@class, 'text-5xl') and not(contains(@class, 'line-through'))]",
                            "//div[contains(@class, 'plus-pricing')]//div[contains(@class, 'text-5xl') and not(contains(@class, 'line-through'))]",
                        ]
                        
                        actual_price = None
                        for selector in price_selectors:
                            price_elements = self.driver.find_elements(By.XPATH, selector)
                            for price_elem in price_elements:
                                if price_elem.is_displayed():
                                    # Double-check class doesn't contain line-through
                                    elem_class = price_elem.get_attribute("class") or ""
                                    if "line-through" in elem_class:
                                        continue  # Skip strikethrough prices
                                    
                                    price_text = price_elem.text.strip()
                                    if price_text:
                                        actual_price = price_text
                                        break
                            if actual_price:
                                break
                        
                        # Check if actual price is NOT free ($0, 0, ₩0)
                        if actual_price and actual_price not in ["$0", "0", "₩0"]:
                            # Check if it contains any non-zero digits
                            digits = ''.join(filter(str.isdigit, actual_price))
                            if digits and int(digits) > 0:
                                self.log(f"Plus price is {actual_price} - no free offer available", Colors.WARNING, "⚠️ ")
                                return "NO_PLUS_OFFER"
                        else:
                            self.log(f"Plus offer detected: {actual_price}", Colors.SUCCESS, "✅ ")
                            
                    except Exception as e:
                        self.log(f"Could not check price, continuing...", Colors.WARNING, "⚠️ ")

                
                # Find and click Get Plus button
                button_xpath = (
                    "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'get plus')]"
                    " | //button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'claim free offer')]"
                    " | //a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'get plus')]"
                )

                
                button = None
                try:
                    button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                except:
                    self.log("Get Plus button not found, ensuring Personal tab...", Colors.WARNING, "⚠️ ")
                    # Try to switch to Personal tab and retry finding button
                    self.ensure_personal_tab_active()
                    time.sleep(1)
                    try:
                        button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                    except:
                        self.log("Get Plus button still not found", Colors.WARNING, "⚠️ ")
                
                if not button:
                    if attempt < max_attempts:
                        continue
                    return None
                
                handles_before = list(self.driver.window_handles)
                button.click()
                self.log("Clicked Get Plus", Colors.SUCCESS, "✅ ")
                time.sleep(1)
                
                # Check for new tab
                handles_after = self.driver.window_handles
                new_handles = [h for h in handles_after if h not in handles_before]
                if new_handles:
                    self.driver.switch_to.window(new_handles[-1])
                    self.log("Switched to checkout tab", Colors.INFO, "🪟 ")
                
                # Use polling to check for checkout URL or error
                self.log("Waiting for checkout page...", Colors.INFO, "⏳ ")
                result = self.poll_for_checkout_or_error(timeout=30, poll_interval=0.5)
                
                if result == "PAYMENT_ERROR" or result == "VERIFY_ERROR":
                    error_type = "payment error" if result == "PAYMENT_ERROR" else "verify URL error"
                    self.log(f"{error_type.capitalize()} detected!", Colors.WARNING, "⚠️ ")
                    if attempt < max_attempts:
                        self.log(f"Retrying due to {error_type}... ({attempt + 1}/{max_attempts})", Colors.WARNING, "🔄 ")
                        # Close error tab if exists
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        # Navigate back to pricing page and ensure Personal tab
                        self.log("Refreshing and navigating back to pricing...", Colors.INFO, "🔄 ")
                        self.driver.get(plus_url)
                        time.sleep(2)
                        self.ensure_personal_tab_active()
                        time.sleep(1)
                        continue
                    else:
                        self.log(f"{error_type.capitalize()} persists, cannot capture Plus link", Colors.ERROR, "❌ ")
                        return None
                elif result:
                    self.log(f"Plus URL: {result[:50]}...", Colors.SUCCESS, "✅ ")
                    return result
                else:
                    # Timeout
                    if attempt < max_attempts:
                        self.log(f"Timeout, retrying... ({attempt + 1}/{max_attempts})", Colors.WARNING, "🔄 ")
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        # Navigate back to pricing page
                        self.driver.get(plus_url)
                        time.sleep(2)
                        self.ensure_personal_tab_active()
                        continue
                    return None
                    
            except Exception as e:
                self.log(f"Error getting Plus checkout: {e}", Colors.ERROR, "❌ ")
                if attempt < max_attempts:
                    if len(self.driver.window_handles) > 1:
                        try:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        except:
                            pass
                    continue
                return None
        
        return None
    
    def get_business_checkout(self):
        """Get Business checkout link with retry logic"""
        business_url = "https://chatgpt.com/?numSeats=5&selectedPlan=month&referrer=#team-pricing-seat-selection"
        max_attempts = 3
        
        for attempt in range(1, max_attempts + 1):
            try:
                self.log(f"Getting Business checkout link (attempt {attempt}/{max_attempts})...", Colors.INFO, "💼 ")
                
                # Navigate to business pricing
                self.driver.get(business_url)
                time.sleep(3)
                
                wait = WebDriverWait(self.driver, self.net["element_timeout"])
                
                # Find Continue to billing button
                button_xpath = (
                    "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'continue to billing')]"
                    " | //button[contains(@class, 'btn-green')]"
                )
                
                try:
                    button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                except:
                    self.log("Continue to billing button not found", Colors.WARNING, "⚠️ ")
                    if attempt < max_attempts:
                        continue
                    return None
                
                handles_before = list(self.driver.window_handles)
                
                # Try multiple click methods
                try:
                    button.click()
                except Exception:
                    try:
                        self.driver.execute_script("arguments[0].click();", button)
                    except Exception:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(self.driver).move_to_element(button).click().perform()
                
                self.log("Clicked Continue to billing", Colors.SUCCESS, "✅ ")
                time.sleep(1)
                
                # Check for new tab
                handles_after = self.driver.window_handles
                new_handles = [h for h in handles_after if h not in handles_before]
                if new_handles:
                    self.driver.switch_to.window(new_handles[-1])
                    self.log("Switched to checkout tab", Colors.INFO, "🪟 ")
                
                # Use polling to check for checkout URL or error
                self.log("Waiting for checkout page...", Colors.INFO, "⏳ ")
                result = self.poll_for_checkout_or_error(timeout=30, poll_interval=0.5)
                
                if result == "PAYMENT_ERROR" or result == "VERIFY_ERROR":
                    error_type = "payment error" if result == "PAYMENT_ERROR" else "verify URL error"
                    self.log(f"{error_type.capitalize()} detected!", Colors.WARNING, "⚠️ ")
                    if attempt < max_attempts:
                        self.log(f"Retrying due to {error_type}... ({attempt + 1}/{max_attempts})", Colors.WARNING, "🔄 ")
                        # Close error tab if exists
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        # Navigate back to business pricing page
                        self.log("Refreshing and navigating back to business pricing...", Colors.INFO, "🔄 ")
                        self.driver.get(business_url)
                        time.sleep(3)
                        continue
                    else:
                        self.log(f"{error_type.capitalize()} persists, cannot capture Business link", Colors.ERROR, "❌ ")
                        return None
                elif result:
                    self.log(f"Business URL: {result[:50]}...", Colors.SUCCESS, "✅ ")
                    return result
                else:
                    # Timeout
                    if attempt < max_attempts:
                        self.log(f"Timeout, retrying... ({attempt + 1}/{max_attempts})", Colors.WARNING, "🔄 ")
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        # Navigate back to business pricing page
                        self.driver.get(business_url)
                        time.sleep(3)
                        continue
                    return None
                    
            except Exception as e:
                self.log(f"Error getting Business checkout: {e}", Colors.ERROR, "❌ ")
                if attempt < max_attempts:
                    if len(self.driver.window_handles) > 1:
                        try:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        except:
                            pass
                    continue
                return None
        
        return None
    
    def save_to_excel(self, plus_url=None, business_url=None):
        """Save checkout URLs to Excel"""
        try:
            with file_lock:
                wb = load_workbook(self.excel_file)
                ws = wb.active
                
                if plus_url:
                    ws.cell(row=self.row_index, column=3, value=plus_url)
                    self.log("Saved Plus URL to Excel", Colors.SUCCESS, "💾 ")
                    
                if business_url:
                    ws.cell(row=self.row_index, column=4, value=business_url)
                    self.log("Saved Business URL to Excel", Colors.SUCCESS, "💾 ")
                
                wb.save(self.excel_file)
                wb.close()
                
            return True
        except Exception as e:
            self.log(f"Failed to save to Excel: {e}", Colors.ERROR, "❌ ")
            return False
    
    def save_no_plus_offer(self):
        """Save 'no Plus offer' to Excel column C when account has no free Plus offer"""
        try:
            with file_lock:
                wb = load_workbook(self.excel_file)
                ws = wb.active
                ws.cell(row=self.row_index, column=3, value="no Plus offer")
                wb.save(self.excel_file)
                wb.close()
                self.log("Saved 'no Plus offer' to Excel", Colors.SUCCESS, "💾 ")
            return True
        except Exception as e:
            self.log(f"Failed to save no Plus offer: {e}", Colors.ERROR, "❌ ")
            return False
    
    def run(self):
        """Run the checkout capture flow"""
        try:
            self.log(f"Starting checkout capture for {self.email}", Colors.HEADER, "🚀 ")
            
            if not self.setup_driver():
                return False
            
            if not self.import_cookies():
                self.cleanup_browser()
                return False
            
            plus_url = None
            business_url = None
            no_plus_offer = False
            
            if self.checkout_type in ["Plus", "Both"]:
                result = self.get_plus_checkout()
                
                # Handle NO_PLUS_OFFER case
                if result == "NO_PLUS_OFFER":
                    no_plus_offer = True
                    self.log("Account has no Plus offer, saving to Excel...", Colors.WARNING, "⚠️ ")
                    self.save_no_plus_offer()
                else:
                    plus_url = result
                
                # Close extra tabs for Business
                if self.checkout_type == "Both" and (plus_url or no_plus_offer):
                    try:
                        while len(self.driver.window_handles) > 1:
                            self.driver.switch_to.window(self.driver.window_handles[-1])
                            self.driver.close()
                        self.driver.switch_to.window(self.driver.window_handles[0])
                        time.sleep(1)
                    except:
                        pass
            
            if self.checkout_type in ["Business", "Both"]:
                business_url = self.get_business_checkout()
            
            # Save results
            if plus_url or business_url:
                self.save_to_excel(plus_url, business_url)
                self.log(f"✅ Completed for {self.email}", Colors.SUCCESS, "🎉 ")
                self.cleanup_browser()
                return True
            elif no_plus_offer and self.checkout_type == "Plus":
                # Only Plus was requested and no offer available
                self.log(f"No Plus offer for {self.email}", Colors.WARNING, "⚠️ ")
                self.cleanup_browser()
                return True  # Still consider success since we saved the info
            elif no_plus_offer and business_url:
                # No Plus offer but got Business
                self.save_to_excel(None, business_url)
                self.log(f"✅ Business link captured for {self.email}", Colors.SUCCESS, "🎉 ")
                self.cleanup_browser()
                return True
            else:
                self.log(f"No checkout URLs captured for {self.email}", Colors.WARNING, "⚠️ ")
                self.cleanup_browser()
                return False

                
        except Exception as e:
            self.log(f"Error: {e}", Colors.ERROR, "❌ ")
            self.cleanup_browser()
            return False


def load_checkout_accounts(excel_file):
    """Load all accounts from Excel for checkout capture"""
    try:
        wb = load_workbook(excel_file)
        ws = wb.active
        
        accounts = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if len(row) < 2:
                continue
            
            account = row[0] if row[0] else ""
            cookie = row[1] if len(row) > 1 and row[1] else ""
            plus_url = row[2] if len(row) > 2 and row[2] else ""
            business_url = row[3] if len(row) > 3 and row[3] else ""
            sold_status = row[5] if len(row) > 5 and row[5] else ""  # Column F (index 5)
            
            if not account or not cookie:
                continue
            
            # Parse email from account
            if ":" in str(account):
                email = str(account).split(":", 1)[0]
            else:
                email = str(account)
            
            # Check if sold (any value in column F means sold)
            is_sold = bool(sold_status and str(sold_status).strip())
            
            accounts.append({
                "row_index": row_idx,
                "email": email,
                "account": account,
                "cookie": cookie,
                "plus_url": plus_url,
                "business_url": business_url,
                "sold_status": sold_status,
                "is_sold": is_sold,
            })
        
        wb.close()
        return accounts
        
    except Exception as e:
        print(f"Error loading accounts: {e}")
        return []


# ============================================================================
# GUI IMPLEMENTATION
# ============================================================================

# Settings
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ============================================================================
# ENHANCED UI UTILITIES
# ============================================================================

class GlowButton(ctk.CTkButton):
    """Custom button with glow effect on hover"""
    def __init__(self, master, glow_color="#00d4ff", **kwargs):
        super().__init__(master, **kwargs)
        self.glow_color = glow_color
        self.default_border = kwargs.get("border_width", 0)
        self.bind("<Enter>", self._on_hover_enter)
        self.bind("<Leave>", self._on_hover_leave)
        
    def _on_hover_enter(self, e=None):
        try:
            self.configure(border_width=2, border_color=self.glow_color)
        except: pass
        
    def _on_hover_leave(self, e=None):
        try:
            self.configure(border_width=self.default_border, border_color="transparent")
        except: pass


class AnimatedCard(ctk.CTkFrame):
    """Card with entrance animation"""
    def __init__(self, master, delay_ms=0, **kwargs):
        # Default styling
        kwargs.setdefault("corner_radius", 16)
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", "#30363d")
        super().__init__(master, **kwargs)
        self.delay_ms = delay_ms
        self._initial_alpha = 0
        
    def animate_in(self, motion):
        """Trigger entrance animation"""
        def do_animate():
            motion.color(self, "fg_color", self.cget("fg_color"), duration_ms=400, steps=30)
        self.after(self.delay_ms, do_animate)


class PulsingDot(ctk.CTkFrame):
    """Animated status dot"""
    def __init__(self, master, color="#00ff88", size=12, **kwargs):
        super().__init__(master, width=size, height=size, corner_radius=size//2, fg_color=color, **kwargs)
        self.base_color = color
        self.is_pulsing = False
        self._pulse_job = None
        
    def start_pulse(self):
        if self.is_pulsing:
            return
        self.is_pulsing = True
        self._do_pulse(True)
        
    def _do_pulse(self, bright):
        if not self.is_pulsing:
            return
        try:
            target = self.base_color if bright else self._dim_color(self.base_color)
            self.configure(fg_color=target)
            self._pulse_job = self.after(500, lambda: self._do_pulse(not bright))
        except: pass
        
    def _dim_color(self, hex_color):
        """Dim a hex color by 50%"""
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return f"#{r//2:02x}{g//2:02x}{b//2:02x}"
        
    def stop_pulse(self):
        self.is_pulsing = False
        if self._pulse_job:
            try:
                self.after_cancel(self._pulse_job)
            except: pass
        self.configure(fg_color=self.base_color)


class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
        # ANSI codes map
        self.tag_map = {
            '\033[31m': 'error',     # Red
            '\033[32m': 'success',   # Green
            '\033[33m': 'warning',   # Yellow
            '\033[34m': 'info',      # Blue
            '\033[36m': 'info',      # Cyan (mapped to info)
            '\033[35m': 'header',    # Magenta
            '\033[0m': 'normal',     # Reset
            '\x1b[31m': 'error',
            '\x1b[32m': 'success',
            '\x1b[33m': 'warning',
            '\x1b[34m': 'info',
            '\x1b[36m': 'info',
            '\x1b[35m': 'header',
            '\x1b[0m': 'normal',
            '\x1b[39m': 'normal'
        }
        self.ansi_re = re.compile(r'(\x1B\[[0-9;]*m)')

    def write(self, str_data):
        if str_data:
            self.widget.after(0, self._append_text, str_data)

    def _append_text(self, text):
        try:
            self.widget.configure(state="normal")
            parts = self.ansi_re.split(text)
            current_tag = "normal"
            
            for part in parts:
                if part in self.tag_map:
                    current_tag = self.tag_map[part]
                elif part.startswith('\x1B'):
                    pass 
                elif part:
                    self.widget.insert("end", part, current_tag)
            
            self.widget.see("end")
            self.widget.configure(state="disabled")
        except:
            pass

    def flush(self):
        pass

# =========================
# UI MOTION SYSTEM (CTk)
# =========================
import time
import math

class MotionTokens:
    # Durations (ms)
    fast = 120
    normal = 180
    slow = 260
    pulse_period = 950

    # Easing
    @staticmethod
    def ease_out_quad(t: float) -> float:
        return 1.0 - (1.0 - t) * (1.0 - t)

    @staticmethod
    def ease_out_cubic(t: float) -> float:
        return 1.0 - (1.0 - t) ** 3

def _hex_to_rgb(h: str):
    h = h.strip()
    if h.startswith("#"): h = h[1:]
    if len(h) == 3: h = "".join([c * 2 for c in h])
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

def _rgb_to_hex(rgb):
    return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"

def _lerp(a, b, t): 
    return a + (b - a) * t

class Motion:
    """
    Motion engine cho CustomTkinter:
    - cancel theo key
    - color tween
    - number tween
    - pulse loop
    - hover transitions
    """
    def __init__(self, app, tokens=MotionTokens):
        self.app = app
        self.tok = tokens
        self._jobs = {}          # key -> after job id
        self._pulse_jobs = {}    # key -> after job id
        self._pulse_on = set()   # keys đang pulse

    def _key(self, widget, prop, suffix=""):
        return f"{id(widget)}::{prop}::{suffix}"

    def cancel(self, key: str):
        job = self._jobs.pop(key, None)
        if job is not None:
            try: self.app.after_cancel(job)
            except: pass

    def color(self, widget, prop: str, to_hex: str, *,
              duration_ms=None, steps=18, easing=None):
        duration_ms = duration_ms if duration_ms is not None else self.tok.normal
        easing = easing if easing is not None else self.tok.ease_out_quad

        try:
            cur = widget.cget(prop)
        except Exception:
            try:
                widget.configure(**{prop: to_hex})
                return
            except:
                return

        if not isinstance(cur, str) or not cur.startswith("#") or len(cur) not in (4, 7):
            try:
                widget.configure(**{prop: to_hex})
            except:
                pass
            return

        try:
            sr, sg, sb = _hex_to_rgb(cur)
            er, eg, eb = _hex_to_rgb(to_hex)
        except:
            try: widget.configure(**{prop: to_hex})
            except: pass
            return

        key = self._key(widget, prop, "color")
        self.cancel(key)

        start_time = time.time()

        def tick():
            t = (time.time() - start_time) * 1000.0 / max(duration_ms, 1)
            t = max(0.0, min(1.0, t))
            te = easing(t)

            r = int(_lerp(sr, er, te))
            g = int(_lerp(sg, eg, te))
            b = int(_lerp(sb, eb, te))

            try:
                widget.configure(**{prop: _rgb_to_hex((r, g, b))})
            except:
                pass

            if t < 1.0:
                job = self.app.after(max(1, duration_ms // steps), tick)
                self._jobs[key] = job
            else:
                self._jobs.pop(key, None)

        tick()

    def number(self, setter, start: float, end: float, *,
               duration_ms=None, steps=18, easing=None, fmt=None):
        duration_ms = duration_ms if duration_ms is not None else self.tok.normal
        easing = easing if easing is not None else self.tok.ease_out_quad
        key = f"{id(setter)}::number"

        self.cancel(key)
        start_time = time.time()

        def tick():
            t = (time.time() - start_time) * 1000.0 / max(duration_ms, 1)
            t = max(0.0, min(1.0, t))
            te = easing(t)
            val = _lerp(start, end, te)
            try:
                setter(fmt(val) if fmt else val)
            except:
                pass

            if t < 1.0:
                job = self.app.after(max(1, duration_ms // steps), tick)
                self._jobs[key] = job
            else:
                self._jobs.pop(key, None)

        tick()

    def pulse(self, widget, prop: str, a: str, b: str, *,
              period_ms=None):
        period_ms = period_ms if period_ms is not None else self.tok.pulse_period
        key = self._key(widget, prop, "pulse")

        self._pulse_on.add(key)

        job = self._pulse_jobs.pop(key, None)
        if job is not None:
            try: self.app.after_cancel(job)
            except: pass

        def loop(state=0):
            if key not in self._pulse_on:
                return
            start, end = (a, b) if state == 0 else (b, a)
            self.color(widget, prop, end, duration_ms=period_ms // 2, steps=20, easing=self.tok.ease_out_quad)
            job2 = self.app.after(period_ms // 2, lambda: loop(1 - state))
            self._pulse_jobs[key] = job2

        loop(0)

    def stop_pulse(self, widget, prop: str):
        key = self._key(widget, prop, "pulse")
        if key in self._pulse_on:
            self._pulse_on.remove(key)
        job = self._pulse_jobs.pop(key, None)
        if job is not None:
            try: self.app.after_cancel(job)
            except: pass

    def hover(self, widget, *,
              enter=None, leave=None,
              duration_ms=None):
        duration_ms = duration_ms if duration_ms is not None else self.tok.fast
        enter = enter or {}
        leave = leave or {}

        def on_enter(_):
            for prop, val in enter.items():
                self.color(widget, prop, val, duration_ms=duration_ms)

        def on_leave(_):
            for prop, val in leave.items():
                self.color(widget, prop, val, duration_ms=duration_ms)

        try:
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
        except:
            pass

def _toast(app, message, duration_ms=2500, toast_type="info"):
    """Modern animated toast notification"""
    try:
        toast = ctk.CTkToplevel(app)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.attributes("-alpha", 0.0)  # Start invisible
        
        app.update_idletasks()
        w, h = 320, 56
        x = app.winfo_x() + app.winfo_width() - w - 24
        y = app.winfo_y() + 50
        toast.geometry(f"{w}x{h}+{x}+{y}")
        
        # Color based on type
        colors = {
            "info": ("#0f172a", "#00f0ff", "#1a3a4a"),
            "success": ("#0f172a", "#10b981", "#1a3d32"),
            "error": ("#0f172a", "#ef4444", "#3d1a1a"),
            "warning": ("#0f172a", "#f59e0b", "#3d2e1a"),
        }
        bg, accent, border = colors.get(toast_type, colors["info"])
        
        frame = ctk.CTkFrame(
            toast, 
            corner_radius=14, 
            fg_color=bg, 
            border_width=1, 
            border_color=border
        )
        frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        # Accent bar
        accent_bar = ctk.CTkFrame(frame, width=4, height=40, corner_radius=2, fg_color=accent)
        accent_bar.place(x=12, rely=0.5, anchor="w")
        
        ctk.CTkLabel(
            frame, 
            text=message, 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#f8fafc"
        ).pack(expand=True, padx=(28, 16))
        
        # Fade in animation
        def fade_in(alpha=0.0):
            if alpha < 0.95:
                toast.attributes("-alpha", alpha)
                toast.after(16, lambda: fade_in(alpha + 0.1))
            else:
                toast.attributes("-alpha", 0.95)
        
        # Fade out animation
        def fade_out(alpha=0.95):
            if alpha > 0.05:
                toast.attributes("-alpha", alpha)
                toast.after(16, lambda: fade_out(alpha - 0.08))
            else:
                toast.destroy()
        
        fade_in()
        toast.after(duration_ms, fade_out)
        
    except: pass

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # UI Motion Engine
        self.motion = Motion(self, MotionTokens)

        # Window setup
        self.title("⚡ ChatGPT Auto Tools")
        self.geometry("1100x920")
        self.minsize(900, 750)
        
        # Configure window background
        self.configure(fg_color="#080b12")
        
        # Premium Fonts - More distinctive choices
        self.font_title = ctk.CTkFont(family="Segoe UI", size=36, weight="bold")
        self.font_subtitle = ctk.CTkFont(family="Segoe UI Light", size=15, weight="normal")
        self.font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="normal")
        self.font_button = ctk.CTkFont(family="Segoe UI Semibold", size=14, weight="bold")
        self.font_mono = ctk.CTkFont(family="Cascadia Code", size=11)
        self.font_stats = ctk.CTkFont(family="Segoe UI", size=28, weight="bold")
        self.font_stats_label = ctk.CTkFont(family="Segoe UI", size=11, weight="normal")
        
        # 🎨 VIBRANT Cyberpunk Color Palette
        self.colors = {
            # Primary Accents - Electric Neons
            "accent_primary": "#00f0ff",    # Electric Cyan
            "accent_secondary": "#ff00aa",  # Hot Pink/Magenta
            "accent_tertiary": "#7c3aed",   # Purple
            "accent_green": "#00ff9f",      # Neon Mint
            "accent_orange": "#ff6b35",     # Coral Orange
            "accent_yellow": "#ffd600",     # Electric Yellow
            
            # Backgrounds - Deep Space
            "bg_base": "#080b12",           # Near Black
            "bg_dark": "#0c1018",           # Dark Navy
            "bg_card": "#111827",           # Card Background
            "bg_card_hover": "#1a2332",     # Card Hover
            "bg_elevated": "#1e293b",       # Elevated surfaces
            
            # Glass Effects
            "glass_bg": "#0f172a",          # Glassmorphism base
            "glass_border": "#334155",      # Glass border
            
            # Gradients (start colors)
            "gradient_cyan": "#00f0ff",
            "gradient_purple": "#a855f7",
            "gradient_pink": "#ec4899",
            
            # Text
            "text_primary": "#f8fafc",      # Bright white
            "text_secondary": "#94a3b8",    # Slate gray
            "text_muted": "#64748b",        # Muted
            
            # Status Colors
            "success": "#10b981",           # Emerald
            "success_glow": "#34d399",      # Light emerald
            "error": "#ef4444",             # Red
            "error_glow": "#f87171",        # Light red
            "warning": "#f59e0b",           # Amber
            "info": "#3b82f6",              # Blue
            
            # Borders
            "border_subtle": "#1e293b",
            "border_glow": "#334155",
            "border_accent": "#1a3a4a",   # Cyan tint (muted)
        }

        # Grid layout (2 cols: Controls | Status)
        self.grid_columnconfigure(0, weight=3)  # Left side (Controls) - larger
        self.grid_columnconfigure(1, weight=2)  # Right side (Status)
        self.grid_rowconfigure(0, weight=0)     # Header
        self.grid_rowconfigure(1, weight=0)     # Main Content
        self.grid_rowconfigure(2, weight=1)     # Logs
        
        # --- ANIMATED HEADER WITH GRADIENT EFFECT ---
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent", height=100)
        self.header_frame.grid(row=0, column=0, columnspan=2, padx=32, pady=(28, 16), sticky="ew")
        self.header_frame.grid_columnconfigure(1, weight=1)
        
        # Logo/Icon container with glow
        self.logo_frame = ctk.CTkFrame(
            self.header_frame, 
            width=56, height=56, 
            corner_radius=14,
            fg_color=self.colors["bg_card"],
            border_width=1,
            border_color=self.colors["accent_primary"]
        )
        self.logo_frame.grid(row=0, column=0, rowspan=2, padx=(0, 16), sticky="w")
        self.logo_frame.grid_propagate(False)
        
        self.logo_icon = ctk.CTkLabel(
            self.logo_frame, 
            text="⚡", 
            font=ctk.CTkFont(size=28),
            text_color=self.colors["accent_primary"]
        )
        self.logo_icon.place(relx=0.5, rely=0.5, anchor="center")
        
        # Title with gradient-like effect (we'll animate colors)
        self.header_label = ctk.CTkLabel(
            self.header_frame, 
            text="ChatGPT Auto Tools", 
            font=self.font_title, 
            text_color=self.colors["accent_primary"]
        )
        self.header_label.grid(row=0, column=1, sticky="sw", pady=(0, 0))
        
        # Subtitle with typing effect
        self.subtitle_label = ctk.CTkLabel(
            self.header_frame, 
            text="✨ Premium Automation Dashboard", 
            font=self.font_subtitle, 
            text_color=self.colors["text_secondary"]
        )
        self.subtitle_label.grid(row=1, column=1, sticky="nw", pady=(2, 0))
        
        # Version badge
        self.version_badge = ctk.CTkLabel(
            self.header_frame,
            text="v2.0",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            fg_color=self.colors["accent_tertiary"],
            corner_radius=6,
            text_color="white",
            width=40, height=20
        )
        self.version_badge.grid(row=0, column=2, sticky="ne", padx=(12, 0))
        
        # Separator line with gradient effect
        self.header_separator = ctk.CTkFrame(
            self, 
            height=2, 
            fg_color=self.colors["border_subtle"]
        )
        self.header_separator.grid(row=0, column=0, columnspan=2, padx=32, pady=(90, 0), sticky="sew")
        
        # Animate header title through vibrant colors
        self._header_colors = ["#00f0ff", "#00ff9f", "#ff00aa", "#a855f7", "#ffd600", "#ff6b35"]
        self._header_idx = 0
        self._animate_header()

        # Animate logo glow
        self._animate_logo_glow()

        # --- LEFT COLUMN: CONTROLS (ENHANCED TABVIEW) ---
        self.controls_container = ctk.CTkFrame(
            self,
            fg_color=self.colors["glass_bg"],
            corner_radius=20,
            border_width=1,
            border_color=self.colors["glass_border"]
        )
        self.controls_container.grid(row=1, column=0, padx=(32, 12), pady=12, sticky="nsew")
        
        self.tabview = ctk.CTkTabview(
            self.controls_container, 
            width=600,
            fg_color="transparent",
            segmented_button_fg_color=self.colors["bg_card"],
            segmented_button_selected_color=self.colors["accent_primary"],
            segmented_button_selected_hover_color=self.colors["accent_primary"],
            segmented_button_unselected_color=self.colors["bg_card"],
            segmented_button_unselected_hover_color=self.colors["bg_card_hover"],
            text_color=self.colors["text_primary"],
            corner_radius=16
        )
        self.tabview.pack(fill="both", expand=True, padx=16, pady=16)
        self.tabview.add("🚀 Registration")
        self.tabview.add("🔐 MFA Automation")
        self.tabview.add("💳 Checkout Capture")
        
        # Enhanced tab font
        self.tabview._segmented_button.configure(
            font=ctk.CTkFont(family="Segoe UI Semibold", size=13, weight="bold"),
            corner_radius=12
        )

        self.setup_registration_tab()
        self.setup_mfa_tab()
        self.setup_checkout_tab()
        
        # --- RIGHT COLUMN: STATUS PANEL (GLASSMORPHISM STYLE) ---
        self.status_frame = ctk.CTkFrame(
            self, 
            fg_color=self.colors["glass_bg"],
            corner_radius=20,
            border_width=1,
            border_color=self.colors["glass_border"]
        )
        self.status_frame.grid(row=1, column=1, padx=(12, 32), pady=12, sticky="nsew")
        self.status_frame.grid_columnconfigure(0, weight=1)
        
        # ═══ STATUS SECTION ═══
        self.status_header = ctk.CTkFrame(self.status_frame, fg_color="transparent")
        self.status_header.grid(row=0, column=0, padx=20, pady=(20, 12), sticky="ew")
        self.status_header.grid_columnconfigure(1, weight=1)
        
        # Pulsing status dot
        self.status_dot = PulsingDot(self.status_header, color=self.colors["text_muted"], size=10)
        self.status_dot.grid(row=0, column=0, padx=(0, 10), sticky="w")
        
        self.status_label = ctk.CTkLabel(
            self.status_header, 
            text="SYSTEM STATUS", 
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold", slant="roman"),
            text_color=self.colors["text_muted"]
        )
        self.status_label.grid(row=0, column=1, sticky="w")
        
        # Status indicator pill
        self.status_indicator = ctk.CTkButton(
            self.status_frame, 
            text="● IDLE", 
            fg_color=self.colors["bg_elevated"], 
            state="disabled", 
            width=140, height=36, 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), 
            corner_radius=18,
            border_width=1,
            border_color=self.colors["border_subtle"]
        )
        self.status_indicator.grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")
        
        # ═══ STATS CARDS ═══
        self.stats_container = ctk.CTkFrame(self.status_frame, fg_color="transparent")
        self.stats_container.grid(row=2, column=0, padx=20, pady=8, sticky="ew")
        self.stats_container.grid_columnconfigure(0, weight=1)
        self.stats_container.grid_columnconfigure(1, weight=1)
        
        # Success Card
        self.success_card = ctk.CTkFrame(
            self.stats_container,
            fg_color=self.colors["bg_card"],
            corner_radius=14,
            border_width=1,
            border_color="#1a3d32"  # Green tint
        )
        self.success_card.grid(row=0, column=0, padx=(0, 6), pady=4, sticky="ew")
        
        self.success_icon = ctk.CTkLabel(
            self.success_card,
            text="✓",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=self.colors["success"]
        )
        self.success_icon.pack(pady=(12, 4))
        
        self.success_count_label = ctk.CTkLabel(
            self.success_card,
            text="0",
            font=self.font_stats,
            text_color=self.colors["success_glow"]
        )
        self.success_count_label.pack(pady=(0, 2))
        
        self.success_text_label = ctk.CTkLabel(
            self.success_card,
            text="SUCCESS",
            font=self.font_stats_label,
            text_color=self.colors["text_muted"]
        )
        self.success_text_label.pack(pady=(0, 12))
        
        # Failed Card
        self.fail_card = ctk.CTkFrame(
            self.stats_container,
            fg_color=self.colors["bg_card"],
            corner_radius=14,
            border_width=1,
            border_color="#3d1a1a"  # Red tint
        )
        self.fail_card.grid(row=0, column=1, padx=(6, 0), pady=4, sticky="ew")
        
        self.fail_icon = ctk.CTkLabel(
            self.fail_card,
            text="✗",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=self.colors["error"]
        )
        self.fail_icon.pack(pady=(12, 4))
        
        self.fail_count_label = ctk.CTkLabel(
            self.fail_card,
            text="0",
            font=self.font_stats,
            text_color=self.colors["error_glow"]
        )
        self.fail_count_label.pack(pady=(0, 2))
        
        self.fail_text_label = ctk.CTkLabel(
            self.fail_card,
            text="FAILED",
            font=self.font_stats_label,
            text_color=self.colors["text_muted"]
        )
        self.fail_text_label.pack(pady=(0, 12))
        
        # ═══ PROGRESS SECTION ═══
        self.progress_label = ctk.CTkLabel(
            self.status_frame, 
            text="PROGRESS",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=self.colors["text_muted"]
        )
        self.progress_label.grid(row=3, column=0, padx=20, pady=(20, 8), sticky="w")
        
        # Custom gradient-like progress bar container
        self.progress_container = ctk.CTkFrame(
            self.status_frame,
            fg_color=self.colors["bg_card"],
            corner_radius=8,
            height=20
        )
        self.progress_container.grid(row=4, column=0, padx=20, pady=(0, 8), sticky="ew")
        
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_container, 
            orientation="horizontal", 
            mode="determinate",
            height=6,
            progress_color=self.colors["accent_primary"],
            fg_color=self.colors["bg_elevated"],
            corner_radius=3
        )
        self.progress_bar.pack(fill="x", padx=8, pady=7)
        self.progress_bar.set(0)
        
        # Custom 60fps animation state
        self._progress_phase = 0.0
        self._progress_job = None

        # Progress percentage
        self.progress_percent = ctk.CTkLabel(
            self.status_frame,
            text="0%",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.colors["accent_primary"]
        )
        self.progress_percent.grid(row=5, column=0, padx=20, pady=(0, 12), sticky="e")

        # ═══ ACTIVITY LOG (Mini) ═══
        self.activity_label = ctk.CTkLabel(
            self.status_frame,
            text="ACTIVITY",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=self.colors["text_muted"]
        )
        self.activity_label.grid(row=6, column=0, padx=20, pady=(8, 8), sticky="w")
        
        self.info_box = ctk.CTkTextbox(
            self.status_frame, 
            height=80, 
            font=self.font_mono, 
            fg_color=self.colors["bg_card"], 
            text_color=self.colors["accent_primary"], 
            corner_radius=10,
            border_width=1,
            border_color=self.colors["border_subtle"]
        )
        self.info_box.grid(row=7, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.info_box.insert("0.0", "⚡ Ready to start automation...")
        self.info_box.configure(state="disabled")

        # --- LOG CONSOLE (TERMINAL STYLE) ---
        self.log_frame = ctk.CTkFrame(
            self,
            fg_color=self.colors["glass_bg"],
            corner_radius=20,
            border_width=1,
            border_color=self.colors["glass_border"]
        )
        self.log_frame.grid(row=2, column=0, columnspan=2, padx=24, pady=(8, 16), sticky="nsew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        
        # Terminal-style toolbar
        self.log_toolbar = ctk.CTkFrame(self.log_frame, fg_color="transparent", height=40)
        self.log_toolbar.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="ew")
        
        # Terminal window dots (macOS style)
        self.dots_frame = ctk.CTkFrame(self.log_toolbar, fg_color="transparent")
        self.dots_frame.pack(side="left", padx=(4, 12))
        
        for color in ["#ff5f57", "#febc2e", "#28c840"]:
            dot = ctk.CTkFrame(self.dots_frame, width=12, height=12, corner_radius=6, fg_color=color)
            dot.pack(side="left", padx=2)
        
        self.log_title = ctk.CTkLabel(
            self.log_toolbar, 
            text="SYSTEM CONSOLE", 
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.colors["text_muted"]
        )
        self.log_title.pack(side="left", padx=8)
        
        # Terminal blinking cursor indicator
        self.cursor_indicator = ctk.CTkLabel(
            self.log_toolbar,
            text="▊",
            font=ctk.CTkFont(size=12),
            text_color=self.colors["accent_primary"]
        )
        self.cursor_indicator.pack(side="left", padx=4)
        self._blink_cursor()
        
        # Toolbar buttons with modern styling
        btn_style = {
            "width": 36, "height": 28,
            "fg_color": self.colors["bg_card"],
            "hover_color": self.colors["bg_card_hover"],
            "corner_radius": 8,
            "font": ctk.CTkFont(size=14),
            "text_color": self.colors["text_secondary"],
            "border_width": 1,
            "border_color": self.colors["border_subtle"]
        }
        
        self.btn_clear = ctk.CTkButton(self.log_toolbar, text="🗑", command=self.clear_logs, **btn_style)
        self.btn_clear.pack(side="right", padx=3)
        
        self.btn_copy = ctk.CTkButton(self.log_toolbar, text="📋", command=self.copy_logs, **btn_style)
        self.btn_copy.pack(side="right", padx=3)
        
        self.btn_export = ctk.CTkButton(self.log_toolbar, text="💾", command=self.export_logs, **btn_style)
        self.btn_export.pack(side="right", padx=3)

        # Hover effects for toolbar buttons
        for btn in [self.btn_clear, self.btn_copy, self.btn_export]:
            self.motion.hover(
                btn, 
                enter={"fg_color": self.colors["bg_elevated"], "border_color": self.colors["accent_primary"]}, 
                leave={"fg_color": self.colors["bg_card"], "border_color": self.colors["border_subtle"]}, 
                duration_ms=120
            )
        
        # Log textbox with terminal aesthetics
        self.log_textbox = ctk.CTkTextbox(
            self.log_frame, 
            font=self.font_mono, 
            fg_color=self.colors["bg_base"],
            text_color="#e2e8f0",
            corner_radius=12,
            border_width=1,
            border_color=self.colors["border_subtle"],
            scrollbar_button_color=self.colors["bg_elevated"],
            scrollbar_button_hover_color=self.colors["accent_primary"]
        )
        self.log_textbox.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")
        
        # Configure Tags for Colored Logs (Neon Cyberpunk)
        self.log_textbox.tag_config("error", foreground="#f87171")      # Light red
        self.log_textbox.tag_config("success", foreground="#4ade80")    # Light green
        self.log_textbox.tag_config("warning", foreground="#fbbf24")    # Amber
        self.log_textbox.tag_config("info", foreground="#38bdf8")       # Sky blue
        self.log_textbox.tag_config("header", foreground="#c084fc")     # Purple
        self.log_textbox.tag_config("normal", foreground="#e2e8f0")     # Slate
        
        self.log_textbox.configure(state="disabled")

        # Redirect stdout
        sys.stdout = TextRedirector(self.log_textbox)
        sys.stderr = TextRedirector(self.log_textbox)
        self.running = False
        self.stop_event = threading.Event()
        self.log_buffer = []
        
        # ═══ ENTRANCE ANIMATIONS ═══
        self.after(100, self._play_entrance_animations)

    # --- SETUP TABS ---
    def setup_registration_tab(self):
        tab = self.tabview.tab("🚀 Registration")
        tab.grid_columnconfigure(0, weight=1)
        
        # ═══ SETTINGS CARD ═══
        settings_card = ctk.CTkFrame(
            tab, 
            fg_color=self.colors["bg_card"],
            corner_radius=16,
            border_width=1,
            border_color=self.colors["border_subtle"]
        )
        settings_card.pack(fill="x", padx=8, pady=(4, 8))
        
        # Card Header
        card_header = ctk.CTkFrame(settings_card, fg_color="transparent")
        card_header.pack(fill="x", padx=16, pady=(6, 4))
        
        ctk.CTkLabel(
            card_header, 
            text="⚙️  Configuration", 
            font=ctk.CTkFont(family="Segoe UI Semibold", size=14, weight="bold"),
            text_color=self.colors["text_primary"]
        ).pack(side="left")
        
        # Divider
        ctk.CTkFrame(settings_card, height=1, fg_color=self.colors["border_subtle"]).pack(fill="x", padx=16)
        
        # Mode Selection
        self.reg_mode_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        self.reg_mode_frame.pack(fill="x", padx=16, pady=6)
        
        mode_label = ctk.CTkLabel(
            self.reg_mode_frame, 
            text="Execution Mode", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        )
        mode_label.pack(side="left", padx=(0, 16))
        
        self.reg_mode_var = ctk.StringVar(value="Sequential")
        self.reg_mode_menu = ctk.CTkOptionMenu(
            self.reg_mode_frame, 
            values=["Sequential", "Multithread"], 
            variable=self.reg_mode_var, 
            command=self.toggle_reg_inputs, 
            width=180,
            height=36,
            font=self.font_label,
            dropdown_font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            button_color=self.colors["accent_tertiary"],
            button_hover_color=self.colors["accent_secondary"],
            dropdown_fg_color=self.colors["bg_card"],
            dropdown_hover_color=self.colors["bg_card_hover"],
            corner_radius=10
        )
        self.reg_mode_menu.pack(side="right")
        
        # Count Input Row (will be modified for Multithread to include Threads)
        self.reg_count_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        self.reg_count_frame.pack(fill="x", padx=16, pady=(0, 6))
        
        # Left side: Total Accounts
        self.reg_count_left = ctk.CTkFrame(self.reg_count_frame, fg_color="transparent")
        self.reg_count_left.pack(side="left", fill="x", expand=True)
        
        self.reg_count_label = ctk.CTkLabel(
            self.reg_count_left, 
            text="Account Count", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        )
        self.reg_count_label.pack(side="left", padx=(0, 8))
        
        self.reg_count_entry = ctk.CTkEntry(
            self.reg_count_left, 
            placeholder_text="1", 
            width=70,
            height=36,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            corner_radius=10
        )
        self.reg_count_entry.insert(0, "1")
        self.reg_count_entry.pack(side="left", padx=(0, 0))
        
        # Right side: Threads (only visible in Multithread mode)
        self.reg_threads_frame = ctk.CTkFrame(self.reg_count_frame, fg_color="transparent")
        # Hidden by default (Sequential mode)
        
        self.reg_threads_label = ctk.CTkLabel(
            self.reg_threads_frame, 
            text="Threads", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        )
        self.reg_threads_label.pack(side="left", padx=(0, 8))
        
        self.reg_threads_entry = ctk.CTkEntry(
            self.reg_threads_frame, 
            placeholder_text="2", 
            width=70,
            height=36,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            corner_radius=10
        )
        self.reg_threads_entry.insert(0, "2")
        self.reg_threads_entry.pack(side="left")
        
        # Thread Delay Input (only visible in Multithread mode)
        self.reg_delay_frame = ctk.CTkFrame(settings_card, fg_color="transparent")
        self.reg_delay_frame.pack(fill="x", padx=16, pady=(0, 10))
        self.reg_delay_frame.pack_forget()  # Hidden by default (Sequential mode)
        
        self.reg_delay_label = ctk.CTkLabel(
            self.reg_delay_frame, 
            text="Delay Between Browsers", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        )
        self.reg_delay_label.pack(side="left", padx=(0, 8))
        
        # Delay unit badge
        self.delay_unit_badge = ctk.CTkLabel(
            self.reg_delay_frame,
            text="seconds",
            font=ctk.CTkFont(size=10),
            text_color=self.colors["text_muted"]
        )
        self.delay_unit_badge.pack(side="left")
        
        self.reg_delay_entry = ctk.CTkEntry(
            self.reg_delay_frame, 
            placeholder_text="2", 
            width=80,
            height=36,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            corner_radius=10
        )
        self.reg_delay_entry.insert(0, "2")
        self.reg_delay_entry.pack(side="right")
        
        # Info tooltip for delay
        self.delay_info = ctk.CTkLabel(
            self.reg_delay_frame,
            text="ℹ️",
            font=ctk.CTkFont(size=14),
            text_color=self.colors["text_muted"]
        )
        self.delay_info.pack(side="right", padx=(0, 8))
        
        # ═══ ADVANCED OPTIONS CARD ═══
        adv_card = ctk.CTkFrame(
            tab, 
            fg_color=self.colors["bg_card"],
            corner_radius=16,
            border_width=1,
            border_color=self.colors["border_subtle"]
        )
        adv_card.pack(fill="x", padx=8, pady=(0, 8))
        
        # Advanced Header
        adv_header = ctk.CTkFrame(adv_card, fg_color="transparent")
        adv_header.pack(fill="x", padx=16, pady=(6, 4))
        
        ctk.CTkLabel(
            adv_header, 
            text="🎯  Advanced Options", 
            font=ctk.CTkFont(family="Segoe UI Semibold", size=14, weight="bold"),
            text_color=self.colors["text_primary"]
        ).pack(side="left")
        
        ctk.CTkFrame(adv_card, height=1, fg_color=self.colors["border_subtle"]).pack(fill="x", padx=16)
        
        # Network Mode selector
        self.reg_network_frame = ctk.CTkFrame(adv_card, fg_color="transparent")
        self.reg_network_frame.pack(fill="x", padx=16, pady=(6, 4))
        
        ctk.CTkLabel(
            self.reg_network_frame, 
            text="🌐  Network Mode:", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        ).pack(side="left", padx=(0, 16))
        
        self.reg_network_var = ctk.StringVar(value="Fast")
        self.reg_network_menu = ctk.CTkOptionMenu(
            self.reg_network_frame, 
            values=["Fast", "VPN/Slow"], 
            variable=self.reg_network_var, 
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            button_color=self.colors["accent_tertiary"],
            button_hover_color=self.colors["accent_primary"],
            dropdown_fg_color=self.colors["bg_elevated"],
            dropdown_hover_color=self.colors["bg_card_hover"],
            dropdown_text_color=self.colors["text_primary"],
            text_color=self.colors["text_primary"],
            width=120,
            height=32,
            corner_radius=8
        )
        self.reg_network_menu.pack(side="left")
        
        # Network mode info
        self.reg_network_info = ctk.CTkLabel(
            self.reg_network_frame,
            text="⚡ Fast: Stable network | 🐢 VPN/Slow: Unstable/VPN",
            font=ctk.CTkFont(size=11),
            text_color=self.colors["text_muted"]
        )
        self.reg_network_info.pack(side="right")
        
        # Email Mode selector
        self.reg_email_mode_frame = ctk.CTkFrame(adv_card, fg_color="transparent")
        self.reg_email_mode_frame.pack(fill="x", padx=16, pady=(4, 4))
        
        ctk.CTkLabel(
            self.reg_email_mode_frame, 
            text="📧  Email Mode:", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        ).pack(side="left", padx=(0, 16))
        
        self.reg_email_mode_var = ctk.StringVar(value="TinyHost")
        self.reg_email_mode_menu = ctk.CTkOptionMenu(
            self.reg_email_mode_frame, 
            values=["TinyHost", "OAuth2"], 
            variable=self.reg_email_mode_var,
            command=self.on_email_mode_change,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            button_color=self.colors["accent_tertiary"],
            button_hover_color=self.colors["accent_primary"],
            dropdown_fg_color=self.colors["bg_elevated"],
            dropdown_hover_color=self.colors["bg_card_hover"],
            dropdown_text_color=self.colors["text_primary"],
            text_color=self.colors["text_primary"],
            width=120,
            height=32,
            corner_radius=8
        )
        self.reg_email_mode_menu.pack(side="left")
        
        # OAuth2 status label (shows loaded accounts count)
        self.reg_oauth2_status = ctk.CTkLabel(
            self.reg_email_mode_frame,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=self.colors["text_muted"]
        )
        self.reg_oauth2_status.pack(side="left", padx=(12, 0))
        
        # OAuth2 Refresh button
        self.reg_oauth2_refresh = ctk.CTkButton(
            self.reg_email_mode_frame,
            text="🔄",
            width=36,
            height=32,
            font=ctk.CTkFont(size=18),
            fg_color=self.colors["bg_elevated"],
            hover_color=self.colors["bg_card_hover"],
            corner_radius=6,
            command=self.refresh_oauth2_accounts
        )
        self.reg_oauth2_refresh.pack(side="left", padx=(8, 0))
        self.reg_oauth2_refresh.pack_forget()  # Hidden by default
        
        # Password Config

        self.reg_password_frame = ctk.CTkFrame(adv_card, fg_color="transparent")
        self.reg_password_frame.pack(fill="x", padx=16, pady=(4, 6))
        
        ctk.CTkLabel(
            self.reg_password_frame, 
            text="🔑  Password:", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        ).pack(side="left", padx=(0, 12))
        
        self.reg_password_var = ctk.StringVar(value=DEFAULT_PASSWORD)
        self.reg_password_entry = ctk.CTkEntry(
            self.reg_password_frame,
            textvariable=self.reg_password_var,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            text_color=self.colors["text_primary"],
            width=180,
            height=32,
            corner_radius=8,
            show="•"  # Hide password by default
        )
        self.reg_password_entry.pack(side="left")
        
        # Show/Hide password button
        self.password_visible = False
        self.reg_password_toggle = ctk.CTkButton(
            self.reg_password_frame,
            text="👁",
            width=32,
            height=32,
            fg_color=self.colors["bg_elevated"],
            hover_color=self.colors["bg_card_hover"],
            corner_radius=8,
            command=self.toggle_password_visibility
        )
        self.reg_password_toggle.pack(side="left", padx=(4, 0))
        
        # Save password button
        self.reg_password_save = ctk.CTkButton(
            self.reg_password_frame,
            text="💾 Save",
            width=60,
            height=32,
            fg_color=self.colors["accent_tertiary"],
            hover_color=self.colors["accent_primary"],
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            command=self.save_password_to_file
        )
        self.reg_password_save.pack(side="left", padx=(8, 0))
        
        # Password info
        self.reg_password_info = ctk.CTkLabel(
            self.reg_password_frame,
            text="Auto-saved to code",
            font=ctk.CTkFont(size=10),
            text_color=self.colors["text_muted"]
        )
        self.reg_password_info.pack(side="right")
        
        # Checkout Switch
        self.reg_adv_frame = ctk.CTkFrame(adv_card, fg_color="transparent")
        self.reg_adv_frame.pack(fill="x", padx=16, pady=(6, 10))
        
        self.reg_checkout_var = ctk.BooleanVar(value=False)
        self.reg_checkout_switch = ctk.CTkSwitch(
            self.reg_adv_frame, 
            text="  Capture Checkout Link",
            variable=self.reg_checkout_var,
            font=self.font_label,
            text_color=self.colors["text_secondary"],
            fg_color=self.colors["bg_elevated"],
            progress_color=self.colors["accent_primary"],
            button_color=self.colors["text_primary"],
            button_hover_color=self.colors["accent_primary"],
            command=self.toggle_checkout_type
        )
        self.reg_checkout_switch.pack(side="left")
        
        # Checkout Type dropdown frame
        self.reg_checkout_type_frame = ctk.CTkFrame(self.reg_adv_frame, fg_color="transparent")
        self.reg_checkout_type_frame.pack(side="left", padx=(20, 0))
        
        ctk.CTkLabel(
            self.reg_checkout_type_frame,
            text="Type:",
            font=self.font_label,
            text_color=self.colors["text_muted"]
        ).pack(side="left", padx=(0, 8))
        
        self.reg_checkout_type_var = ctk.StringVar(value="Plus")
        self.reg_checkout_type_dropdown = ctk.CTkOptionMenu(
            self.reg_checkout_type_frame,
            values=["Plus", "Business", "Both"],
            variable=self.reg_checkout_type_var,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            button_color=self.colors["accent_secondary"],
            button_hover_color=self.colors["accent_primary"],
            dropdown_fg_color=self.colors["bg_elevated"],
            dropdown_hover_color=self.colors["bg_card_hover"],
            dropdown_text_color=self.colors["text_primary"],
            text_color=self.colors["text_primary"],
            width=100,
            height=32,
            corner_radius=8
        )
        self.reg_checkout_type_dropdown.pack(side="left")
        self.reg_checkout_type_frame.pack_forget()  # Hide initially
        
        # Info tooltip
        ctk.CTkLabel(
            self.reg_adv_frame,
            text="💳",
            font=ctk.CTkFont(size=16),
            text_color=self.colors["text_muted"]
        ).pack(side="right")
        
        # ═══ ACTION BUTTONS ═══
        self.reg_btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        self.reg_btn_frame.pack(fill="x", padx=8, pady=(8, 4))
        
        # Start Button with glow effect
        self.reg_start_btn = GlowButton(
            self.reg_btn_frame, 
            text="▶  START REGISTRATION", 
            command=self.start_registration_thread, 
            fg_color=self.colors["accent_green"], 
            hover_color="#00dd88",
            text_color="#0a0a0a",
            height=52, 
            corner_radius=14,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            glow_color=self.colors["accent_green"]
        )
        self.reg_start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        # Stop Button
        self.reg_stop_btn = ctk.CTkButton(
            self.reg_btn_frame, 
            text="⏹", 
            command=self.stop_process, 
            fg_color=self.colors["bg_card"], 
            border_width=2, 
            border_color=self.colors["error"], 
            text_color=self.colors["error"], 
            hover_color="#2a1515", 
            width=52,
            height=52, 
            corner_radius=14,
            font=ctk.CTkFont(size=18), 
            state="disabled"
        )
        self.reg_stop_btn.pack(side="right")

        # Enhanced Hovers
        self.motion.hover(
            self.reg_start_btn, 
            enter={"fg_color": "#33ffaa"}, 
            leave={"fg_color": self.colors["accent_green"]}, 
            duration_ms=150
        )

    def setup_mfa_tab(self):
        tab = self.tabview.tab("🔐 MFA Automation")
        tab.grid_columnconfigure(0, weight=1)
        
        # ═══ MFA SETTINGS CARD ═══
        mfa_settings_card = ctk.CTkFrame(
            tab, 
            fg_color=self.colors["bg_card"],
            corner_radius=16,
            border_width=1,
            border_color=self.colors["border_subtle"]
        )
        mfa_settings_card.pack(fill="x", padx=8, pady=(4, 8))
        
        # Card Header
        mfa_header = ctk.CTkFrame(mfa_settings_card, fg_color="transparent")
        mfa_header.pack(fill="x", padx=16, pady=(10, 8))
        
        ctk.CTkLabel(
            mfa_header, 
            text="🛡️  MFA Configuration", 
            font=ctk.CTkFont(family="Segoe UI Semibold", size=14, weight="bold"),
            text_color=self.colors["text_primary"]
        ).pack(side="left")
        
        # Badge
        ctk.CTkLabel(
            mfa_header,
            text="2FA",
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color=self.colors["accent_secondary"],
            corner_radius=6,
            text_color="white",
            width=36, height=20
        ).pack(side="right")
        
        # Divider
        ctk.CTkFrame(mfa_settings_card, height=1, fg_color=self.colors["border_subtle"]).pack(fill="x", padx=16)
        
        # Mode Selection
        self.mfa_mode_frame = ctk.CTkFrame(mfa_settings_card, fg_color="transparent")
        self.mfa_mode_frame.pack(fill="x", padx=16, pady=10)
        
        ctk.CTkLabel(
            self.mfa_mode_frame, 
            text="Execution Mode", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        ).pack(side="left", padx=(0, 16))
        
        self.mfa_mode_var = ctk.StringVar(value="Sequential")
        self.mfa_mode_menu = ctk.CTkOptionMenu(
            self.mfa_mode_frame, 
            values=["Sequential", "Multithread"], 
            variable=self.mfa_mode_var, 
            command=self.toggle_mfa_inputs, 
            width=180,
            height=36,
            font=self.font_label,
            dropdown_font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            button_color=self.colors["accent_tertiary"],
            button_hover_color=self.colors["accent_secondary"],
            dropdown_fg_color=self.colors["bg_card"],
            dropdown_hover_color=self.colors["bg_card_hover"],
            corner_radius=10
        )
        self.mfa_mode_menu.pack(side="right")
        
        # Threads Input
        self.mfa_threads_frame = ctk.CTkFrame(mfa_settings_card, fg_color="transparent")
        self.mfa_threads_frame.pack(fill="x", padx=16, pady=(0, 10))
        
        self.mfa_threads_label = ctk.CTkLabel(
            self.mfa_threads_frame, 
            text="Thread Count", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        )
        self.mfa_threads_label.pack(side="left", padx=(0, 16))
        
        # Thread limit badge
        self.thread_limit_badge = ctk.CTkLabel(
            self.mfa_threads_frame,
            text="max 5",
            font=ctk.CTkFont(size=9),
            text_color=self.colors["text_muted"]
        )
        self.thread_limit_badge.pack(side="left")
        
        self.mfa_threads_entry = ctk.CTkEntry(
            self.mfa_threads_frame, 
            placeholder_text="1", 
            width=180,
            height=36,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            corner_radius=10
        )
        self.mfa_threads_entry.insert(0, "1")
        self.mfa_threads_entry.pack(side="right")
        
        # Thread Delay Input (only visible in Multithread mode)
        self.mfa_delay_frame = ctk.CTkFrame(mfa_settings_card, fg_color="transparent")
        self.mfa_delay_frame.pack(fill="x", padx=16, pady=(0, 10))
        self.mfa_delay_frame.pack_forget()  # Hidden by default (Sequential mode)
        
        self.mfa_delay_label = ctk.CTkLabel(
            self.mfa_delay_frame, 
            text="Delay Between Browsers", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        )
        self.mfa_delay_label.pack(side="left", padx=(0, 8))
        
        # Delay unit badge
        self.mfa_delay_unit_badge = ctk.CTkLabel(
            self.mfa_delay_frame,
            text="seconds",
            font=ctk.CTkFont(size=10),
            text_color=self.colors["text_muted"]
        )
        self.mfa_delay_unit_badge.pack(side="left")
        
        self.mfa_delay_entry = ctk.CTkEntry(
            self.mfa_delay_frame, 
            placeholder_text="2", 
            width=80,
            height=36,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            corner_radius=10
        )
        self.mfa_delay_entry.insert(0, "2")
        self.mfa_delay_entry.pack(side="right")
        
        # Info tooltip for delay
        self.mfa_delay_info = ctk.CTkLabel(
            self.mfa_delay_frame,
            text="ℹ️",
            font=ctk.CTkFont(size=14),
            text_color=self.colors["text_muted"]
        )
        self.mfa_delay_info.pack(side="right", padx=(0, 8))
        
        # ═══ INFO CARD ═══
        info_card = ctk.CTkFrame(
            tab, 
            fg_color=self.colors["bg_card"],
            corner_radius=16,
            border_width=1,
            border_color="#3b82f6"  # Blue tint
        )
        info_card.pack(fill="x", padx=8, pady=(0, 8))
        
        info_content = ctk.CTkFrame(info_card, fg_color="transparent")
        info_content.pack(fill="x", padx=16, pady=10)
        
        ctk.CTkLabel(
            info_content,
            text="ℹ️",
            font=ctk.CTkFont(size=18),
            text_color=self.colors["info"]
        ).pack(side="left", padx=(0, 12))
        
        ctk.CTkLabel(
            info_content,
            text="MFA will be enabled for accounts in chatgpt.xlsx\nwithout existing 2FA configuration.",
            font=ctk.CTkFont(size=12),
            text_color=self.colors["text_secondary"],
            justify="left"
        ).pack(side="left")
        
        # ═══ ACTION BUTTONS ═══
        self.mfa_btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        self.mfa_btn_frame.pack(fill="x", padx=8, pady=(8, 4))
        
        # Start Button
        self.mfa_start_btn = GlowButton(
            self.mfa_btn_frame, 
            text="🔐  START MFA AUTOMATION", 
            command=self.start_mfa_thread, 
            fg_color=self.colors["accent_primary"], 
            hover_color="#33ddff",
            text_color="#0a0a0a",
            height=52, 
            corner_radius=14,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            glow_color=self.colors["accent_primary"]
        )
        self.mfa_start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        # Stop Button
        self.mfa_stop_btn = ctk.CTkButton(
            self.mfa_btn_frame, 
            text="⏹", 
            command=self.stop_process, 
            fg_color=self.colors["bg_card"], 
            border_width=2, 
            border_color=self.colors["error"], 
            text_color=self.colors["error"], 
            hover_color="#2a1515", 
            width=52,
            height=52, 
            corner_radius=14,
            font=ctk.CTkFont(size=18), 
            state="disabled"
        )
        self.mfa_stop_btn.pack(side="right")

        # Enhanced Hovers
        self.motion.hover(
            self.mfa_start_btn, 
            enter={"fg_color": "#66f0ff"}, 
            leave={"fg_color": self.colors["accent_primary"]}, 
            duration_ms=150
        )
        
        # Initial State
        self.toggle_mfa_inputs("Sequential")

    def setup_checkout_tab(self):
        """Setup the Checkout Capture tab with account selection table"""
        tab = self.tabview.tab("💳 Checkout Capture")
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        
        # ═══ CONTROLS CARD ═══
        controls_card = ctk.CTkFrame(
            tab, 
            fg_color=self.colors["bg_card"],
            corner_radius=16,
            border_width=1,
            border_color=self.colors["border_subtle"]
        )
        controls_card.grid(row=0, column=0, padx=8, pady=(4, 8), sticky="ew")
        
        # Header row
        header_frame = ctk.CTkFrame(controls_card, fg_color="transparent")
        header_frame.pack(fill="x", padx=16, pady=(10, 8))
        
        ctk.CTkLabel(
            header_frame, 
            text="💳  Checkout Link Capture", 
            font=ctk.CTkFont(family="Segoe UI Semibold", size=14, weight="bold"),
            text_color=self.colors["text_primary"]
        ).pack(side="left")
        
        # Refresh button
        self.checkout_refresh_btn = ctk.CTkButton(
            header_frame,
            text="🔄 Refresh",
            width=80,
            height=28,
            fg_color=self.colors["bg_elevated"],
            hover_color=self.colors["bg_card_hover"],
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            command=self.load_checkout_accounts
        )
        self.checkout_refresh_btn.pack(side="right")
        
        ctk.CTkFrame(controls_card, height=1, fg_color=self.colors["border_subtle"]).pack(fill="x", padx=16)
        
        # Options row
        options_frame = ctk.CTkFrame(controls_card, fg_color="transparent")
        options_frame.pack(fill="x", padx=16, pady=10)
        
        # Checkout type
        ctk.CTkLabel(
            options_frame, 
            text="Capture Type:", 
            font=self.font_label,
            text_color=self.colors["text_secondary"]
        ).pack(side="left", padx=(0, 8))
        
        self.checkout_type_var = ctk.StringVar(value="Plus")
        self.checkout_type_menu = ctk.CTkOptionMenu(
            options_frame, 
            values=["Plus", "Business", "Both"], 
            variable=self.checkout_type_var,
            width=100,
            height=32,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            button_color=self.colors["accent_tertiary"],
            button_hover_color=self.colors["accent_primary"],
            corner_radius=8
        )
        self.checkout_type_menu.pack(side="left", padx=(0, 16))
        
        # Select All / Deselect All
        self.checkout_select_all_btn = ctk.CTkButton(
            options_frame,
            text="☑ Select All",
            width=90,
            height=28,
            fg_color=self.colors["bg_elevated"],
            hover_color=self.colors["accent_primary"],
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            command=self.checkout_select_all
        )
        self.checkout_select_all_btn.pack(side="left", padx=(0, 4))
        
        self.checkout_deselect_all_btn = ctk.CTkButton(
            options_frame,
            text="☐ Deselect All",
            width=100,
            height=28,
            fg_color=self.colors["bg_elevated"],
            hover_color=self.colors["error"],
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            command=self.checkout_deselect_all
        )
        self.checkout_deselect_all_btn.pack(side="left")
        
        # Account count label
        self.checkout_count_label = ctk.CTkLabel(
            options_frame,
            text="0 accounts | 0 selected",
            font=ctk.CTkFont(size=11),
            text_color=self.colors["text_muted"]
        )
        self.checkout_count_label.pack(side="right")
        
        # ═══ MULTITHREAD OPTIONS (always visible, but disabled when < 2 selected) ═══
        self.checkout_mt_frame = ctk.CTkFrame(controls_card, fg_color="transparent", height=40)
        self.checkout_mt_frame.pack(fill="x", padx=16, pady=(0, 8))
        self.checkout_mt_frame.pack_propagate(False)  # Fixed height to prevent resize
        
        # Multithread enable switch
        self.checkout_mt_var = ctk.BooleanVar(value=False)
        self.checkout_mt_switch = ctk.CTkSwitch(
            self.checkout_mt_frame,
            text="  Multithread Mode",
            variable=self.checkout_mt_var,
            font=self.font_label,
            text_color=self.colors["text_muted"],  # Start muted
            fg_color=self.colors["bg_elevated"],
            progress_color=self.colors["accent_orange"],
            button_color=self.colors["text_primary"],
            button_hover_color=self.colors["accent_orange"],
            command=self.toggle_checkout_multithread,
            state="disabled"  # Start disabled
        )
        self.checkout_mt_switch.pack(side="left", padx=(0, 16), pady=4)
        
        # Thread count label
        self.checkout_threads_label = ctk.CTkLabel(
            self.checkout_mt_frame,
            text="Threads:",
            font=self.font_label,
            text_color=self.colors["text_muted"]
        )
        self.checkout_threads_label.pack(side="left", padx=(0, 8), pady=4)
        
        self.checkout_threads_entry = ctk.CTkEntry(
            self.checkout_mt_frame,
            placeholder_text="2",
            width=50,
            height=32,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            corner_radius=8,
            state="disabled"
        )
        self.checkout_threads_entry.insert(0, "2")
        self.checkout_threads_entry.pack(side="left", pady=4)
        
        # Delay label and entry
        self.checkout_delay_label = ctk.CTkLabel(
            self.checkout_mt_frame,
            text="Delay:",
            font=self.font_label,
            text_color=self.colors["text_muted"]
        )
        self.checkout_delay_label.pack(side="left", padx=(16, 8), pady=4)
        
        self.checkout_delay_entry = ctk.CTkEntry(
            self.checkout_mt_frame,
            placeholder_text="3",
            width=50,
            height=32,
            font=self.font_label,
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"],
            corner_radius=8,
            state="disabled"
        )
        self.checkout_delay_entry.insert(0, "3")  # Default 3s delay
        self.checkout_delay_entry.pack(side="left", pady=4)
        
        self.checkout_delay_unit = ctk.CTkLabel(
            self.checkout_mt_frame,
            text="s",
            font=self.font_label,
            text_color=self.colors["text_muted"]
        )
        self.checkout_delay_unit.pack(side="left", padx=(2, 0), pady=4)
        
        # Max threads info
        self.checkout_mt_info = ctk.CTkLabel(
            self.checkout_mt_frame,
            text="(select 2+)",
            font=ctk.CTkFont(size=10),
            text_color=self.colors["text_muted"]
        )
        self.checkout_mt_info.pack(side="left", padx=(8, 0), pady=4)
        
        # Recommendation badge
        self.checkout_mt_badge = ctk.CTkLabel(
            self.checkout_mt_frame,
            text="💡 Rec: 2 threads, 3s delay",
            font=ctk.CTkFont(size=10),
            text_color=self.colors["accent_yellow"]
        )
        self.checkout_mt_badge.pack(side="right")
        
        # ═══ ACCOUNTS TABLE ═══
        table_frame = ctk.CTkFrame(
            tab, 
            fg_color=self.colors["bg_card"],
            corner_radius=16,
            border_width=1,
            border_color=self.colors["border_subtle"]
        )
        table_frame.grid(row=1, column=0, padx=8, pady=(0, 8), sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        
        # Scrollable frame for accounts
        self.checkout_scroll = ctk.CTkScrollableFrame(
            table_frame,
            fg_color="transparent",
            corner_radius=0
        )
        self.checkout_scroll.pack(fill="both", expand=True, padx=8, pady=8)
        
        # Configure fixed column widths
        self.checkout_scroll.grid_columnconfigure(0, weight=0, minsize=40)   # Checkbox
        self.checkout_scroll.grid_columnconfigure(1, weight=1, minsize=200)  # Email (flexible)
        self.checkout_scroll.grid_columnconfigure(2, weight=0, minsize=50)   # Plus
        self.checkout_scroll.grid_columnconfigure(3, weight=0, minsize=70)   # Business
        self.checkout_scroll.grid_columnconfigure(4, weight=0, minsize=60)   # Sold
        
        # Header row for table
        header_row = ctk.CTkFrame(self.checkout_scroll, fg_color=self.colors["bg_elevated"], corner_radius=8, height=30)
        header_row.grid(row=0, column=0, columnspan=5, sticky="ew", pady=(0, 4))
        header_row.grid_columnconfigure(0, weight=0, minsize=40)
        header_row.grid_columnconfigure(1, weight=1, minsize=200)
        header_row.grid_columnconfigure(2, weight=0, minsize=50)
        header_row.grid_columnconfigure(3, weight=0, minsize=70)
        header_row.grid_columnconfigure(4, weight=0, minsize=60)
        
        ctk.CTkLabel(header_row, text="✓", font=ctk.CTkFont(size=11, weight="bold"), text_color=self.colors["text_secondary"]).grid(row=0, column=0, padx=8, sticky="w")
        ctk.CTkLabel(header_row, text="Email", font=ctk.CTkFont(size=11, weight="bold"), text_color=self.colors["text_secondary"], anchor="w").grid(row=0, column=1, padx=8, sticky="w")
        ctk.CTkLabel(header_row, text="Plus", font=ctk.CTkFont(size=11, weight="bold"), text_color=self.colors["text_secondary"]).grid(row=0, column=2, padx=4, sticky="w")
        ctk.CTkLabel(header_row, text="Business", font=ctk.CTkFont(size=11, weight="bold"), text_color=self.colors["text_secondary"]).grid(row=0, column=3, padx=4, sticky="w")
        ctk.CTkLabel(header_row, text="Sold", font=ctk.CTkFont(size=11, weight="bold"), text_color=self.colors["text_secondary"]).grid(row=0, column=4, padx=4, sticky="w")
        
        # Store for account checkboxes
        self.checkout_account_vars = []
        self.checkout_account_widgets = []
        
        # ═══ ACTION BUTTONS ═══
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=2, column=0, padx=8, pady=(0, 4), sticky="ew")
        
        self.checkout_start_btn = GlowButton(
            btn_frame, 
            text="💳  START CAPTURE", 
            command=self.start_checkout_capture_thread, 
            fg_color=self.colors["accent_orange"], 
            hover_color="#ff8855",
            text_color="#0a0a0a",
            height=48, 
            corner_radius=14,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            glow_color=self.colors["accent_orange"]
        )
        self.checkout_start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        self.checkout_stop_btn = ctk.CTkButton(
            btn_frame, 
            text="⏹", 
            command=self.stop_process, 
            fg_color="transparent",
            hover_color=self.colors["error"],
            text_color=self.colors["text_secondary"],
            width=52, height=48,
            corner_radius=14,
            font=ctk.CTkFont(size=18),
            border_width=2,
            border_color=self.colors["border_subtle"],
            state="disabled"
        )
        self.checkout_stop_btn.pack(side="right")
        
        # Load accounts on init
        self.after(500, self.load_checkout_accounts)
    
    def load_checkout_accounts(self):
        """Load accounts from Excel and display in table"""
        # Clear existing widgets safely to avoid TclError with CTkCheckBox
        # First, clear all variable traces by setting variables to False and unbinding
        for item in self.checkout_account_vars:
            try:
                var = item.get("var")
                if var:
                    var.set(False)  # Reset variable value
            except:
                pass
        
        # Clear variables list first to avoid callbacks during destroy
        self.checkout_account_vars.clear()
        
        # Now safely destroy widgets
        for widget in self.checkout_account_widgets:
            try:
                # For CTkCheckBox, we need to handle the variable trace issue
                if isinstance(widget, ctk.CTkCheckBox):
                    # Try to remove the variable callback before destroy
                    try:
                        if hasattr(widget, '_variable') and widget._variable:
                            if hasattr(widget, '_variable_callback_name') and widget._variable_callback_name:
                                try:
                                    widget._variable.trace_remove("write", widget._variable_callback_name)
                                except:
                                    pass
                                widget._variable_callback_name = None
                    except:
                        pass
                widget.destroy()
            except:
                pass
        self.checkout_account_widgets.clear()
        
        # Load accounts
        excel_file = "chatgpt.xlsx"
        if not os.path.exists(excel_file):
            self.checkout_count_label.configure(text="No chatgpt.xlsx found")
            return
        
        accounts = load_checkout_accounts(excel_file)
        
        if not accounts:
            self.checkout_count_label.configure(text="No accounts found in Excel")
            return
        
        # Create row for each account using grid layout
        for idx, account in enumerate(accounts):
            try:
                row_num = idx + 1
                is_sold = account.get("is_sold", False)
                
                # Create row frame - keep normal background, use border to highlight sold accounts
                if is_sold:
                    # Use bright orange border to highlight sold rows (keep normal background)
                    row_border_color = "#ff6b4a"  # Bright orange border
                    text_color = "#ff8c6b"  # Light orange text for better visibility
                    border_width = 2  # Thicker border for better visibility
                else:
                    row_border_color = None  # Don't set border_color when not needed
                    text_color = self.colors["text_primary"]
                    border_width = 0
                
                # Create frame - only set border_color if sold, keep normal background (transparent)
                frame_kwargs = {
                    "master": self.checkout_scroll,
                    "fg_color": "transparent",  # Keep same background as scrollable frame
                    "border_width": border_width,
                    "corner_radius": 6
                }
                if is_sold:
                    frame_kwargs["border_color"] = row_border_color
                
                row_frame = ctk.CTkFrame(**frame_kwargs)
                row_frame.grid(row=row_num, column=0, columnspan=5, sticky="ew", pady=1, padx=2)
                row_frame.grid_columnconfigure(0, weight=0, minsize=40)
                row_frame.grid_columnconfigure(1, weight=1, minsize=200)
                row_frame.grid_columnconfigure(2, weight=0, minsize=50)
                row_frame.grid_columnconfigure(3, weight=0, minsize=70)
                row_frame.grid_columnconfigure(4, weight=0, minsize=60)
                self.checkout_account_widgets.append(row_frame)
                
                # Checkbox
                var = ctk.BooleanVar(value=False)
                self.checkout_account_vars.append({
                    "var": var,
                    "account": account
                })
                
                cb = ctk.CTkCheckBox(
                    row_frame, 
                    text="",
                    variable=var,
                    width=30,
                    checkbox_width=18,
                    checkbox_height=18,
                    fg_color=self.colors["accent_primary"],
                    hover_color=self.colors["accent_secondary"],
                    command=self.update_checkout_selection_count,
                    state="disabled" if is_sold else "normal"
                )
                cb.grid(row=0, column=0, padx=(8, 4), pady=2, sticky="w")
                self.checkout_account_widgets.append(cb)
                
                # Email (truncate to fit)
                email_text = account["email"][:30] + "..." if len(account["email"]) > 30 else account["email"]
                email_label = ctk.CTkLabel(
                    row_frame, 
                    text=email_text,
                    font=ctk.CTkFont(size=11),
                    text_color=text_color,
                    anchor="w"
                )
                email_label.grid(row=0, column=1, padx=4, pady=2, sticky="w")
                self.checkout_account_widgets.append(email_label)
                
                # Plus status
                if account["plus_url"] == "no Plus offer":
                    plus_status = "⛔"
                    plus_color = self.colors["error"]
                elif account["plus_url"]:
                    plus_status = "✅"
                    plus_color = self.colors["accent_green"]
                else:
                    plus_status = "❌"
                    plus_color = self.colors["text_muted"]

                plus_label = ctk.CTkLabel(
                    row_frame, 
                    text=plus_status,
                    font=ctk.CTkFont(size=12),
                    text_color=plus_color
                )
                plus_label.grid(row=0, column=2, padx=4, pady=2, sticky="w")
                self.checkout_account_widgets.append(plus_label)
                
                # Business status
                business_status = "✅" if account["business_url"] else "❌"
                business_color = self.colors["accent_green"] if account["business_url"] else self.colors["text_muted"]
                business_label = ctk.CTkLabel(
                    row_frame, 
                    text=business_status,
                    font=ctk.CTkFont(size=12),
                    text_color=business_color
                )
                business_label.grid(row=0, column=3, padx=4, pady=2, sticky="w")
                self.checkout_account_widgets.append(business_label)
                
                # Sold status
                sold_status = "✅" if is_sold else "❌"
                sold_color = "#ff6b4a" if is_sold else self.colors["text_muted"]  # Bright orange for sold
                sold_label = ctk.CTkLabel(
                    row_frame, 
                    text=sold_status,
                    font=ctk.CTkFont(size=12),
                    text_color=sold_color
                )
                sold_label.grid(row=0, column=4, padx=4, pady=2, sticky="w")
                self.checkout_account_widgets.append(sold_label)
            except Exception as e:
                print(f"Error creating row for account {account.get('email', 'unknown')}: {e}")
                continue
        
        self.update_checkout_selection_count()
    
    def update_checkout_selection_count(self):
        """Update the selection count label and enable/disable multithread options"""
        # Count only non-sold accounts
        total = sum(1 for item in self.checkout_account_vars if not item["account"].get("is_sold", False))
        selected = sum(1 for item in self.checkout_account_vars if item["var"].get() and not item["account"].get("is_sold", False))
        self.checkout_count_label.configure(text=f"{total} accounts | {selected} selected")
        
        # Enable multithread options when 2+ accounts selected (no pack/unpack to avoid resize animation)
        if selected >= 2:
            # Enable switch and update colors
            self.checkout_mt_switch.configure(
                state="normal",
                text_color=self.colors["text_secondary"]
            )
            # Update max threads info
            max_threads = min(5, selected)
            self.checkout_mt_info.configure(
                text=f"(max {max_threads})",
                text_color=self.colors["accent_primary"]
            )
            self.checkout_threads_label.configure(text_color=self.colors["text_secondary"])
            self.checkout_delay_label.configure(text_color=self.colors["text_secondary"])
        else:
            # Disable switch and mute colors
            self.checkout_mt_switch.configure(
                state="disabled",
                text_color=self.colors["text_muted"]
            )
            self.checkout_mt_var.set(False)
            self.checkout_threads_entry.configure(state="disabled")
            self.checkout_delay_entry.configure(state="disabled")
            self.checkout_mt_info.configure(
                text="(select 2+ accounts)",
                text_color=self.colors["text_muted"]
            )
            self.checkout_threads_label.configure(text_color=self.colors["text_muted"])
            self.checkout_delay_label.configure(text_color=self.colors["text_muted"])
    
    def toggle_checkout_multithread(self):
        """Toggle multithread mode for checkout capture"""
        if self.checkout_mt_var.get():
            self.checkout_threads_entry.configure(state="normal")
            self.checkout_delay_entry.configure(state="normal")
        else:
            self.checkout_threads_entry.configure(state="disabled")
            self.checkout_delay_entry.configure(state="disabled")
    
    def checkout_select_all(self):
        """Select all non-sold accounts"""
        for item in self.checkout_account_vars:
            if not item["account"].get("is_sold", False):
                item["var"].set(True)
        self.update_checkout_selection_count()
    
    def checkout_deselect_all(self):
        """Deselect all accounts"""
        for item in self.checkout_account_vars:
            item["var"].set(False)
        self.update_checkout_selection_count()
    
    def start_checkout_capture_thread(self):
        """Start checkout capture in thread"""
        threading.Thread(target=self.run_checkout_capture).start()
    
    def run_checkout_capture(self):
        """Run checkout capture for selected accounts (sequential or multithread)"""
        # Get selected accounts (exclude sold accounts)
        selected = [item["account"] for item in self.checkout_account_vars 
                   if item["var"].get() and not item["account"].get("is_sold", False)]
        
        if not selected:
            _toast(self, "⚠️ No accounts selected!", toast_type="warning")
            return
        
        self.lock_ui(True)
        self.stop_event.clear()
        self.update_stats(0, 0)
        
        checkout_type = self.checkout_type_var.get()
        excel_file = "chatgpt.xlsx"
        
        # Check if multithread mode is enabled
        use_multithread = self.checkout_mt_var.get() and len(selected) >= 2
        
        if use_multithread:
            try:
                threads = int(self.checkout_threads_entry.get())
                max_allowed = min(5, len(selected))
                threads = max(1, min(threads, max_allowed))
            except:
                threads = 2
            
            # Get delay between browsers
            try:
                thread_delay = int(self.checkout_delay_entry.get())
                thread_delay = max(1, min(thread_delay, 10))  # 1-10s
            except:
                thread_delay = 3  # Default 3s
            
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting Checkout Capture | Type: {checkout_type} | Accounts: {len(selected)} | Threads: {threads} | Delay: {thread_delay}s")
            self.update_status("RUNNING", self.colors["accent_orange"], f"Capturing {len(selected)} accounts with {threads} threads...")
            
            success_count = 0
            fail_count = 0
            count_lock = threading.Lock()
            
            def run_checkout_worker(thread_id, account, account_idx, start_delay):
                nonlocal success_count, fail_count
                
                # Use pre-calculated start delay
                if start_delay > 0:
                    safe_print(thread_id, f"Waiting {start_delay}s before starting...", Colors.INFO, "⏳ ")
                    # Sleep in 0.5s intervals to check stop_event
                    for _ in range(int(start_delay * 2)):
                        if self.stop_event.is_set():
                            return False, account["email"]
                        time.sleep(0.5)
                
                if self.stop_event.is_set():
                    return False, account["email"]
                
                worker = CheckoutCaptureWorker(
                    thread_id=thread_id,
                    email=account["email"],
                    cookie_json=account["cookie"],
                    excel_file=excel_file,
                    row_index=account["row_index"],
                    checkout_type=checkout_type
                )
                worker.stop_event = self.stop_event
                
                result = worker.run()
                
                with count_lock:
                    if result:
                        nonlocal success_count
                        success_count += 1
                        print(f"✅ Captured: {account['email']}")
                    else:
                        nonlocal fail_count
                        fail_count += 1
                        print(f"❌ Failed: {account['email']}")
                    self.update_stats(success_count, fail_count)
                
                return result, account["email"]
            
            # Run with ThreadPoolExecutor
            with concurrent.futures.ThreadPoolExecutor(max_workers=threads) as executor:
                futures = []
                for idx, account in enumerate(selected):
                    if self.stop_event.is_set():
                        break
                    
                    # Slot index is (idx % threads) for staggered delay in each wave
                    slot_idx = idx % threads
                    wave_idx = idx // threads
                    
                    if wave_idx == 0:
                        # First wave: stagger by slot
                        start_delay = slot_idx * thread_delay
                    else:
                        # Subsequent waves: Base delay (5s) + stagger
                        start_delay = 5.0 + (slot_idx * thread_delay)
                    
                    thread_id = (idx % threads) + 1
                    future = executor.submit(run_checkout_worker, thread_id, account, idx, start_delay)
                    futures.append(future)
                
                # Wait for all futures
                for future in concurrent.futures.as_completed(futures):
                    if self.stop_event.is_set():
                        break
                    try:
                        future.result()
                    except Exception as e:
                        print(f"Error in worker: {e}")
        else:
            # Sequential mode
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting Checkout Capture | Type: {checkout_type} | Accounts: {len(selected)}")
            self.update_status("RUNNING", self.colors["accent_orange"], f"Capturing {len(selected)} accounts...")
            
            success_count = 0
            fail_count = 0
            
            for idx, account in enumerate(selected):
                if self.stop_event.is_set():
                    break
                
                self.update_status("RUNNING", self.colors["accent_orange"], f"Account {idx+1}/{len(selected)}: {account['email']}")
                
                worker = CheckoutCaptureWorker(
                    thread_id=1,
                    email=account["email"],
                    cookie_json=account["cookie"],
                    excel_file=excel_file,
                    row_index=account["row_index"],
                    checkout_type=checkout_type
                )
                worker.stop_event = self.stop_event
                
                result = worker.run()
                
                if result:
                    success_count += 1
                    print(f"✅ Captured: {account['email']}")
                else:
                    fail_count += 1
                    print(f"❌ Failed: {account['email']}")
                
                self.update_stats(success_count, fail_count)
        
        # Schedule UI updates on main thread to avoid Tkinter threading issues
        final_success = success_count
        final_fail = fail_count
        
        def finish_capture():
            # Refresh table safely
            self.load_checkout_accounts()
            
            # Final status
            self.lock_ui(False)
            if self.stop_event.is_set():
                self.update_status("STOPPED", self.colors["warning"], f"Stopped. Success: {final_success} | Failed: {final_fail}")
            else:
                # Use neutral color for COMPLETED (like IDLE)
                self.update_status("COMPLETED", None, f"Success: {final_success} | Failed: {final_fail}")
            
            _toast(self, f"💳 Captured {final_success} checkout links!", toast_type="success" if final_success > 0 else "warning")
        
        self.after(100, finish_capture)

    # --- UI ANIMATIONS ---
    def _animate_header(self):
        """Cycle header title through vibrant accent colors"""
        try:
            self._header_idx = (self._header_idx + 1) % len(self._header_colors)
            next_color = self._header_colors[self._header_idx]
            self.motion.color(self.header_label, "text_color", next_color, duration_ms=1200, steps=50)
            self.after(4000, self._animate_header)
        except:
            pass
    
    def _animate_logo_glow(self):
        """Animate logo border with color cycling"""
        try:
            colors = [self.colors["accent_primary"], self.colors["accent_secondary"], self.colors["accent_tertiary"]]
            if not hasattr(self, '_logo_color_idx'):
                self._logo_color_idx = 0
            self._logo_color_idx = (self._logo_color_idx + 1) % len(colors)
            self.motion.color(self.logo_frame, "border_color", colors[self._logo_color_idx], duration_ms=1500, steps=40)
            self.after(3000, self._animate_logo_glow)
        except:
            pass
    
    def _blink_cursor(self):
        """Animate terminal cursor blinking"""
        try:
            if not hasattr(self, '_cursor_visible'):
                self._cursor_visible = True
            
            self._cursor_visible = not self._cursor_visible
            if self._cursor_visible:
                self.cursor_indicator.configure(text_color=self.colors["accent_primary"])
            else:
                self.cursor_indicator.configure(text_color=self.colors["glass_bg"])
            
            self.after(530, self._blink_cursor)
        except:
            pass
    
    def _play_entrance_animations(self):
        """Staggered entrance animations for UI elements"""
        try:
            # Animate cards with staggered delays
            elements = [
                (self.controls_container, 0),
                (self.status_frame, 100),
                (self.log_frame, 200),
            ]
            
            for widget, delay in elements:
                self.after(delay, lambda w=widget: self._fade_in_widget(w))
                
        except Exception as e:
            pass
    
    def _fade_in_widget(self, widget):
        """Simple fade-in by animating border glow"""
        try:
            self.motion.color(widget, "border_color", self.colors["glass_border"], duration_ms=400, steps=25)
        except:
            pass
    
    def _start_progress_animation(self):
        """Start custom 60fps progress bar animation with color cycling"""
        import math
        self._progress_phase = 0.0
        self._progress_color_idx = 0
        
        progress_colors = [
            self.colors["accent_primary"],
            self.colors["accent_tertiary"],
            self.colors["accent_secondary"],
        ]
        
        def tick():
            try:
                # Smooth sine wave animation (0 to 1 and back)
                self._progress_phase += 0.025  # Speed control
                value = (math.sin(self._progress_phase) + 1) / 2  # Normalize to 0-1
                self.progress_bar.set(value)
                
                # Cycle colors every ~2 seconds
                if int(self._progress_phase) % 2 == 0 and self._progress_phase % 1 < 0.03:
                    self._progress_color_idx = (self._progress_color_idx + 1) % len(progress_colors)
                    self.motion.color(
                        self.progress_bar, 
                        "progress_color", 
                        progress_colors[self._progress_color_idx], 
                        duration_ms=500
                    )
                
                self._progress_job = self.after(16, tick)  # ~60fps
            except:
                pass
        
        tick()
    
    def _stop_progress_animation(self):
        """Stop the custom progress animation"""
        if self._progress_job:
            try:
                self.after_cancel(self._progress_job)
            except:
                pass
            self._progress_job = None
        self.progress_bar.set(0)
        # Reset to primary color
        try:
            self.progress_bar.configure(progress_color=self.colors["accent_primary"])
        except:
            pass
    
    # --- UI LOGIC ---
    def toggle_reg_inputs(self, choice):
        if choice == "Sequential":
            self.reg_count_label.configure(text="Account Count")
            # Hide threads and delay inputs for Sequential mode
            self.reg_threads_frame.pack_forget()
            self.reg_delay_frame.pack_forget()
        else:
            self.reg_count_label.configure(text="Total")
            # Show threads inline with count, delay on separate row
            self.reg_threads_frame.pack(side="right", padx=(20, 0))
            self.reg_delay_frame.pack(fill="x", padx=16, pady=(0, 10), after=self.reg_count_frame)

    def toggle_mfa_inputs(self, choice):
        if choice == "Sequential":
            self.mfa_threads_label.configure(text="Threads (Disabled):")
            self.mfa_threads_entry.configure(state="disabled")
            # Hide delay input for Sequential mode
            self.mfa_delay_frame.pack_forget()
        else:
            self.mfa_threads_label.configure(text="Threads (Max 5):")
            self.mfa_threads_entry.configure(state="normal")
            # Show delay input for Multithread mode
            self.mfa_delay_frame.pack(fill="x", padx=16, pady=(0, 10), after=self.mfa_threads_frame)
    
    def toggle_checkout_type(self):
        """Show/hide checkout type dropdown based on switch state"""
        if self.reg_checkout_var.get():
            self.reg_checkout_type_frame.pack(side="left", padx=(20, 0))
        else:
            self.reg_checkout_type_frame.pack_forget()
    
    def on_email_mode_change(self, mode):
        """Handle email mode change between TinyHost and OAuth2"""
        global oauth2_accounts
        
        if mode == "OAuth2":
            # Show refresh button
            self.reg_oauth2_refresh.pack(side="left", padx=(8, 0))
            
            # Load oauth2 accounts from oauth2.xlsx
            excel_file = "oauth2.xlsx"
            if os.path.exists(excel_file):
                oauth2_accounts = load_oauth2_accounts_from_excel(excel_file)
                count = len(oauth2_accounts)
                if count > 0:
                    self.reg_oauth2_status.configure(
                        text=f"✅ Loaded {count} OAuth2 accounts",
                        text_color=self.colors["accent_green"]
                    )
                else:
                    self.reg_oauth2_status.configure(
                        text="⚠️ No OAuth2 accounts found in oauth2.xlsx",
                        text_color=self.colors["warning"]
                    )
            else:
                self.reg_oauth2_status.configure(
                    text="❌ oauth2.xlsx not found",
                    text_color=self.colors["error"]
                )
                oauth2_accounts = []
        else:
            # TinyHost mode - clear status and hide refresh button
            self.reg_oauth2_status.configure(text="")
            self.reg_oauth2_refresh.pack_forget()
            oauth2_accounts = []
    
    def refresh_oauth2_accounts(self):
        """Refresh OAuth2 accounts from oauth2.xlsx"""
        global oauth2_accounts
        
        excel_file = "oauth2.xlsx"
        if os.path.exists(excel_file):
            oauth2_accounts = load_oauth2_accounts_from_excel(excel_file)
            count = len(oauth2_accounts)
            if count > 0:
                self.reg_oauth2_status.configure(
                    text=f"✅ Loaded {count} OAuth2 accounts",
                    text_color=self.colors["accent_green"]
                )
            else:
                self.reg_oauth2_status.configure(
                    text="⚠️ No OAuth2 accounts found",
                    text_color=self.colors["warning"]
                )
        else:
            self.reg_oauth2_status.configure(
                text="❌ oauth2.xlsx not found",
                text_color=self.colors["error"]
            )
            oauth2_accounts = []
    
    def toggle_password_visibility(self):

        """Toggle password visibility"""
        self.password_visible = not self.password_visible
        if self.password_visible:
            self.reg_password_entry.configure(show="")
            self.reg_password_toggle.configure(text="🙈")
        else:
            self.reg_password_entry.configure(show="•")
            self.reg_password_toggle.configure(text="👁")
    
    def save_password_to_file(self):
        """Save password to the source code file"""
        new_password = self.reg_password_var.get().strip()
        if not new_password:
            _toast(self, "❌ Password cannot be empty!", toast_type="error")
            return
        
        try:
            # Get the current file path
            current_file = os.path.abspath(__file__)
            
            # Read the file content
            with open(current_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Replace the DEFAULT_PASSWORD line using proper regex
            import re
            # Match: DEFAULT_PASSWORD = "anything"
            pattern = r'DEFAULT_PASSWORD\s*=\s*"[^"]*"'
            replacement = f'DEFAULT_PASSWORD = "{new_password}"'
            
            # Check if pattern exists
            if not re.search(pattern, content):
                _toast(self, "❌ Could not find DEFAULT_PASSWORD in code!", toast_type="error")
                return
            
            new_content = re.sub(pattern, replacement, content, count=1)
            
            # Verify the replacement was made
            if new_content == content:
                _toast(self, "❌ Password replacement failed!", toast_type="error")
                return
            
            # Write back to file
            with open(current_file, 'w', encoding='utf-8') as f:
                f.write(new_content)
            
            # Update global variable
            global DEFAULT_PASSWORD
            DEFAULT_PASSWORD = new_password
            
            _toast(self, f"✅ Password saved: {new_password}", toast_type="success")
            self.reg_password_info.configure(text="✓ Saved!", text_color=self.colors["accent_green"])
            
            # Reset info text after 2 seconds
            self.after(2000, lambda: self.reg_password_info.configure(
                text="Auto-saved to code", 
                text_color=self.colors["text_muted"]
            ))
            
        except Exception as e:
            _toast(self, f"❌ Failed to save: {str(e)}", toast_type="error")
            
    def update_status(self, state="IDLE", color=None, details=""):
        # Map state to icon (keep consistent neutral color for all states)
        state_icons = {
            "IDLE": "●",
            "RUNNING": "◉",
            "LOADING": "◎",
            "COMPLETED": "✓",
            "STOPPED": "◼",
            "ERROR": "✗"
        }
        
        icon = state_icons.get(state.upper(), "●")
        self.status_indicator.configure(text=f"{icon} {state.upper()}")
        
        # Use consistent neutral color for all states (like IDLE)
        neutral_color = self.colors["bg_elevated"]
        neutral_border = self.colors["border_subtle"]
        self.status_indicator.configure(fg_color=neutral_color, border_color=neutral_border)
        
        # Update status dot - keep neutral for all states
        st = state.upper()
        if st == "RUNNING":
            self.status_dot.base_color = self.colors["text_muted"]
            self.status_dot.start_pulse()
        else:
            self.status_dot.stop_pulse()
            self.status_dot.configure(fg_color=self.colors["text_muted"])
            
        if details:
            self.info_box.configure(state="normal")
            self.info_box.delete("0.0", "end")
            self.info_box.insert("0.0", f"→ {details}")
            self.info_box.configure(state="disabled")
            
    def update_stats(self, success, failed):
        """Update stats cards with animated counters"""
        # Get current values
        def get_current(lbl):
            try: 
                text = lbl.cget("text")
                return int(text)
            except: 
                return 0
            
        s0 = get_current(self.success_count_label)
        f0 = get_current(self.fail_count_label)
        
        # Animate success counter
        self.motion.number(
            setter=lambda v: self.success_count_label.configure(text=str(int(v))),
            start=s0, end=success, duration_ms=350, steps=30
        )
        
        # Animate fail counter
        self.motion.number(
            setter=lambda v: self.fail_count_label.configure(text=str(int(v))),
            start=f0, end=failed, duration_ms=350, steps=30
        )
        
        # Flash card borders on change
        if success > s0:
            self.motion.color(self.success_card, "border_color", self.colors["success"], duration_ms=150)
            self.after(300, lambda: self.motion.color(self.success_card, "border_color", "#1a3d32", duration_ms=300))
        
        if failed > f0:
            self.motion.color(self.fail_card, "border_color", self.colors["error"], duration_ms=150)
            self.after(300, lambda: self.motion.color(self.fail_card, "border_color", "#3d1a1a", duration_ms=300))
        
        # Update progress percentage
        total = success + failed
        if total > 0:
            percent = int((success / total) * 100) if total > 0 else 0
            self.progress_percent.configure(text=f"{percent}% success rate")

    def lock_ui(self, is_running):
        self.running = is_running
        state = "disabled" if is_running else "normal"
        stop_state = "normal" if is_running else "disabled"
        
        # Update buttons with animated states
        if is_running:
            self.reg_start_btn.configure(
                state=state, 
                text="⏳  PROCESSING...",
                fg_color=self.colors["bg_elevated"]
            )
            self.mfa_start_btn.configure(
                state=state, 
                text="⏳  PROCESSING...",
                fg_color=self.colors["bg_elevated"]
            )
            self.checkout_start_btn.configure(
                state=state, 
                text="⏳  PROCESSING...",
                fg_color=self.colors["bg_elevated"]
            )
        else:
            self.reg_start_btn.configure(
                state=state, 
                text="▶  START REGISTRATION",
                fg_color=self.colors["accent_green"]
            )
            self.mfa_start_btn.configure(
                state=state, 
                text="🔐  START MFA AUTOMATION",
                fg_color=self.colors["accent_primary"]
            )
            self.checkout_start_btn.configure(
                state=state, 
                text="💳  START CAPTURE",
                fg_color=self.colors["accent_orange"]
            )
        
        # Stop buttons
        self.reg_stop_btn.configure(state=stop_state)
        self.mfa_stop_btn.configure(state=stop_state)
        self.checkout_stop_btn.configure(state=stop_state)
        
        # Animate stop button visibility
        if is_running:
            self.motion.color(self.reg_stop_btn, "border_color", self.colors["error"], duration_ms=200)
            self.motion.color(self.mfa_stop_btn, "border_color", self.colors["error"], duration_ms=200)
            self.motion.color(self.checkout_stop_btn, "border_color", self.colors["error"], duration_ms=200)
        else:
            self.motion.color(self.reg_stop_btn, "border_color", self.colors["border_subtle"], duration_ms=200)
            self.motion.color(self.mfa_stop_btn, "border_color", self.colors["border_subtle"], duration_ms=200)
            self.motion.color(self.checkout_stop_btn, "border_color", self.colors["border_subtle"], duration_ms=200)
        
        # Progress Bar (custom 60fps animation)
        if is_running:
            self._start_progress_animation()
            self.update_status("RUNNING", self.colors["info"], "Initializing automation...")
        else:
            self._stop_progress_animation()
            
        # Lock inputs
        self.reg_mode_menu.configure(state=state)
        self.mfa_mode_menu.configure(state=state)
        self.reg_count_entry.configure(state=state)
        self.reg_threads_entry.configure(state=state)
        self.reg_delay_entry.configure(state=state)
        self.reg_checkout_switch.configure(state=state)
        self.reg_checkout_type_dropdown.configure(state=state)
        self.reg_network_menu.configure(state=state)
        self.checkout_type_menu.configure(state=state)
        self.checkout_refresh_btn.configure(state=state)
        self.checkout_select_all_btn.configure(state=state)
        self.checkout_deselect_all_btn.configure(state=state)
        self.checkout_mt_switch.configure(state=state)
        if not is_running and self.checkout_mt_var.get():
            self.checkout_threads_entry.configure(state="normal")
            self.checkout_delay_entry.configure(state="normal")
        else:
            self.checkout_threads_entry.configure(state="disabled" if is_running else ("normal" if self.checkout_mt_var.get() else "disabled"))
            self.checkout_delay_entry.configure(state="disabled" if is_running else ("normal" if self.checkout_mt_var.get() else "disabled"))
        self.reg_password_entry.configure(state=state)
        self.reg_password_toggle.configure(state=state)
        self.reg_password_save.configure(state=state)
        
        if self.mfa_mode_var.get() == "Multithread":
            self.mfa_threads_entry.configure(state=state)
            self.mfa_delay_entry.configure(state=state)

    def stop_process(self):
        self.stop_event.set()
        self.update_status("STOPPING", self.colors["warning"], "Waiting for threads to finish...")
        _toast(self, "⏹ Stopping automation...", toast_type="warning")

    # --- LOG UTILS ---
    def clear_logs(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("0.0", "end")
        self.log_textbox.configure(state="disabled")
        _toast(self, "🗑️ Logs cleared!", toast_type="info")
        
    def copy_logs(self):
        try:
            text = self.log_textbox.get("0.0", "end")
            self.clipboard_clear()
            self.clipboard_append(text)
            _toast(self, "📋 Copied to clipboard!", toast_type="success")
        except:
            pass
            
    def export_logs(self):
        try:
            filename = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
            if filename:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(self.log_textbox.get("0.0", "end"))
                _toast(self, "💾 Exported successfully!", toast_type="success")
        except Exception as e:
            _toast(self, f"❌ Export failed!", toast_type="error")
            messagebox.showerror("Error", f"Could not export logs: {e}")

    # --- WORKERS ---
    def start_registration_thread(self):
        threading.Thread(target=self.run_registration).start()

    def run_registration(self):
        global GET_CHECKOUT_LINK
        
        self.lock_ui(True)
        self.stop_event.clear()
        self.update_stats(0, 0)
        
        # settings
        global GET_CHECKOUT_LINK, GET_CHECKOUT_TYPE, NETWORK_MODE, oauth2_accounts
        GET_CHECKOUT_LINK = self.reg_checkout_var.get()
        GET_CHECKOUT_TYPE = self.reg_checkout_type_var.get()
        NETWORK_MODE = self.reg_network_var.get()
        email_mode = self.reg_email_mode_var.get()  # "TinyHost" or "OAuth2"
        mode = self.reg_mode_var.get()
        try:
            count = int(self.reg_count_entry.get())
        except:
            count = 1
        
        # OAuth2 mode validation
        if email_mode == "OAuth2":
            # Reload oauth2 accounts to get fresh list
            excel_file = "oauth2.xlsx"
            if os.path.exists(excel_file):
                oauth2_accounts = load_oauth2_accounts_from_excel(excel_file)
            else:
                oauth2_accounts = []
            
            available_count = len(oauth2_accounts)
            if available_count == 0:
                print(f"❌ No OAuth2 accounts available in oauth2.xlsx")
                self.lock_ui(False)
                return
            
            if count > available_count:
                print(f"⚠️ Requested {count} accounts but only {available_count} OAuth2 accounts available. Using {available_count}.")
                count = available_count
            
            # Reset used status
            reset_oauth2_accounts()
            print(f"📧 Using OAuth2 mode with {available_count} available accounts")
            
        # Get thread count (only for multithread mode)
        try:
            threads = int(self.reg_threads_entry.get())
            if threads < 1:
                threads = 2
            if threads > count:
                threads = count  # Can't have more threads than accounts
        except:
            threads = 2
            
        # Get thread delay (only for multithread mode)
        try:
            thread_delay = float(self.reg_delay_entry.get())
            if thread_delay < 0:
                thread_delay = 2
        except:
            thread_delay = 2
            
        if mode == "Multithread":
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting Registration | Mode: {mode} | Email: {email_mode} | Total: {count} | Threads: {threads} | Delay: {thread_delay}s")
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting Registration | Mode: {mode} | Email: {email_mode} | Count: {count}")
        
        success_count = 0
        failed_count = 0
        
        if mode == "Sequential":
            for i in range(count):
                if self.stop_event.is_set():
                    break
                
                self.update_status("RUNNING", self.colors["info"], f"Processing Account {i+1}/{count}")
                print(f"--- Account {i+1}/{count} ---")
                
                try:
                    # Get oauth2 account if in OAuth2 mode
                    oauth2_account = None
                    if email_mode == "OAuth2":
                        oauth2_account = get_next_oauth2_account()
                        if not oauth2_account:
                            print(f"❌ No more OAuth2 accounts available")
                            break
                    
                    worker = ChatGPTAutoRegisterWorker(
                        thread_id=1, 
                        email_mode=email_mode, 
                        oauth2_account=oauth2_account
                    )
                    worker.stop_event = self.stop_event
                    success, result = worker.run()
                    if success:
                        success_count += 1
                        print("✅ Success!")
                        # Mark OAuth2 account as registered
                        if email_mode == "OAuth2" and oauth2_account:
                            mark_oauth2_registered(oauth2_account["row_num"])
                    else:
                        failed_count += 1
                        print("❌ Failed!")
                except Exception as e:
                    failed_count += 1
                    print(f"❌ Error: {e}")

                
                self.update_stats(success_count, failed_count)
                
                if i < count - 1 and not self.stop_event.is_set():
                    print("Waiting 3 seconds...")
                    time.sleep(3)
        else:
            # Multithread with fixed number of threads
            self.update_status("RUNNING", self.colors["info"], f"Processing {count} accounts with {threads} threads (delay: {thread_delay}s)...")
            
            processed = 0
            
            # Create a wrapper function that calculates delay based on slot position
            def run_worker_with_slot_delay(account_idx, slot_idx, stop_event, start_delay, email_mode_inner, num_threads_inner):
                """Worker with computed start delay"""
                # Delay is pre-calculated
                if start_delay > 0:
                    safe_print(account_idx, f"Waiting {start_delay}s before starting...", Colors.INFO, "⏳ ")
                    for _ in range(int(start_delay * 2)):
                        if stop_event and stop_event.is_set():
                            return (False, None, None)
                        time.sleep(0.5)
                
                # Get oauth2 account if in OAuth2 mode
                oauth2_account = None
                oauth2_row_num = None
                if email_mode_inner == "OAuth2":
                    oauth2_account = get_next_oauth2_account()
                    if not oauth2_account:
                        safe_print(account_idx, "No more OAuth2 accounts available", Colors.ERROR, "❌ ")
                        return (False, None, None)
                    oauth2_row_num = oauth2_account.get("row_num")
                
                # Use slot_idx+1 as thread_id for proper window positioning
                worker = ChatGPTAutoRegisterWorker(
                    slot_idx + 1,  # Use slot index for window position (1-based)
                    num_threads=num_threads_inner,
                    email_mode=email_mode_inner,
                    oauth2_account=oauth2_account
                )
                worker.stop_event = stop_event
                success, result = worker.run()
                
                # Mark OAuth2 account as registered if success
                if success and email_mode_inner == "OAuth2" and oauth2_row_num:
                    mark_oauth2_registered(oauth2_row_num)
                
                return (success, result, oauth2_row_num)


            
            with ThreadPoolExecutor(max_workers=threads) as executor:
                # Submit tasks
                futures = {}
                for i in range(count):
                    if self.stop_event.is_set():
                        break
                    
                    # Slot index is (i % threads) for staggered delay in each wave
                    slot_idx = i % threads
                    wave_idx = i // threads
                    
                    if wave_idx == 0:
                        # First wave: stagger by slot
                        start_delay = slot_idx * thread_delay
                    else:
                        # Subsequent waves: Base delay (5s) + stagger
                        # This ensures threads wait 5s + stagger before starting next account
                        start_delay = 5.0 + (slot_idx * thread_delay)
                    
                    future = executor.submit(run_worker_with_slot_delay, i+1, slot_idx, self.stop_event, start_delay, email_mode, threads)
                    futures[future] = i+1
                
                for future in as_completed(futures):
                    if self.stop_event.is_set():
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
                    try:
                        result_tuple = future.result()
                        # Handle both 2-tuple (TinyHost) and 3-tuple (OAuth2) returns
                        if len(result_tuple) == 3:
                            success, result, _ = result_tuple
                        else:
                            success, result = result_tuple
                        processed += 1
                        if success:
                            success_count += 1
                        else:
                            failed_count += 1
                        self.update_stats(success_count, failed_count)
                        self.update_status("RUNNING", self.colors["info"], f"Processed {processed}/{count} accounts")
                    except Exception as e:
                        failed_count += 1
                        processed += 1
                        self.update_stats(success_count, failed_count)
                        print(f"Error: {e}")
                    
                    self.update_stats(success_count, failed_count)

        final_msg = "COMPLETED" if not self.stop_event.is_set() else "STOPPED"
        color = self.colors["success"] if not self.stop_event.is_set() else self.colors["warning"]
        self.update_status(final_msg, color, f"✨ Success: {success_count} | Failed: {failed_count}")
        print(f"\n{'🎉' if not self.stop_event.is_set() else '🛑'} {final_msg}! Success: {success_count} | Failed: {failed_count}")
        
        # Show completion toast
        if not self.stop_event.is_set():
            _toast(self, f"✅ Completed! {success_count} accounts", toast_type="success")
        else:
            _toast(self, f"⏹ Stopped. {success_count} completed", toast_type="warning")
        
        self.lock_ui(False)
        # Restore status with neutral color (like IDLE)
        self.status_indicator.configure(
            text=f"{'✓' if not self.stop_event.is_set() else '◼'} {final_msg}",
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"]
        )

    def start_mfa_thread(self):
        threading.Thread(target=self.run_mfa).start()
    
    def run_mfa(self):
        self.lock_ui(True)
        self.stop_event.clear()
        self.update_stats(0, 0)
        
        mode = self.mfa_mode_var.get()
        try:
            threads = int(self.mfa_threads_entry.get())
            threads = max(1, min(5, threads))
        except:
            threads = 1
            
        # Get thread delay (only for multithread mode)
        try:
            mfa_thread_delay = float(self.mfa_delay_entry.get())
            if mfa_thread_delay < 0:
                mfa_thread_delay = 2
        except:
            mfa_thread_delay = 2
            
        if mode == "Multithread":
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting MFA | Mode: {mode} | Threads: {threads} | Delay: {mfa_thread_delay}s")
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting MFA | Mode: {mode}")
        self.update_status("LOADING", self.colors["info"], "Loading accounts from Excel...")
        
        excel_file = "chatgpt.xlsx"
        accounts = load_mfa_accounts(excel_file)
        
        if not accounts:
            print("No suitable accounts found.")
            self.lock_ui(False)
            self.update_status("IDLE", self.colors["text_muted"], "No accounts found in Excel.")
            _toast(self, "⚠️ No accounts found!", toast_type="warning")
            return
            
        print(f"Found {len(accounts)} accounts.")
        
        # Load OAuth2 accounts for lookup (to enable code reading for OAuth2 emails)
        oauth2_map = {}
        if os.path.exists("oauth2.xlsx"):
            try:
                raw_oauth2 = load_oauth2_accounts_from_excel("oauth2.xlsx", skip_registered=False)
                for acc in raw_oauth2:
                    if acc.get('email'):
                        oauth2_map[acc['email'].lower()] = acc
                if len(oauth2_map) > 0:
                    print(f"Loaded {len(oauth2_map)} OAuth2 accounts for reference.")
            except Exception as e:
                print(f"Error loading oauth2.xlsx: {e}")
        
        success_count = 0
        fail_count = 0
        
        if mode == "Sequential":
            for idx, account in enumerate(accounts):
                if self.stop_event.is_set():
                    break
                    
                self.update_status("RUNNING", self.colors["info"], f"🔐 {account['email'][:20]}... ({idx+1}/{len(accounts)})")
                print(f"Processing {account['email']}...")
                
                # Get OAuth2 info if available
                oauth2_acc = oauth2_map.get(account['email'].lower())
                
                try:
                    result, email = run_mfa_worker(1, account, excel_file, self.stop_event, oauth2_account=oauth2_acc)
                    if result:
                        success_count += 1
                        print(f"✅ MFA Enabled: {email}")
                    else:
                        fail_count += 1
                        print(f"❌ Failed: {email}")
                except Exception as e:
                    fail_count += 1
                    print(f"Error: {e}")
                
                self.update_stats(success_count, fail_count)
                
                if idx < len(accounts) - 1:
                    time.sleep(3)
        else:
            self.update_status("RUNNING", self.colors["info"], f"🔐 Running with {threads} threads (delay: {mfa_thread_delay}s)...")
            with ThreadPoolExecutor(max_workers=threads) as executor:
                futures = {}
                for idx, account in enumerate(accounts):
                    if self.stop_event.is_set():
                        break
                    
                    # Get OAuth2 info if available
                    oauth2_acc = oauth2_map.get(account['email'].lower())
                    
                    thread_id = (idx % threads) + 1
                    slot_index = idx % threads           # 0, 1, 2 for 3 threads
                    batch_number = idx // threads        # Which batch this account belongs to
                    is_first_batch = (batch_number == 0) # First batch has no base delay
                    # Pass slot-based delay parameters
                    future = executor.submit(run_mfa_worker, thread_id, account, excel_file, self.stop_event, mfa_thread_delay, slot_index, is_first_batch, oauth2_acc)
                    futures[future] = account
                
                for future in as_completed(futures):
                    if self.stop_event.is_set():
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
                    try:
                        result, email = future.result()
                        if result:
                            success_count += 1
                        else:
                            fail_count += 1
                    except Exception as e:
                        print(f"Error: {e}")
                    
                    self.update_stats(success_count, fail_count)

        final_msg = "COMPLETED" if not self.stop_event.is_set() else "STOPPED"
        color = self.colors["success"] if not self.stop_event.is_set() else self.colors["warning"]
        self.update_status(final_msg, color, f"🔐 MFA Success: {success_count} | Failed: {fail_count}")
        print(f"\n{'🎉' if not self.stop_event.is_set() else '🛑'} MFA {final_msg}! Success: {success_count} | Failed: {fail_count}")
        
        # Show completion toast
        if not self.stop_event.is_set():
            _toast(self, f"🔐 MFA enabled for {success_count} accounts!", toast_type="success")
        else:
            _toast(self, f"⏹ MFA stopped. {success_count} completed", toast_type="warning")
        
        self.lock_ui(False)
        # Restore status with neutral color (like IDLE)
        self.status_indicator.configure(
            text=f"{'✓' if not self.stop_event.is_set() else '◼'} {final_msg}",
            fg_color=self.colors["bg_elevated"],
            border_color=self.colors["border_subtle"]
        )

if __name__ == "__main__":
    app = App()
    app.mainloop()

