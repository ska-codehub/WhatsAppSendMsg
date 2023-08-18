# main file for the scraping
# @author: Sk Khurshid Alam

import os
import sys
import shutil
import signal
import psutil
from contextlib import suppress
import configparser
import argparse
from pathlib import Path
import time
import random
import pandas as pd
import urllib
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import traceback

TRUTHY = [
    'TRUE', 'True', 'true', 'T', 't', True,
    '1', 1,
    'ON', 'On', 'on',
    'YES', 'Yes', 'yes', 'Y', 'y'
]


FALSY = [
    'FALSE', 'False', 'false', 'F', 'f', False,
    '0', 0,
    'OFF', 'Off', 'off',
    'NO', 'No', 'no', 'N', 'n'
]

BASE_DIR = Path("./")
CONFIGURATION_FILE = BASE_DIR / 'settings.config'
CHROME_DIR = BASE_DIR / "chrome"
USER_DATA_DIR = CHROME_DIR / "user-data"
CONTACT_FOLDER_PATH = BASE_DIR / "contacts"


CHROME_DIR.mkdir(exist_ok=True)
USER_DATA_DIR.mkdir(exist_ok=True)
CONTACT_FOLDER_PATH.mkdir(exist_ok=True)

config = configparser.RawConfigParser()
config.read(CONFIGURATION_FILE)
SETTINGS = dict(config.items('settings'))

PROJECT_NAME = SETTINGS.get('project_name').strip()
PROJECT_DESCRIPTION = SETTINGS.get('project_description').strip()
VERSION = SETTINGS.get('version').strip()
SITE_DOMAIN = SETTINGS.get('site_domain').strip()
LOGIN_URL = SETTINGS.get('login_url').strip()
LOGIN_TITLE = SETTINGS.get('login_title').strip()
LOGIN_REDIRECT_TITLE = SETTINGS.get('login_redirect_title').strip()
SEND_URL = SETTINGS.get('send_url').strip()
CONTACT_NUMBER_COLUMN_NAME = SETTINGS.get('message_file').strip()
CONTACT_NAME_COLUMN_NAME = SETTINGS.get('message_file').strip()
MESSAGE_FILE = BASE_DIR / SETTINGS.get('message_file').strip()

HEALTH_CHECK_URL = SETTINGS.get('health_check_url').strip()
HEALTH_CHECK_TITLE = SETTINGS.get('health_check_title').strip()

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.5735.90",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36", 
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36",
]


def confirmation_input(ask_str, ask_type):
    if ask_type not in ['Y/n', 'y/N', 'N/y', 'n/Y']:
        ask_type = 'Y/n'
    ask_str = f"{ask_str} [{ask_type}]: "
    while True:
        ask_value = input(ask_str).lower()
        if not ask_value in [''] + TRUTHY + FALSY:
            print("Please provide a valid confirmation!")
            continue
        if ask_type == 'Y/n':
            return ask_value in [''] + TRUTHY
        elif ask_type == 'y/N':
            return ask_value in TRUTHY
        elif ask_type == 'N/y':
            return ask_value in TRUTHY
        elif ask_type == 'n/Y':
            return ask_value in [''] + TRUTHY


class WhatsAppSendMsg:
    def __init__(self,invisible=False) -> None:
        self.invisible = invisible
        self.user_agent = USER_AGENTS[random.randrange(0, len(USER_AGENTS)-1)]
        self.page_load_timeout = 60
        self.not_ok = 0
        self.retry = 0
        self.max_retries = 3

    def is_head_ready(self):
        try:
            WebDriverWait(self.browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, "head")))
            return self.browser.find_element(By.TAG_NAME, "head") is not None 
        except Exception as e:
            print("WhatsAppSendMsg.is_head_ready Error: ", e, traceback.format_exc())
            return False
    
    def is_dom_ready(self):
        try:
            time.sleep(0.5)
            WebDriverWait(self.browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))                
            try:
                self.browser.execute_script(f'window.scrollTo(0, {random.randrange(100, 1000)})')
            except:
                pass
            return self.browser.find_element(By.TAG_NAME, "body") is not None
        except Exception as e:
            print("WhatsAppSendMsg.is_dom_ready Error: ", e, traceback.format_exc())
            return False        

    def is_title_valid(self, title=""):
        try:
            return self.browser.title==title or self.browser.title.strip()==title.strip()
        except Exception as e:
            print("WhatsAppSendMsg.is_title_valid Error: ", e, traceback.format_exc())
            return False
        
    def is_page_ready(self, title):
        ready = False
        for _ in range(0, 3):
            try:
                time.sleep(1)
                ready = self.is_head_ready() and self.is_dom_ready() and self.is_title_valid(title)
                if ready:
                    break
            except:
                ready = False
        return ready


    def get_page(self, url, title=None):
        try:
            self.browser.get(url)
            time.sleep(1)
            return self.is_page_ready(title)
        except TimeoutException as e:
            print("WhatsAppSendMsg.get_page Error1: ", e, traceback.format_exc())
            if self.retry<=self.max_retries:
                self.retry += 1
                # self.config_browser()
                return self.get_page(url, title)
            return False
        except Exception as e:
            print("WhatsAppSendMsg.get_page Error2: ", e, traceback.format_exc())
            return False

    def test_browser_ok(self):
        print("Testing browser")
        if self.get_page(HEALTH_CHECK_URL, HEALTH_CHECK_TITLE):
            self.not_ok = 0
            print("OK")
            return True
        else:
            self.not_ok += 1
            print("NOT OK")
            return False
    
    def kill_browser_process(self):     
        try:
            if hasattr(self, "browser") and self.browser is not None:
                print("Killing browser instances and process")
                try:
                    pid = int(self.browser.service.process.id)
                except:
                    pid = None
                try:
                    self.browser.service.process.send_signal(signal.SIGTERM)
                except:
                    pass
                try:
                    self.browser.close()
                except:
                    pass
                try:
                    self.browser.quit()
                except:
                    pass
                try:
                    if pid is not None:
                        os.kill(pid, signal.SIGTERM)
                except:
                    pass
            try:
                for process in psutil.process_iter():
                    try:
                        if process.name() == "chrome.exe" \
                            and "--test-type=webdriver" in process.cmdline():
                            with suppress(psutil.NoSuchProcess):
                                try:
                                    os.kill(process.pid, signal.SIGTERM)
                                except:
                                    pass
                    except:
                        pass
            except:
                pass
        except:
            pass
        
        if hasattr(self, "browser"):
            if self.browser is None or not hasattr(self.browser, "service") or self.browser.service:
                print("Browser closed and webdriver process killed!")
                self.browser = None
            else:
                print("Browser and Webdriver process NOT killed !!!!")


    def config_browser(self, *args, **kwargs):
        print("Configuring browser...")
        chrome_driver_path = CHROME_DIR / 'chromedriver.exe'
        self.kill_browser_process()
        options = Options()
        options.page_load_strategy = "none"
        options.add_argument("--start-maximized")
        options.add_argument("--ignore-gpu-blacklist")
        options.add_argument("--use-gl")
        options.add_argument("--allow-insecure-localhost")
        options.add_argument("--allow-running-insecure-content")
        options.add_argument("--ignore-ssl-errors=yes")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--block-insecure-private-network-requests=false")
        options.add_argument(f"--unsafely-treat-insecure-origin-as-secure={SITE_DOMAIN}")
        options.add_argument("--safebrowsing-disable-download-protection")
        options.add_argument("--disable-gpu")
        if self.invisible:
            print("Configuring browser with invisible mode!")
            options.add_argument("--headless")
        else:
            print("Configuring browser with visible mode!")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument(f"user-agent={self.user_agent}")
        options.add_argument("--kiosk-printing")
        options.add_argument("--disable-blink-features")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-notifications")
        options.add_argument(f"user-data-dir={USER_DATA_DIR.absolute()}")
        options.set_capability("acceptInsecureCerts", True)
        options.add_experimental_option("useAutomationExtension", False)
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        prefs = {
            "profile.default_content_setting_values.notifications" : 2,
            "safebrowsing_for_trusted_sources_enabled" : False,
            "safebrowsing.enabled" : False,
            "profile.exit_type" : "Normal"
        }
        options.add_experimental_option("prefs", prefs)
        os.environ["webdriver.chrome.driver"] = str(chrome_driver_path.absolute())
        service = Service(executable_path=chrome_driver_path, service_args=["--verbose"])
        self.browser = webdriver.Chrome(service=service, options=options)
        self.browser.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.browser.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent":self.user_agent})
        self.browser.set_page_load_timeout(self.page_load_timeout)
        self.browser.maximize_window()
        print("browserVersion: ", self.browser.capabilities["browserVersion"])
        print("chromedriverVersion: ", self.browser.capabilities["chrome"]["chromedriverVersion"].split(" ")[0])
        if not self.test_browser_ok():
            self.retry += 1
            if self.retry<=self.max_retries:
                self.config_browser()
            else:
                raise Exception("Failed to configure browser. Possible reason: Blocked Proxy server")
        else:
            self.retry = 0

        
    def get_clickable_element(self, by_tuple):
        el = None
        try:
            el = WebDriverWait(self.browser, 20).until(EC.presence_of_element_located(by_tuple))
            el1 = WebDriverWait(self.browser, 10).until(EC.element_to_be_clickable(by_tuple))
            if el1 is not None:
                el = el1
        except Exception as e:
            print("WhatsAppSendMsg.get_clickable_element Error: ", e)
        return el

    def cleanup_session_login(self):
        if self.is_title_valid(None) or self.is_title_valid(""):
            print("Re-configuring and Re-loging...")
            if USER_DATA_DIR.exists():
                self.kill_browser_process()
                shutil.rmtree(USER_DATA_DIR)
                USER_DATA_DIR.mkdir(exist_ok=True)
            self.config_browser()
            return self.login()
        return False

    def login(self):
        if self.get_page(LOGIN_URL, LOGIN_TITLE):
            time.sleep(3)
            qrcode_el = self.browser.find_elements(By.XPATH, "//div[@data-testid='qrcode']")
            if len(qrcode_el)==0:
                landing_title_el = self.browser.find_elements(By.XPATH, "//div[@class='landing-title']")
                if len(landing_title_el)>0:
                    self.browser.refresh()
                else:
                    time.sleep(10)
                    return True
            print("Please scan the QR Code to login!")
            while True:
                if confirmation_input("Done with scanning QR code?", 'y/N')==True:
                    qrcode_el = self.browser.find_elements(By.XPATH, "//div[@data-testid='qrcode']")
                    if len(qrcode_el)==0:
                        time.sleep(1)
                        print(f"Waiting for login redirect title {LOGIN_REDIRECT_TITLE}!")
                        if not self.is_page_ready(LOGIN_REDIRECT_TITLE):
                            if not self.is_title_valid(LOGIN_REDIRECT_TITLE):
                                if not self.cleanup_session_login():
                                    print("Couldn't login!!! Re-loging....")
                                    return self.login()
                        time.sleep(3)
                        return True
                    else:
                        print("Scanning not done yet! Please scan the QR code...")
        return self.cleanup_session_login()
        
    def click_send(self):
        by_tuple = (By.XPATH, f"//div[@id='main']//footer//div[@data-testid='compose-box']//button[@data-testid='compose-btn-send']")
        el = self.get_clickable_element(by_tuple)
        if el:
            el.click()
            time.sleep(2)    
    
    def start_sending_msg(self):
        try:
            if not MESSAGE_FILE.exists():
                print(f"There's not message file {MESSAGE_FILE}")
                return
            message = None
            with open(MESSAGE_FILE, "r", encoding="utf-8") as f:
                message = f.read()
            if message is None or len(message)==0:
                print(f"There's no message in message file {MESSAGE_FILE}")
                return
            print("Message to send: ", message)
            
            contact_numbers = []
            contact_names = []
            if CONTACT_FOLDER_PATH.exists():
                for f in CONTACT_FOLDER_PATH.iterdir():
                    if f.is_file() and f.suffix=='.xlsx':
                        print(f"Processing {f.name}")
                        sheets = list(pd.read_excel(f, sheet_name=None))
                        print("Total Sheets:", len(sheets))
                        for sheet in sheets:
                            df = pd.read_excel(f, sheet_name=sheet)
                            try:
                                if CONTACT_NUMBER_COLUMN_NAME in df and CONTACT_NAME_COLUMN_NAME in df:
                                    contact_numbers = list(df[CONTACT_NUMBER_COLUMN_NAME])
                                    contact_names = list(df[CONTACT_NAME_COLUMN_NAME])
                            except Exception as e:
                                print("WhatsAppSendMsg.start_sending_msg Error: ", e, traceback.format_exc())
            if len(contact_numbers)==0:
                print("There's no contact number to process.")
                return
            if len(contact_numbers)!=len(contact_names):
                print(f"Each contact number should be mapped to a contact name and vice-versa.\nTotal count of contact numbers: {len(contact_numbers)}, Total count of contact names: {len(contact_names)}")
                return
            max_retries = 3
            retry = 0
            while True:
                try:
                    self.config_browser()
                    break
                except Exception as e:
                    print("WhatsAppSendMsg.start_sending_msg Error: ", e, traceback.format_exc())
                    retry += 1
                if retry>max_retries:
                    raise Exception("Browser can't be configured at this moment!")
            time.sleep(1)
            if self.login():
                print("Logged in")
                WebDriverWait(self.browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, "title")))
                for i, contact_number in enumerate(contact_numbers):
                    contact_number = str(contact_number).strip()
                    contact_name = str(contact_names[i]).strip()
                    if not contact_number.startswith("+91"):
                        contact_number = f"+91{contact_number.lstrip('91')}"
                    print(f"Processing {contact_name}, {contact_number}")
                    try:
                        text = urllib.parse.urlencode({'text' : message % (contact_name, )})
                        send_url = SEND_URL % (contact_number, text)
                        if self.get_page(send_url, LOGIN_TITLE):
                            time.sleep(1)
                            self.click_send()
                            print(f"######################## SENT TO: {contact_name}, {contact_number} ########################")
                    except Exception as e:
                        print("WhatsAppSendMsg.start_sending_msg Error1: ", e, traceback.format_exc())
                    print()
                print("######################## COMPLETED ########################")
        except Exception as e:
            print("WhatsAppSendMsg.start_sending_msg Error2: ", e, traceback.format_exc())
        self.kill_browser_process()
            


if __name__ == "__main__":
    argv = sys.argv
    parser = argparse.ArgumentParser(prog=PROJECT_NAME, description=PROJECT_DESCRIPTION)
    parser.version = VERSION
    parser.add_argument("-v", "--version", action="version", version=parser.version)
    parser.add_argument(
        "-inviz",
        "--invisible",
        dest="INVISIBLE",
        action='store_true',
        default=False,
        required=False,
        help="Run script with visible mode, default: visible"
    )

    args = parser.parse_args(argv[1:])
    print(parser.description)
    INVISIBLE = args.INVISIBLE
    print("INVISIBLE: ", INVISIBLE)
    wp_scraper = WhatsAppSendMsg(invisible=INVISIBLE)
    wp_scraper.start_sending_msg()
