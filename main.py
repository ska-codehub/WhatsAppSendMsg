# main file for the scraping
# @author: Sk Khurshid Alam

import os
import sys
import subprocess
import signal
import psutil
from contextlib import suppress
import argparse
import time
import random
import shutil
import pandas as pd
import numpy as np
import urllib
from copy import deepcopy
from weakref import finalize
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.service import utils
from selenium.webdriver.chromium.service import ChromiumService
from selenium.webdriver.chromium.options import ChromiumOptions
from patcher import Patcher
from utils import *
from settings import *
import traceback

class WhatsAppSendMsg:
    def __init__(self, invisible=True, debug=False) -> None:
        finalize(self, self.kill_browser_process)
        self.invisible = invisible
        self.debug = debug
        self.not_ok = 0
        self.retry = 0
        self.max_retries = 3
        self.force_kill = True

    def is_head_ready(self):
        ready = False
        try:
            WebDriverWait(self.browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, "head")))
            ready = self.browser.find_element(By.TAG_NAME, "head") is not None 
        except Exception as e:
            if "Alert Text" in str(e):
                print("############## Alert detected.")
            else:
                print("WhatsAppSendMsg.is_head_ready Error: ", e, traceback.format_exc())
        print("Head Ready?: ", "Yes" if ready else "No")
        return ready

    def scroll(self):
        try:
            print("Scrolling...")
            scroll_height = self.browser.execute_script("return document.body.scrollHeight;")
            start_time = time.time()
            old_scroll_height = None
            while True:
                i = 10
                while i>0:
                    self.browser.execute_script(f"window.scrollTo(0, {scroll_height/i});")
                    time.sleep(random.randrange(3, 5)/100)
                    i -= 1
                    if time.time() - start_time >= self.scroll_timeout:
                        break
                time.sleep(random.randrange(3, 5)/100)
                new_scroll_height = self.browser.execute_script("return document.body.scrollHeight;")
                if new_scroll_height==scroll_height:
                    break
                if old_scroll_height is None or old_scroll_height!=new_scroll_height:
                    old_scroll_height = new_scroll_height
                else:
                    break                
                if time.time() - start_time >= self.scroll_timeout:
                    break
        except:
            pass
        finally:
            self.browser.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.5)

    def is_dom_ready(self):
        ready = False
        try:
            time.sleep(0.5)
            WebDriverWait(self.browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))                
            try:
                self.scroll()
            except:
                pass
            ready = self.browser.find_element(By.TAG_NAME, "body") is not None
        except Exception as e:
            if "Alert Text" in str(e):
                print("############## Alert detected.")

            else:
                print("WhatsAppSendMsg.is_dom_ready Error: ", e, traceback.format_exc())
        print("Dom Ready?: ", "Yes" if ready else "No")
        return ready        

    def is_title_valid(self, title=None, invalid_title=None):
        valid = False
        print(f"Checking Title {title}")
        if title is None or title=="":
            if invalid_title is not None:
                try:
                    WebDriverWait(self.browser, 1).until(EC.title_contains(invalid_title.strip()))
                except TimeoutException:
                    valid = True

                if not valid:
                    print("BOT DETECTED")
            else:
                valid = True
        else:
            try:
                WebDriverWait(self.browser, 15).until(EC.title_contains(title.strip()))
                valid = True
            except TimeoutException:
                pass
        print("Is Title Valid?: ", "Yes" if valid else "No")
        return valid
        
    def is_page_ready(self, title=None, invalid_title=None, max_try=3):
        print("Is Page Ready?")
        ready = False
        for _ in range(0, max_try):
            try:
                ready = self.is_head_ready() and self.is_dom_ready() and self.is_title_valid(title=title, invalid_title=invalid_title)
                if ready:
                    break
            except Exception as e:
                print("is_page_ready Error: ", e, traceback.format_exc())
                ready = False
        print("Page Ready Status:", "Yes" if ready else "No")
        return ready


    def get_page(self, url, title=None, invalid_title=None):
        try:
            self.browser.get(url)
            time.sleep(1)
            return self.is_page_ready(title=title, invalid_title=invalid_title)
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

    
    def kill_browser_process(self, all=False):     
        try:
            if hasattr(self, "browser") and self.browser is not None:
                print("Killing browser instances and process")
                try:
                    pid = int(self.browser.service.process.pid)
                except:
                    pid = None
                try:
                    self.browser.close()
                except:
                    pass
                try:
                    self.browser.quit()
                except:
                    pass
                try:
                    self.browser.service.process.terminate()
                except:
                    pass                
                try:
                    self.browser.service.process.kill()
                except:
                    pass
                try:
                    self.browser.service.process.send_signal(signal.SIGTERM)
                except:
                    pass
                try:
                    if pid is not None:
                        os.kill(pid, signal.SIGTERM)
                except PermissionError:
                    try:
                        subprocess.check_output("Taskkill /PID %d /F" % pid)
                    except:
                        pass
                except:
                    pass

            if all:
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
            if self.browser is None or not hasattr(self.browser, "service") or (
                self.browser is not None and
                hasattr(self.browser, "service") and
                not self.browser.service.is_connectable()):
                print("Browser closed and webdriver process killed!")
                self.browser = None
            else:
                print("Browser and Webdriver process NOT killed !!!!")
                if self.force_kill:
                    self.force_kill = False
                    self.kill_browser_process(all=True)



    def _configure_headless(self):
        orig_get = self.browser.get
        print("setting properties for headless")

        def get_wrapped(*args, **kwargs):
            if self.browser.execute_script("return navigator.webdriver"):
                print("patch navigator.webdriver")
                self.browser.execute_cdp_cmd(
                    "Page.addScriptToEvaluateOnNewDocument",
                    {
                        "source": """

                           Object.defineProperty(window, "navigator", {
                                Object.defineProperty(window, "navigator", {
                                  value: new Proxy(navigator, {
                                    has: (target, key) => (key === "webdriver" ? false : key in target),
                                    get: (target, key) =>
                                      key === "webdriver"
                                        ? false
                                        : typeof target[key] === "function"
                                        ? target[key].bind(target)
                                        : target[key],
                                  }),
                                });
                    """
                    },
                )

                print("patch user-agent string")
                self.browser.execute_cdp_cmd(
                    "Network.setUserAgentOverride",
                    {
                        "userAgent": self.browser.execute_script(
                            "return navigator.userAgent"
                        ).replace("Headless", "")
                    },
                )
                self.browser.execute_cdp_cmd(
                    "Page.addScriptToEvaluateOnNewDocument",
                    {
                        "source": """
                            Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 1});
                            Object.defineProperty(navigator.connection, 'rtt', {get: () => 100});
                            window.chrome = {
                                app: {
                                    isInstalled: false,
                                    InstallState: {
                                        DISABLED: 'disabled',
                                        INSTALLED: 'installed',
                                        NOT_INSTALLED: 'not_installed'
                                    },
                                    RunningState: {
                                        CANNOT_RUN: 'cannot_run',
                                        READY_TO_RUN: 'ready_to_run',
                                        RUNNING: 'running'
                                    }
                                },
                                runtime: {
                                    OnInstalledReason: {
                                        CHROME_UPDATE: 'chrome_update',
                                        INSTALL: 'install',
                                        SHARED_MODULE_UPDATE: 'shared_module_update',
                                        UPDATE: 'update'
                                    },
                                    OnRestartRequiredReason: {
                                        APP_UPDATE: 'app_update',
                                        OS_UPDATE: 'os_update',
                                        PERIODIC: 'periodic'
                                    },
                                    PlatformArch: {
                                        ARM: 'arm',
                                        ARM64: 'arm64',
                                        MIPS: 'mips',
                                        MIPS64: 'mips64',
                                        X86_32: 'x86-32',
                                        X86_64: 'x86-64'
                                    },
                                    PlatformNaclArch: {
                                        ARM: 'arm',
                                        MIPS: 'mips',
                                        MIPS64: 'mips64',
                                        X86_32: 'x86-32',
                                        X86_64: 'x86-64'
                                    },
                                    PlatformOs: {
                                        ANDROID: 'android',
                                        CROS: 'cros',
                                        LINUX: 'linux',
                                        MAC: 'mac',
                                        OPENBSD: 'openbsd',
                                        WIN: 'win'
                                    },
                                    RequestUpdateCheckStatus: {
                                        NO_UPDATE: 'no_update',
                                        THROTTLED: 'throttled',
                                        UPDATE_AVAILABLE: 'update_available'
                                    }
                                }
                            }

                            // https://github.com/microlinkhq/browserless/blob/master/packages/goto/src/evasions/navigator-permissions.js
                            if (!window.Notification) {
                                window.Notification = {
                                    permission: 'denied'
                                }
                            }

                            const originalQuery = window.navigator.permissions.query
                            window.navigator.permissions.__proto__.query = parameters =>
                                parameters.name === 'notifications'
                                    ? Promise.resolve({ state: window.Notification.permission })
                                    : originalQuery(parameters)

                            const oldCall = Function.prototype.call
                            function call() {
                                return oldCall.apply(this, arguments)
                            }
                            Function.prototype.call = call

                            const nativeToStringFunctionString = Error.toString().replace(/Error/g, 'toString')
                            const oldToString = Function.prototype.toString

                            function functionToString() {
                                if (this === window.navigator.permissions.query) {
                                    return 'function query() { [native code] }'
                                }
                                if (this === functionToString) {
                                    return nativeToStringFunctionString
                                }
                                return oldCall.call(oldToString, this)
                            }
                            // eslint-disable-next-line
                            Function.prototype.toString = functionToString
                            """
                    },
                )
            return orig_get(*args, **kwargs)
        self.browser.get = get_wrapped


    def find_chrome_executable(self):
        candidates = set()
        if IS_POSIX:
            for item in os.environ.get("PATH").split(os.pathsep):
                for subitem in (
                    "google-chrome",
                    "chromium",
                    "chromium-browser",
                    "chrome",
                    "google-chrome-stable",
                ):
                    candidates.add(os.sep.join((item, subitem)))
            if "darwin" in sys.platform:
                candidates.update(
                    [
                        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                        "/Applications/Chromium.app/Contents/MacOS/Chromium",
                    ]
                )
        else:
            for item in map(
                os.environ.get,
                ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA", "PROGRAMW6432"),
            ):
                if item is not None:
                    for subitem in (
                        "Google/Chrome/Application",
                    ):
                        candidates.add(os.sep.join((item, subitem, "chrome.exe")))
        for candidate in candidates:
            print('checking if %s exists and is executable' % candidate)
            if os.path.exists(candidate) and os.access(candidate, os.X_OK):
                print('found! using %s' % candidate)
                return os.path.normpath(candidate)


    def config_browser(self):
        print("Configuring browser...")
        chrome_driver_filename = 'undetected_chromedriver'
        if not IS_POSIX:
            chrome_driver_filename += '.exe'
        chrome_driver_path = CHROME_DIR / chrome_driver_filename
        self.patcher = Patcher(user_multi_procs=True)
        self.patcher.auto(executable_path=chrome_driver_path)
        options = ChromiumOptions()
        # options.page_load_strategy = "normal"
        options.add_argument("--disable-extensions")
        options.add_argument("--lang=en-IN")
        options.arguments.extend(["--no-default-browser-check", "--no-first-run"])
        options.arguments.extend(["--no-sandbox", "--test-type"])
        if self.invisible:
            print("Configuring browser with invisible mode!")
            options.add_argument("--headless=new")
            if self.debug:
                debug_host = "127.0.0.1"
                debug_port = utils.free_port()
                options.add_argument(f"--remote-debugging-host={debug_host}")
                options.add_argument(f"--remote-debugging-port={debug_port}")
                options.debugger_address = "%s:%d" % (debug_host, debug_port)
        else:
            print("Configuring browser with visible mode!")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument('--disable-application-cache')
        options.add_argument("--disable-session-crashed-bubble")
        if USER_DATA_DIR.exists():
            options.add_argument(f"user-data-dir={str(USER_DATA_DIR.absolute())}")
        options.add_experimental_option("useAutomationExtension", False)
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option("prefs", {
            "profile.exit_type" : "Normal",
            "profile.default_content_setting_values.notifications": 2
        })

        os.environ["webdriver.chrome.driver"] = str(chrome_driver_path.absolute())
        
        desired_capabilities = options.to_capabilities()
        # print("desired_capabilities: ", desired_capabilities)

        # binary_location = self.find_chrome_executable()
        # if binary_location is not None:
        #     options.binary_location = binary_location

        service = ChromiumService(executable_path=chrome_driver_path)
        self.browser = webdriver.chrome.webdriver.WebDriver(service=service, options=options, keep_alive=False)
        self.browser._delay = 3
        # self.browser.user_data_dir = str(USER_DATA_DIR.absolute())
        self.browser.keep_user_data_dir = True
        if self.invisible:
            self._configure_headless()
        # self.browser.set_page_load_timeout(self.page_load_timeout)
        self.browser.maximize_window()
        print("browserVersion: ", self.browser.capabilities["browserVersion"])
        print("chromedriverVersion: ", self.browser.capabilities["chrome"]["chromedriverVersion"].split(" ")[0])
        if not self.test_browser_ok():
            self.retry += 1
            if self.retry<=self.max_retries:
                self.kill_browser_process(all=True)
                self.config_browser()
            else:
                raise Exception("Failed to configure browser.")
        else:
            self.retry = 0

    def get_prensented_elements(self, by_tuple, timeout=10):
        els = []
        try:
            els = WebDriverWait(self.browser, timeout).until(EC.presence_of_all_elements_located(by_tuple))
        except Exception as e:
            print(f"WhatsAppSendMsg.get_prensented_elements {e}")
            els = []
        return els

    def get_clickable_element(self, by_tuple, timeout=20):
        el = None
        error = ""
        sep = ""
        try:
            els = self.get_prensented_elements(by_tuple=by_tuple)
            if len(els)>0:
                el = els[0]
        except Exception as e:
            error += f"Error1: {e}"
            sep = "\n"
        try:
            el1 = WebDriverWait(self.browser, timeout).until(EC.element_to_be_clickable(by_tuple))
            if el1 is not None:
                el = el1
        except Exception as e:
            error += f"{sep}Error2: {e}"
        if len(error)>0:
            error = f"WhatsAppSendMsg.get_clickable_element {error}"
            print(error)
        return el


    def cleanup_session_login(self):
        if self.is_title_valid(None) or self.is_title_valid(""):
            print("Re-configuring and Re-loging...")
            if USER_DATA_DIR.exists():
                self.kill_browser_process()
                shutil.rmtree(USER_DATA_DIR)
                USER_DATA_DIR.mkdir(exist_ok=True)
            self.kill_browser_process(all=True)
            self.config_browser()
            return self.login()
        return False

    def login(self):
        if self.get_page(LOGIN_URL, LOGIN_TITLE):
            time.sleep(3)
            initial_startup_el = self.browser.find_elements(By.XPATH, "//div[@id='initial_startup']")
            if len(initial_startup_el)==0:
                landing_title_el = self.browser.find_elements(By.XPATH, "//div[@class='landing-title']")
                if len(landing_title_el)>0:
                    self.browser.refresh()
                else:
                    time.sleep(10)
                    return True
            print("Please scan the QR Code to login!")
            while True:
                if confirmation_input("Done with scanning QR code?", 'y/N')==True:
                    initial_startup_el = self.browser.find_elements(By.XPATH, "//div[@id='initial_startup']")
                    if len(initial_startup_el)==0:
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
                        print("Not logged in yet. Possible reasons: \n1> Scanning not done yet! Please scan the QR code...\n2> Slow network, still loading...")
        return self.cleanup_session_login()
        
    def is_message_link_rendered(self):
        if MESSAGE_LINK_RENDERED_TITLE is not None and len(MESSAGE_LINK_RENDERED_TITLE)>0:
            el = None
            while not el:
                by_tuple = (By.XPATH, f"//div[@id='main']//footer//div[@title='{MESSAGE_LINK_RENDERED_TITLE}']")
                el = self.get_prensented_elements(by_tuple=by_tuple)
                time.sleep(0.1)

    def attach_message_image(self):
        by_tuple = (By.XPATH, f"//div[@id='main']//footer//div[@title='Type a message']")
        el = self.get_clickable_element(by_tuple=by_tuple)
        if el:
            el.click()
            ActionChains(self.browser).key_down(Keys.CONTROL).send_keys('v').perform()

    def click_send(self, send_button=True):
        if send_button:
            by_tuple = (By.XPATH, f"//div[@id='main']//footer//button[@aria-label='Send']")
        else:
            by_tuple = (By.XPATH, f"//div[@id='app']//div[@aria-label='Send']")
        el = self.get_clickable_element(by_tuple)
        if el:
            el.click()
            return True
        return False

    def wait_until_sent(self):
        by_tuple = (By.XPATH, f"//div[@id='app']//span[@aria-label=' Pending ']")
        els = self.get_prensented_elements(by_tuple=by_tuple, timeout=5)
        while len(els)>0:
            print("Message Pending...")
            els = self.get_prensented_elements(by_tuple=by_tuple, timeout=1)
        
    def start_sending_msg(self):
        try:
            if not MESSAGE_BODY_FILE.exists():
                print(f"There's not message body file {MESSAGE_BODY_FILE}")
                return
            message_body = None
            with open(MESSAGE_BODY_FILE, "r", encoding="utf-8") as f:
                message_body = f.read()
            if message_body is None or len(message_body)==0:
                print(f"There's no body in message body file {MESSAGE_BODY_FILE}")
                return
            print("Message body to send: ", message_body)

            if CONTACT_FOLDER_PATH.exists():            
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
                if self.login():
                    WebDriverWait(self.browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, "title")))
                    print("Logged in")
                    file_process_durations = []
                    chunk_process_durations = []
                    row_process_durations = []
                    start_time = time.time()
                    for f in CONTACT_FOLDER_PATH.iterdir():
                        filename = f.name
                        if filename.startswith('~$') or filename=='Contacts Template.xlsx':continue
                        if f.is_file() and f.suffix=='.xlsx':
                            if not confirmation_input(f"Process {filename}?", 'N/y'):continue
                            start_time1 = time.time()
                            print(f"Processing {filename}")
                            sheets = list(pd.read_excel(f, sheet_name=None))
                            print("Total Sheets:", len(sheets))
                            previous_image_name = None 
                            is_image_in_clipboard = False                        
                            sheets_df = {}
                            for sheet in sheets:
                                orginal_df = pd.read_excel(f, sheet_name=sheet)
                                df = deepcopy(orginal_df)
                                try:
                                    has_contact_number_col = CONTACT_NUMBER_COLUMN_NAME in df
                                    has_contact_name_col = CONTACT_NAME_COLUMN_NAME in df
                                    has_image_name_col = IMAGE_NAME_COLUMN_NAME in df
                                    has_status_col = STATUS_COLUMN_NAME in df
                                    if has_contact_number_col and has_contact_name_col and has_image_name_col and has_status_col:
                                        size = roundoff(NUMBER_OF_ROWS_TO_PROCESS)
                                        chunks = chunker(df, size)
                                        no_of_chunks = roundoff(len(df)/NUMBER_OF_ROWS_TO_PROCESS)
                                        if no_of_chunks==0:
                                            no_of_chunks = 1
                                        print("no_of_chunks: ", no_of_chunks, NUMBER_OF_ROWS_TO_PROCESS, size)
                                        for i, chunk in enumerate(chunks):
                                            if not confirmation_input(f"Process chunk {i+1}/{no_of_chunks} of sheet '{sheet}'?", 'N/y'):continue
                                            print(f"Processing chunk {i+1}/{no_of_chunks} of sheet '{sheet}'")
                                            start_time2 = time.time()
                                            for row in chunk.iterrows():
                                                start_time3 = time.time()
                                                status = "Unknown"
                                                idx = row[0]
                                                r = row[1]
                                                contact_number = r[CONTACT_NUMBER_COLUMN_NAME]
                                                contact_name = r[CONTACT_NAME_COLUMN_NAME]
                                                image_name = r[IMAGE_NAME_COLUMN_NAME]
                                                if pd.isnull(contact_number):continue
                                                if pd.isnull(contact_name):
                                                    contact_name = ""
                                                if pd.isnull(image_name):
                                                    image_name = None
                                                    is_image_in_clipboard = False
                                                else:
                                                    image_name = str(image_name).strip()
                                                    if len(image_name)==0:
                                                        is_image_in_clipboard = False
                                                    elif previous_image_name!=image_name:
                                                        is_image_in_clipboard = send_image_to_clipboard(
                                                            image_dir=MESSAGE_IMAGE_DIR, 
                                                            image_name=image_name
                                                        )
                                                        if is_image_in_clipboard:
                                                            previous_image_name = image_name
                                                contact_number = str(contact_number).strip()
                                                contact_name = str(contact_name).strip()
                                                if not contact_number.startswith(f"+{COUNTRY_CODE}"):
                                                    contact_number = f"+{COUNTRY_CODE}{contact_number.lstrip(COUNTRY_CODE)}"
                                                print(f"Processing {contact_name}, {contact_number}")
                                                try:
                                                    text = urllib.parse.urlencode({'text' : message_body.format(contact_name=contact_name)})
                                                    send_url = SEND_URL % (contact_number, text)
                                                    if self.get_page(send_url, LOGIN_TITLE):
                                                        if is_image_in_clipboard:
                                                            self.attach_message_image()
                                                        else:
                                                            self.is_message_link_rendered()
                                                        if self.click_send(send_button=not is_image_in_clipboard):
                                                            self.wait_until_sent()
                                                            print(f"######################## SENT TO: {contact_name}, {contact_number} ########################")
                                                            status = "Success"
                                                        else:
                                                            print(f"######################## Falied to SENT TO: {contact_name}, {contact_number} ########################")
                                                            status = "Fail"
                                                except Exception as e:
                                                    print("WhatsAppSendMsg.start_sending_msg Error1: ", e, traceback.format_exc())
                                                    status = "Fail"
                                                orginal_df[STATUS_COLUMN_NAME][idx] = status
                                                row_process_duration = time.time()-start_time3
                                                row_process_durations.append(row_process_duration)
                                            chunk_process_duration = time.time()-start_time2
                                            chunk_process_durations.append(chunk_process_duration)
                                            print(f"Processing of chunk {i+1}/{no_of_chunks} of sheet '{sheet}' took: {chunk_process_duration} seconds.")
                                    else:
                                        if not has_contact_number_col:
                                            raise Exception(f"{filename} has missing column: '{CONTACT_NUMBER_COLUMN_NAME}'")
                                        elif not has_contact_name_col:
                                            raise Exception(f"{filename} has missing column: '{CONTACT_NAME_COLUMN_NAME}'")
                                        elif not has_image_name_col:
                                            raise Exception(f"{filename} has missing column: '{IMAGE_NAME_COLUMN_NAME}'")
                                        elif not has_status_col:
                                            raise Exception(f"{filename} has missing column: '{STATUS_COLUMN_NAME}'")      
                                except Exception as e:
                                    print("WhatsAppSendMsg.start_sending_msg Error2 : ", e, traceback.format_exc())
                                sheets_df[sheet] = orginal_df
                            file_process_duration = time.time() - start_time1
                            file_process_durations.append(file_process_duration)
                            print(f"Processing of filename '{filename}' took: {file_process_duration} seconds.")

                            if confirmation_input(f"\n###################### Please close '{filename}' if it's opened anywhere. ######################\nClosed?", 'N/y'):
                                with pd.ExcelWriter(f, engine="openpyxl", mode="w") as writer:
                                    for sheet, df in sheets_df.items():
                                        df.to_excel(writer, sheet_name=sheet, index=False)

                    print(f"""Row processing duration metrics:
                          Max: {max(row_process_durations)}
                          Min: {min(row_process_durations)}
                          Avg: {np.average(row_process_durations)}
                          Mean: {np.mean(row_process_durations)}""")
                    print(f"""Chunk processing duration metrics:
                          Max: {max(chunk_process_durations)}
                          Min: {min(chunk_process_durations)}
                          Avg: {np.average(chunk_process_durations)}
                          Mean: {np.mean(chunk_process_durations)}""")
                    print(f"""File processing duration metrics:
                          Max: {max(file_process_durations)}
                          Min: {min(file_process_durations)}
                          Avg: {np.average(file_process_durations)}
                          Mean: {np.mean(file_process_durations)}""")
                    print(f"Total processing time: {time.time() -start_time} seconds.")                
                    print("######################## COMPLETED ########################")
                else:
                    print("Couldn't log in !!!")
        except Exception as e:
            print("WhatsAppSendMsg.start_sending_msg Error2: ", e, traceback.format_exc())
        # self.kill_browser_process()
            


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

    parser.add_argument(
        "-d",
        "--debug",
        dest="DEBUG",
        action='store_true',
        default=False,
        required=False,
        help="Run script with debug mode, default: off"
    )

    args = parser.parse_args(argv[1:])
    print(parser.description)
    INVISIBLE = args.INVISIBLE
    DEBUG = args.DEBUG
    print("INVISIBLE: ", INVISIBLE)
    print("DEBUG: ", DEBUG)
    wp_scraper = WhatsAppSendMsg(invisible=INVISIBLE, debug=DEBUG)
    wp_scraper.start_sending_msg()
