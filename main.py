# main file for the scraping
# @author: Sk Khurshid Alam

import os
import sys
import signal
import psutil
from contextlib import suppress
import argparse
import time
import random
import tempfile
import shutil
import pandas as pd
import urllib
from weakref import finalize
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.service import utils
from selenium.webdriver.chromium.service import ChromiumService
from selenium.webdriver.chromium.options import ChromiumOptions
from dprocess import start_detached
from patcher import Patcher
from settings import *
import traceback



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
    def __init__(self, invisible=True, debug=False) -> None:
        finalize(self, self.kill_browser_process)
        self.invisible = invisible
        self.debug = debug
        self.not_ok = 0
        self.retry = 0
        self.max_retries = 3
        self.browser_pid = None

    def is_head_ready(self):
        ready = False
        try:
            WebDriverWait(self.browser, 20).until(EC.presence_of_element_located((By.TAG_NAME, "head")))
            ready = self.browser.find_element(By.TAG_NAME, "head") is not None 
        except Exception as e:
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
            except:
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
                self.browser.service.process.kill()
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

        try:
            if self.browser_pid is not None:
                os.kill(self.browser_pid, signal.SIGTERM)
        except:
            pass
        finally:
            try:
                if self.browser_pid is not None:
                    from dprocess import REGISTERED
                    if self.browser_pid in REGISTERED:
                        REGISTERED.remove(self.browser_pid)
            except:
                pass        

        if hasattr(self, "browser"):
            if self.browser is None or not hasattr(self.browser, "service") or self.browser.service:
                print("Browser closed and webdriver process killed!")
                self.browser = None
            else:
                print("Browser and Webdriver process NOT killed !!!!")


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
        user_data_dir = os.path.normpath(tempfile.mkdtemp())
        options.add_argument(f"user-data-dir={user_data_dir}")
        # options.add_experimental_option("useAutomationExtension", False)
        # options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        os.environ["webdriver.chrome.driver"] = str(chrome_driver_path.absolute())
        
        desired_capabilities = options.to_capabilities()
        # print("desired_capabilities: ", desired_capabilities)

        binary_location = self.find_chrome_executable()
        if binary_location is not None:
            options.binary_location = binary_location
            self.browser_pid = start_detached(
                options.binary_location, *options.arguments
            )

        service = ChromiumService(executable_path=chrome_driver_path)
        self.browser = webdriver.chrome.webdriver.WebDriver(service=service, options=options, keep_alive=True)
        self.browser._delay = 3
        self.browser.user_data_dir = user_data_dir
        self.browser.keep_user_data_dir = False
        if self.invisible:
            self._configure_headless()
        # self.browser.set_page_load_timeout(self.page_load_timeout)
        self.browser.maximize_window()
        print("browserVersion: ", self.browser.capabilities["browserVersion"])
        print("chromedriverVersion: ", self.browser.capabilities["chrome"]["chromedriverVersion"].split(" ")[0])
        if not self.test_browser_ok():
            self.retry += 1
            if self.retry<=self.max_retries:
                self.config_browser()
            else:
                raise Exception("Failed to configure browser.")
        else:
            self.retry = 0


    def get_prensented_elements(self, by_tuple):
        els = []
        try:
            els = WebDriverWait(self.browser, 10).until(EC.presence_of_all_elements_located(by_tuple))
        except Exception as e:
            print(f"WhatsAppSendMsg.get_clickable_element {e}")
            els = []
        return els

    def get_clickable_element(self, by_tuple):
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
            el1 = WebDriverWait(self.browser, 20).until(EC.element_to_be_clickable(by_tuple))
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
        
    def click_send(self):
        by_tuple = (By.XPATH, f"//div[@id='main']//footer//button[@aria-label='Send']")
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
                            print(df)
                            try:
                                print(CONTACT_NUMBER_COLUMN_NAME, CONTACT_NUMBER_COLUMN_NAME in df)
                                print(CONTACT_NAME_COLUMN_NAME, CONTACT_NAME_COLUMN_NAME in df)
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
