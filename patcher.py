from distutils.version import LooseVersion
import io
import json
import logging
import os
from pathlib import Path
import random
import re
import shutil
import string
import sys
import time
from urllib.request import urlopen
from urllib.request import urlretrieve
import zipfile
from multiprocessing import Lock
from settings import CHROME_DIR


IS_POSIX = sys.platform.startswith(("darwin", "cygwin", "linux", "linux2"))


class Patcher(object):
    lock = Lock()
    exe_name = "chromedriver%s"

    platform = sys.platform
    data_path = CHROME_DIR
    data_path.mkdir(exist_ok=True)

    def __init__(
        self,
        executable_path=None,
        force=False,
        user_multi_procs=False,
    ):
        self.force = force
        self._custom_exe_path = False
        prefix = "undetected"
        self.user_multi_procs = user_multi_procs

        # Needs to be called before self.exe_name is accessed
        self._set_platform_name()

        if not os.path.exists(self.data_path):
            os.makedirs(self.data_path, exist_ok=True)

        if not executable_path:
            self.executable_path = os.path.join(
                self.data_path, "_".join([prefix, self.exe_name])
            )

        if not IS_POSIX:
            if executable_path:
                if not executable_path[-4:] == ".exe":
                    executable_path += ".exe"

        self.zip_path = os.path.join(self.data_path, prefix)

        if not executable_path:
            if not self.user_multi_procs:
                self.executable_path = os.path.abspath(
                    os.path.join(".", self.executable_path)
                )

        if executable_path:
            self._custom_exe_path = True
            self.executable_path = executable_path

        self.url_repo = "https://googlechromelabs.github.io/chrome-for-testing"

    def _set_platform_name(self):
        """
        Set the platform and exe name based on the platform undetected_chromedriver is running on
        in order to download the correct chromedriver.
        """
        if self.platform.endswith("win32"):
            self.platform_name = "win32"
            self.exe_name %= ".exe"
        if self.platform.endswith(("linux", "linux2")):
            self.platform_name = "linux64"
            self.exe_name %= ""
        if self.platform.endswith("darwin"):
            self.platform_name = "mac-x64"
            self.exe_name %= ""
        # self.platform_name = "linux64"
        # self.exe_name %= ""

    def auto(self, executable_path=None, force=False, _=None):
        p = Path(self.data_path)
        if self.user_multi_procs:
            with Lock():
                if IS_POSIX:
                    driver_name = "*chromedriver"
                else:
                    driver_name = "*chromedriver.exe"
                files = list(p.rglob(driver_name))
                if len(files)>0:
                    most_recent = max(files, key=lambda f: f.stat().st_mtime)
                    files.remove(most_recent)
                    list(map(lambda f: f.unlink(), files))
                    if self.is_binary_patched(most_recent):
                        self.executable_path = str(most_recent)
                        os.chmod(self.executable_path, mode=0o755)
                        return True
        print("executable_path: ", executable_path)
        
        if executable_path and executable_path.exists():
            self.executable_path = executable_path
            self._custom_exe_path = True

        if self._custom_exe_path:
            ispatched = self.is_binary_patched(self.executable_path)
            print("ispatched: ",ispatched)
            if not ispatched:
                os.chmod(self.executable_path, mode=0o755)
                return self.patch_exe()
            else:
                os.chmod(self.executable_path, mode=0o755)
                return
            
        if force is True:
            self.force = force

        try:
            os.unlink(self.executable_path)
        except PermissionError:
            if self.force:
                self.force_kill_instances(self.executable_path)
                return self.auto(force=not self.force)
            try:
                if self.is_binary_patched():
                    # assumes already running AND patched
                    return True
            except PermissionError:
                pass
            # return False
        except FileNotFoundError:
            pass

        release = self.fetch_release_number()
        self.version_full = release
        self.unzip_package(self.fetch_package())
        return self.patch()

    def driver_binary_in_use(self, path: str = None) -> bool:
        if not path:
            path = self.executable_path
        p = Path(path)

        if not p.exists():
            raise OSError("file does not exist: %s" % p)
        try:
            with open(p, mode="a+b") as fs:
                exc = []
                try:

                    fs.seek(0, 0)
                except PermissionError as e:
                    exc.append(e)  # since some systems apprently allow seeking
                    # we conduct another test
                try:
                    fs.readline()
                except PermissionError as e:
                    exc.append(e)

                if exc:

                    return True
                return False
            # ok safe to assume this is in use
        except Exception as e:
            # print("whoops ", e)
            pass

    def cleanup_unused_files(self):
        p = Path(self.data_path)
        items = list(p.glob("*undetected*"))
        for item in items:
            try:
                item.unlink()
            except:
                pass

    def patch(self):
        self.patch_exe()
        return self.is_binary_patched()

    def fetch_release_number(self):
        path = "/last-known-good-versions-with-downloads.json"
        print("getting release number from %s" % path)
        with urlopen(self.url_repo + path) as conn:
            response = conn.read().decode()

        last_versions = json.loads(response)
        return LooseVersion(last_versions["channels"]["Stable"]["version"])

    def parse_exe_version(self):
        with io.open(self.executable_path, "rb") as f:
            for line in iter(lambda: f.readline(), b""):
                match = re.search(rb"platform_handle\x00content\x00([0-9.]*)", line)
                if match:
                    return LooseVersion(match[1].decode())

    def fetch_package(self):
        zip_name = f"chromedriver_{self.platform_name}.zip"
        zip_name = zip_name.replace("_", "-", 1)
        download_url = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/%s/%s/%s"
        download_url %= (self.version_full.vstring, self.platform_name, zip_name)
        print("downloading from %s" % download_url)
        return urlretrieve(download_url)[0]

    def unzip_package(self, fp):
        exe_path = self.exe_name
        zip_name = f"chromedriver-{self.platform_name}"
        exe_path = os.path.join(zip_name, self.exe_name)
        print("unzipping %s" % fp)
        try:
            os.unlink(self.zip_path)
        except (FileNotFoundError, OSError):
            pass

        os.makedirs(self.zip_path, mode=0o755, exist_ok=True)
        with zipfile.ZipFile(fp, mode="r") as zf:
            zf.extractall(self.zip_path)
        os.rename(os.path.join(self.zip_path, exe_path), self.executable_path)
        os.remove(fp)
        shutil.rmtree(self.zip_path)
        os.chmod(self.executable_path, 0o755)
        return self.executable_path

    @staticmethod
    def force_kill_instances(exe_name):
        exe_name = os.path.basename(exe_name)
        if IS_POSIX:
            r = os.system("kill -f -9 $(pidof %s)" % exe_name)
        else:
            r = os.system("taskkill /f /im %s" % exe_name)
        return not r

    @staticmethod
    def gen_random_cdc():
        cdc = random.choices(string.ascii_letters, k=27)
        return "".join(cdc).encode()

    def is_binary_patched(self, executable_path=None):
        executable_path = executable_path or self.executable_path
        try:
            with io.open(executable_path, "rb") as fh:
                return fh.read().find(b"undetected chromedriver") != -1
        except FileNotFoundError:
            return False

    def patch_exe(self):
        start = time.perf_counter()
        print("patching driver executable %s" % self.executable_path)
        with io.open(self.executable_path, "r+b") as fh:
            content = fh.read()
            # match_injected_codeblock = re.search(rb"{window.*;}", content)
            match_injected_codeblock = re.search(rb"\{window\.cdc.*?;\}", content)
            if match_injected_codeblock:
                target_bytes = match_injected_codeblock[0]
                new_target_bytes = (
                    b'{console.log("undetected chromedriver 1337!")}'.ljust(
                        len(target_bytes), b" "
                    )
                )
                new_content = content.replace(target_bytes, new_target_bytes)
                if new_content == content:
                    print(
                        "something went wrong patching the driver binary. could not find injection code block"
                    )
                else:
                    print(
                        "found block:\n%s\nreplacing with:\n%s"
                        % (target_bytes, new_target_bytes)
                    )
                fh.seek(0)
                fh.write(new_content)
        print(
            "patching took us {:.2f} seconds".format(time.perf_counter() - start)
        )

    def __repr__(self):
        return "{0:s}({1:s})".format(
            self.__class__.__name__,
            self.executable_path,
        )

    def __del__(self):
        if self._custom_exe_path:
            # if the driver binary is specified by user
            # we assume it is important enough to not delete it
            return
        else:
            timeout = 3  # stop trying after this many seconds
            t = time.monotonic()
            now = lambda: time.monotonic()
            while now() - t > timeout:
                # we don't want to wait until the end of time
                try:
                    if self.user_multi_procs:
                        break
                    os.unlink(self.executable_path)
                    print("successfully unlinked %s" % self.executable_path)
                    break
                except (OSError, RuntimeError, PermissionError):
                    time.sleep(0.01)
                    continue
                except FileNotFoundError:
                    break
