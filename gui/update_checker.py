from packaging import version
import urllib.request
import urllib.parse
import logging
import ctypes
import json
import sys
import re


logger = logging.getLogger(__name__)


class VS_FIXEDFILEINFO(ctypes.Structure):
    _fields_ = [
        ("dwSignature", ctypes.c_uint32),
        ("dwStrucVersion", ctypes.c_uint32),
        ("dwFileVersionMS", ctypes.c_uint32),
        ("dwFileVersionLS", ctypes.c_uint32),
        ("dwProductVersionMS", ctypes.c_uint32),
        ("dwProductVersionLS", ctypes.c_uint32),
        ("dwFileFlagsMask", ctypes.c_uint32),
        ("dwFileFlags", ctypes.c_uint32),
        ("dwFileOS", ctypes.c_uint32),
        ("dwFileType", ctypes.c_uint32),
        ("dwFileSubtype", ctypes.c_uint32),
        ("dwFileDateMS", ctypes.c_uint32),
        ("dwFileDateLS", ctypes.c_uint32),
    ]


class UpdateChecker:
    def __init__(self, github_repo: str, exe_path: str = sys.executable):
        self.github_repo = github_repo  # :param github_repo: GitHub repository in the format "owner/repo",

    def check_app_update_status(self, file_version: str | None = None) -> bool | None:
        if not file_version:
            file_version, _ = self._get_local_version()
        latest_online_version  = self._get_latest_github_version()

        logging.debug(f"Obecna wersja zainstalowana: {file_version}", f"Najnowsza wersja dostępna online: {latest_online_version}")

        try:
            if latest_online_version:
                latest_online_version = version.parse(latest_online_version)
            file_version = version.parse(file_version)
        except Exception as e:
            logging.exception(e)
            return None

        if latest_online_version is None:
            logging.debug("Offline")
            return "Offline"
        elif latest_online_version == file_version:
            logging.debug("The installed version is the same as the version available online.")
            return False  # Installed version is up to date.
        elif latest_online_version > file_version:
            logging.debug("The installed version is outdated.")
            return True  # Installed version is outdated.
        elif latest_online_version < file_version:
            logging.debug("The installed version is newer than the version available online.")
            return False  # Installed version is up to date.
        else:
            logging.debug("Error: app version check")
            return None  # Error

    def _get_local_version(self, exe_path=sys.executable):
        size = ctypes.windll.version.GetFileVersionInfoSizeW(exe_path, None)
        if size == 0:
            return None, None

        res = ctypes.create_string_buffer(size)
        ctypes.windll.version.GetFileVersionInfoW(exe_path, None, size, res)

        r = ctypes.c_void_p()
        l = ctypes.c_uint()

        # Query the root block for the VS_FIXEDFILEINFO structure
        ctypes.windll.version.VerQueryValueW(res, '\\', ctypes.byref(r), ctypes.byref(l))

        info = ctypes.cast(r, ctypes.POINTER(VS_FIXEDFILEINFO)).contents

        def print_vs_fixedfileinfo(info):
            print("VS_FIXEDFILEINFO:")
            print(f"  dwFileVersionMS: {info.dwFileVersionMS:#010x}")
            print(f"  dwFileVersionLS: {info.dwFileVersionLS:#010x}")
            print(f"  dwProductVersionMS: {info.dwProductVersionMS:#010x}")
            print(f"  dwProductVersionLS: {info.dwProductVersionLS:#010x}")
        #print_vs_fixedfileinfo(info)

        # Extract file version numbers
        file_version_ms = info.dwFileVersionMS
        file_version_ls = info.dwFileVersionLS
        file_version = (file_version_ms >> 16, file_version_ms & 0xffff, file_version_ls >> 16, file_version_ls & 0xffff)
        file_version = ".".join(map(str, file_version)) if file_version else "Unknown file version"

        # Extract product version numbers
        product_version_ms = info.dwProductVersionMS
        product_version_ls = info.dwProductVersionLS
        product_version = (product_version_ms >> 16, product_version_ms & 0xffff, product_version_ls >> 16, product_version_ls & 0xffff)
        product_version = ".".join(map(str, product_version)) if product_version else "Unknown file version"

        return file_version, product_version

    def _get_latest_github_version(self) -> str | None:
        api_url = f"https://api.github.com/repos/{self.github_repo}/releases/latest"

        request = urllib.request.Request(api_url, headers={"Accept": "application/json"})

        try:
            with urllib.request.urlopen(request) as response:
                if response.status == 200:
                    data = json.loads(response.read().decode())
                    return data["tag_name"]
                else:
                    print(f"Request failed with status code: {response.status}")
                    return False
        except urllib.error.URLError as e:
            logging.exception(e)
            print(f"Failed to reach the server. Reason: {e.reason}")
            return None


if __name__ == "__main__":
    pass