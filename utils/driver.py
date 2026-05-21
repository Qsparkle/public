"""
ChromeDriver 管理 + Selenium 工具。
get_chromedriver_path 使用 @st.cache_resource，全局只下载一次。
"""
import os
import glob
import shutil
import time
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

XHS_LOGIN_URL = "https://www.xiaohongshu.com/"


def _find_cached_chromedriver() -> str:
    """
    搜索 webdriver_manager 已下载到本地的 ChromeDriver，无需联网。
    返回最新修改的可执行文件路径，找不到返回空字符串。
    """
    exe = "chromedriver.exe" if os.name == "nt" else "chromedriver"
    wdm_cache = os.path.expanduser("~/.wdm/drivers/chromedriver")
    if os.path.isdir(wdm_cache):
        matches = [
            p for p in glob.glob(os.path.join(wdm_cache, "**", exe), recursive=True)
            if os.path.isfile(p)
        ]
        if matches:
            return max(matches, key=os.path.getmtime)  # 取最新的
    return ""


@st.cache_resource
def get_chromedriver_path() -> str:
    """
    获取 ChromeDriver 路径，全局只运行一次。
    优先顺序：本地缓存 > PATH > 联网下载
    """
    os.environ.setdefault("WDM_PROGRESS_BAR", "0")
    os.environ.setdefault("WDM_SSL_VERIFY", "0")

    # 1. 先找 webdriver_manager 的本地缓存（无需联网）
    cached = _find_cached_chromedriver()
    if cached:
        return cached

    # 2. 系统 PATH
    path = shutil.which("chromedriver") or shutil.which("chromedriver.exe")
    if path:
        return path

    # 3. Windows 常用目录
    for p in [
        r"C:\chromedriver\chromedriver.exe",
        r"C:\Program Files\chromedriver\chromedriver.exe",
    ]:
        if os.path.exists(p):
            return p

    # 4. 最后才联网下载（中国网络可能失败）
    try:
        return ChromeDriverManager().install()
    except Exception as e:
        raise RuntimeError(
            f"ChromeDriver 未找到（{e}）。\n"
            "解决方案：\n"
            "1. 开启全局代理后重启应用\n"
            "2. 手动下载 ChromeDriver 放入 C:\\chromedriver\\chromedriver.exe\n"
            "   下载地址：https://googlechromelabs.github.io/chrome-for-testing/"
        )


def init_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--ignore-certificate-errors")   # 忽略 SSL 证书错误
    options.add_argument("--allow-running-insecure-content")
    # 注意：--single-process 在 Windows 上会导致崩溃，只在 Linux 容器中启用
    if os.path.isdir("/usr/bin"):
        options.add_argument("--single-process")

    # Streamlit Cloud 上 Chromium 可能在不同路径，依次尝试
    _chromium_candidates = [
        ("/usr/bin/chromium", "/usr/bin/chromedriver"),
        ("/usr/bin/chromium-browser", "/usr/bin/chromedriver"),
        ("/usr/bin/chromium-browser", "/usr/bin/chromium-driver"),
    ]
    linux_service = None
    for _chrom, _driver in _chromium_candidates:
        if os.path.exists(_chrom) and os.path.exists(_driver):
            options.binary_location = _chrom
            linux_service = Service(executable_path=_driver)
            break

    if linux_service:
        # Linux 云端环境
        service = linux_service
    else:
        # 本地环境
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        service = Service(executable_path=get_chromedriver_path())

    try:
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        err = str(e)
        if "executable needs to be in PATH" in err or "chromedriver" in err.lower():
            raise RuntimeError(
                "ChromeDriver 启动失败。请确认：\n"
                "1. 本机已安装 Google Chrome 浏览器\n"
                "2. Chrome 版本与 ChromeDriver 版本匹配\n"
                f"原始错误：{err}"
            )
        raise

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
    })
    driver.set_page_load_timeout(30)
    return driver


def add_cookies_to_driver(driver: webdriver.Chrome, cookies: list, url: str = XHS_LOGIN_URL):
    try:
        driver.get(url)
    except Exception as e:
        err = str(e)
        if "timeout" not in err.lower():
            raise RuntimeError(
                f"浏览器导航失败（{err}）。\n"
                "请检查：网络连接是否正常、是否需要代理/VPN 访问小红书。"
            )

    time.sleep(2)

    # 检测 Chrome 是否静默加载了错误页（不抛异常但页面是 chrome-error://）
    try:
        current = driver.current_url
        if current.startswith("chrome-error://") or current.startswith("data:"):
            raise RuntimeError(
                f"浏览器无法访问小红书（落地页：{current}）。\n"
                "请检查：\n"
                "1. 网络连接是否正常\n"
                "2. 是否需要代理/VPN 才能访问小红书\n"
                "3. 尝试在普通 Chrome 中打开小红书确认可以访问"
            )
    except RuntimeError:
        raise
    except Exception:
        pass

    for ck in cookies:
        try:
            if "name" not in ck or "value" not in ck:
                continue
            driver.add_cookie(ck)
        except Exception:
            continue
