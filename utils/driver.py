"""
ChromeDriver 管理 + Selenium 工具。
get_chromedriver_path 使用 @st.cache_resource，全局只下载一次。
包含：warmup_browser、human_like_drag 等拡展函数。
"""
import os
import glob
import math
import random
import shutil
import time
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

XHS_LOGIN_URL = "https://www.xiaohongshu.com/"

# ===== 全面的反自动化检测 JS =====
_ANTI_DETECT_JS = """
(function() {
    // 1. 覆盖 webdriver 检测
    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    
    // 2. 覆盖 plugins（headless 下 plugins.length 为 0）
    Object.defineProperty(navigator, 'plugins', {
        get: () => [1, 2, 3, 4, 5]
    });
    
    // 3. 覆盖 languages
    Object.defineProperty(navigator, 'languages', {
        get: () => ['zh-CN', 'zh', 'en']
    });
    
    // 4. 覆盖 chrome.runtime
    if (!window.chrome) {
        window.chrome = { runtime: {} };
    }
    
    // 5. 覆盖 permissions API
    try {
        const originalQuery = window.navigator.permissions.query;
        window.navigator.permissions.query = function(params) {
            if (params.name === 'notifications') {
                return Promise.resolve({ state: 'granted' });
            }
            return originalQuery(params);
        };
    } catch(e) {}
    
    // 6. 覆盖 WebGL 指纹（避免暴露虚拟机/headless 渲染器）
    try {
        const getParameter = WebGLRenderingContext.prototype.getParameter;
        WebGLRenderingContext.prototype.getParameter = function(p) {
            if (p === 37445) return 'Intel Inc.';
            if (p === 37446) return 'Intel Iris OpenGL Engine';
            if (p === 7936) return 'WebKit WebGL';
            if (p === 7937) return 'Mozilla';
            return getParameter(p);
        };
    } catch(e) {}
    
    // 7. 覆盖 Canvas 指纹（制造微小随机噪点，防止被追踪）
    try {
        const originalToDataURL = HTMLCanvasElement.prototype.toDataURL;
        HTMLCanvasElement.prototype.toDataURL = function(type) {
            const ctx = this.getContext('2d');
            if (ctx) {
                const imageData = ctx.getImageData(0, 0, this.width, this.height);
                for (let i = 0; i < imageData.data.length; i += 4) {
                    imageData.data[i] = imageData.data[i] ^ 1;
                }
                ctx.putImageData(imageData, 0, 0);
            }
            return originalToDataURL.apply(this, arguments);
        };
    } catch(e) {}
    
    // 8. 覆盖 navigator.hardwareConcurrency（headless 可能返回特殊值）
    Object.defineProperty(navigator, 'hardwareConcurrency', {
        get: () => Math.floor(Math.random() * 4) + 4
    });
    
    // 9. 覆盖 navigator.deviceMemory
    Object.defineProperty(navigator, 'deviceMemory', {
        get: () => 8
    });
    
    // 10. 覆盖 screen 属性
    try {
        const fakeScreen = {
            width: window.screen.width,
            height: window.screen.height,
            availWidth: window.screen.availWidth,
            availHeight: window.screen.availHeight,
            colorDepth: 24,
            pixelDepth: 24
        };
        Object.defineProperty(window.screen, 'colorDepth', { get: () => 24 });
        Object.defineProperty(window.screen, 'pixelDepth', { get: () => 24 });
    } catch(e) {}
    
    // 11. 覆盖 navigator.connection（模拟真实网络环境）
    try {
        Object.defineProperty(navigator, 'connection', {
            get: () => ({
                effectiveType: '4g',
                downlink: 10,
                rtt: 50,
                saveData: false
            })
        });
    } catch(e) {}
})();
"""


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


def init_driver(window_size: tuple = None) -> webdriver.Chrome:
    """
    初始化 Chrome 浏览器驱动。
    
    Args:
        window_size: (width, height) 可选，不提供则随机生成相近尺寸。
    """
    options = Options()
    
    # ----- headless 模式 -----
    # 使用 Chrome 112+ 的新无头模式（--headless=new），比旧版更不易被检测
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    # ----- 窗口尺寸随机化（避免总是 1920x1080 被识别）-----
    if window_size:
        w, h = window_size
    else:
        # 随机在几个常见分辨率中选择，避免总是相同尺寸
        sizes = [
            (1920, 1080), (1920, 1200), (1366, 768),
            (1440, 900), (1536, 864), (1600, 900),
            (1280, 720), (1280, 800), (1920, 1080),
        ]
        base_w, base_h = random.choice(sizes)
        # 微调 ±30px，避免每次完全一样
        w = base_w + random.randint(-30, 30)
        h = base_h + random.randint(-30, 30)
    options.add_argument(f"--window-size={w},{h}")
    
    # ----- User-Agent（真实 Windows Chrome 120-126 随机）-----
    chrome_versions = [120, 121, 122, 123, 124, 125, 126]
    ver = random.choice(chrome_versions)
    ua = (
        f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        f"AppleWebKit/537.36 (KHTML, like Gecko) "
        f"Chrome/{ver}.0.0.0 Safari/537.36"
    )
    options.add_argument(f"user-agent={ua}")
    
    # ----- 反检测 -----
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-running-insecure-content")
    
    # 禁用自动化特征
    options.add_argument("--disable-automation")
    
    # 禁用部分可能暴露 headless 的 feature
    options.add_argument("--disable-features=VizDisplayCompositor")
    
    # 设置语言
    options.add_argument("--lang=zh-CN")
    
    # ----- 实验性选项 -----
    prefs = {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    options.add_experimental_option("prefs", prefs)
    
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

    # 注入全面的反检测脚本
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": _ANTI_DETECT_JS
    })
    
    driver.set_page_load_timeout(30)
    return driver


def warmup_browser(driver: webdriver.Chrome, url: str = "https://www.douyin.com/"):
    """
    浏览器预热：模拟人类浏览行为，降低首次请求触发验证码的概率。
    在正式截图前调用一次即可。
    """
    try:
        driver.get(url)
        time.sleep(3)
        
        # 模拟缓慢滚动
        for _ in range(random.randint(2, 4)):
            scroll_y = random.randint(200, 600)
            driver.execute_script(f"window.scrollBy(0, {scroll_y})")
            time.sleep(random.uniform(0.8, 1.8))
        
        # 模拟鼠标随机移动
        ac = ActionChains(driver)
        ac.move_by_offset(random.randint(100, 500), random.randint(100, 500))
        ac.pause(random.uniform(0.3, 0.6))
        ac.perform()
        
        time.sleep(random.uniform(1, 2))
    except Exception:
        pass  # 预热失败不影响主流程


def add_cookies_to_driver(driver: webdriver.Chrome, cookies: list, url: str = XHS_LOGIN_URL):
    # 推断当前平台，用于错误提示和域名修正
    is_douyin = "douyin.com" in url
    platform_name = "抖音" if is_douyin else "小红书"

    try:
        driver.get(url)
    except Exception as e:
        err = str(e)
        if "timeout" not in err.lower():
            raise RuntimeError(
                f"浏览器导航失败（{err}）。\n"
                f"请检查：网络连接是否正常、是否需要代理/VPN 访问{platform_name}。"
            )

    time.sleep(2)

    # 检测 Chrome 是否静默加载了错误页（不抛异常但页面是 chrome-error://）
    try:
        current = driver.current_url
        if current.startswith("chrome-error://") or current.startswith("data:"):
            raise RuntimeError(
                f"浏览器无法访问{platform_name}（落地页：{current}）。\n"
                f"请检查：\n"
                f"1. 网络连接是否正常\n"
                f"2. 是否需要代理/VPN 才能访问{platform_name}\n"
                f"3. 尝试在普通 Chrome 中打开{platform_name}确认可以访问"
            )
    except RuntimeError:
        raise
    except Exception:
        pass

    for ck in cookies:
        try:
            if "name" not in ck or "value" not in ck:
                continue
            # 修正拖音 Cookie domain：将 www.douyin.com 统一修正为 .douyin.com
            # 不带点的根域名才能在子域名页面生效
            if is_douyin and "domain" in ck:
                domain = ck["domain"]
                if domain and not domain.startswith(".") and "douyin.com" in domain:
                    ck = dict(ck)  # 不修改原对象
                    ck["domain"] = ".douyin.com"
            driver.add_cookie(ck)
        except Exception:
            continue


def bezier_curve(start: int, end: int, steps: int = 35) -> list:
    """
    生成贝塞尔曲线路径点。
    模拟人类拖动鼠标的加速/减速/微抖自然轨迹。
    
    Args:
        start: 起始 X 坐标（通常为 0）
        end: 目标 X 坐标（滑块需要移动的距离）
        steps: 路径点数量
        
    Returns:
        [(x, y), ...] 路径点列表
    """
    if steps < 10:
        steps = 30
    
    # 控制点：贝塞尔曲线使用 3 个控制点来模拟自然加速/减速
    # 控制点 1：略微超出目标（overshoot），模拟人类的手部惯性
    overshoot = int(end * random.uniform(0.08, 0.15))
    
    # 随机生成控制点偏移
    cp1_x = int(end * random.uniform(0.2, 0.35))
    cp1_y = random.randint(-8, 8)
    
    cp2_x = int(end * random.uniform(0.6, 0.8))
    cp2_y = random.randint(-5, 5)
    
    # 终点 Y 轴的微小偏移（模拟手部松脱前的抖动）
    end_y = random.randint(-3, 3)
    
    points = []
    for i in range(steps + 1):
        t = i / steps
        
        # 三次贝塞尔曲线公式
        x = (1-t)**3 * start + 3*(1-t)**2*t * cp1_x + 3*(1-t)*t**2 * cp2_x + t**3 * (end + overshoot)
        y = 3*(1-t)**2*t * cp1_y + 3*(1-t)*t**2 * cp2_y + t**3 * end_y
        
        # 在中间阶段添加更多随机微抖（模拟人类手部震颤）
        if 0.2 < t < 0.8:
            x += random.gauss(0, 1.2)
            y += random.gauss(0, 0.8)
        
        points.append((max(0, int(x)), int(y)))
    
    return points


def human_like_drag(driver, slider, target_x: int) -> bool:
    """
    拟人化拖动滑块。
    使用贝塞尔曲线 + 变加速度 + 随机微抖 + 尾部落差复位。
    
    Returns:
        True 表示拖动过程无异常
    """
    try:
        ac = ActionChains(driver)
        
        # 1. 先移动到滑块位置（自然的鼠标移动）
        ac.move_to_element(slider)
        ac.pause(random.uniform(0.3, 0.8))  # 反应时间
        
        # 2. 按住滑块
        ac.click_and_hold(slider)
        ac.pause(random.uniform(0.08, 0.25))  # 刚按住时的犹豫
        
        # 3. 生成贝塞尔曲线路径
        points = bezier_curve(0, target_x, steps=random.randint(30, 50))
        
        prev_x = 0
        for i, (x, y) in enumerate(points):
            dx = x - prev_x
            if dx <= 0 and i > 0:
                continue  # 不允许后退（人类拖动滑块不会后退）
            
            # 加入微颤（高频率小幅度抖动）
            jitter_x = random.gauss(0, 0.3)
            jitter_y = random.gauss(0, 0.4)
            
            ac.move_by_offset(max(1, int(dx + jitter_x)), int(y + jitter_y))
            
            # 根据进度动态调整停顿时间
            progress = i / len(points)
            if progress < 0.1:
                pause = random.uniform(0.008, 0.015)  # 起步加速
            elif progress < 0.3:
                pause = random.uniform(0.006, 0.012)  # 加速阶段
            elif progress < 0.7:
                pause = random.uniform(0.01, 0.02)    # 匀速阶段
            elif progress < 0.85:
                pause = random.uniform(0.015, 0.03)   # 减速接近目标
            else:
                pause = random.uniform(0.03, 0.06)    # 微调阶段
            
            ac.pause(pause)
            prev_x = x
        
        # 4. 松手前短暂停顿（人类会确认位置再松手）
        ac.pause(random.uniform(0.12, 0.3))
        ac.release()
        
        # 5. 执行拖动
        ac.perform()
        return True
    except Exception:
        return False