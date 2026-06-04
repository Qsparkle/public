"""
抖音达人批量截图工具页面
上传含达人主页链接的 Excel，自动打开页面截图，按 Excel名/Sheet名 分类打包下载。
"""
import io
import os
import time
import base64
import random
import zipfile
import threading
import pandas as pd
import streamlit as st
from PIL import Image as _PILImage
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

from utils.helpers import init_session, safe_filename, create_example_excel
from utils.cookie import cookie_section
from utils.driver import init_driver, add_cookies_to_driver
from utils.task_store import get_task, reset_task

st.set_page_config(page_title="抖音达人批量截图", page_icon="🎵", layout="wide")

DEFAULT_WAIT = 5
DY_URL = "https://www.douyin.com/"

# ===== 日期元素隐藏 JS =====
_JS_HIDE_DATE = """
(function() {
    function hideEl(el) {
        if (!el.dataset._hiddenByTool) {
            el.dataset._hiddenByTool = el.style.visibility || '';
        }
        el.style.visibility = 'hidden';
    }

    // 1. CSS 选择器批量隐藏（涵盖抖音常见日期/时间相关类名）
    var dateSelectors = [
        '.date', '.note-date', '.publish-date', '.post-date',
        '.time', '.post-time', '.create-time', '.upload-time',
        '[class*="date"]', '[class*="time"]',
        '[class*="publish"]', '[class*="create"]',
        '[class*="Duration"]', '[class*="duration"]'
    ];
    dateSelectors.forEach(function(sel) {
        try {
            document.querySelectorAll(sel).forEach(function(el) { hideEl(el); });
        } catch(e) {}
    });

    // 2. 包含「发布时间」「上传时间」「发布于」「创建时间」等中文标签的元素
    var dateKeywords = ['发布时间', '上传时间', '发布于', '创建时间', '更新时间', '拍摄时间'];

    // 3. 纯日期/时间文本正则（全文匹配）
    var dateReExact = /^(\d{4}[-年\/]\d{1,2}[-月\/]\d{1,2}[日]?(\s+\d{2}:\d{2}(:\d{2})?)?)$|^\d{1,2}-\d{1,2}$|^\d{1,2}月\d{1,2}日$|^\d{2}-\d{2}-\d{2}$|^\d{1,3}天前$|^\d{1,2}小时前$|^\d{1,2}分钟前$|^刚刚$|^昨天$|^前天$/;

    // 4. 包含日期的文本正则（含前缀/后缀，用 test 而非全文匹配）
    var dateReContains = /\d{4}[-年\/]\d{1,2}[-月\/]\d{1,2}[日]?(\s+\d{2}:\d{2})?/;

    var walker = document.createTreeWalker(
        document.body, NodeFilter.SHOW_ELEMENT, null, false
    );
    while (walker.nextNode()) {
        var node = walker.currentNode;
        if (node.childElementCount === 0) {
            var txt = node.textContent.trim();
            if (!txt) continue;
            // 全文精确匹配纯日期
            if (dateReExact.test(txt)) { hideEl(node); continue; }
            // 文本包含「发布时间」等中文关键词
            var hasKw = dateKeywords.some(function(kw) { return txt.indexOf(kw) !== -1; });
            if (hasKw && dateReContains.test(txt)) { hideEl(node); continue; }
        }
    }

    // 5. 针对抖音视频卡片：查找包含「发布时间：」文字的祖先容器并隐藏
    //    XPath 方式查找含「发布时间」的文本节点，再向上找到合适的父元素
    try {
        var xpathResult = document.evaluate(
            '//*[contains(text(),"发布时间") or contains(text(),"上传时间") or contains(text(),"发布于")]',
            document.body, null, XPathResult.UNORDERED_NODE_SNAPSHOT_TYPE, null
        );
        for (var i = 0; i < xpathResult.snapshotLength; i++) {
            hideEl(xpathResult.snapshotItem(i));
        }
    } catch(e) {}
})();
"""
_JS_RESTORE_DATE = """
(function() {
    document.querySelectorAll('[data-_hidden-by-tool]').forEach(function(el) {
        el.style.visibility = el.dataset._hiddenByTool || '';
        delete el.dataset._hiddenByTool;
    });
})();
"""


# 只保留非常具体的中文关键词，避免误判
_CAPTCHA_KEYWORDS = ["请完成下列验证后继续", "按住左边按钮拖动完成上方拼图"]
# 验证码弹窗的 DOM 选择器（双重确认）
_CAPTCHA_ELEM_SELECTORS = [
    '[class*="secsdk-captcha"]',
    '[class*="captcha-dialog"]',
    '[class*="verify-dialog"]',
]

def _has_captcha(driver) -> bool:
    """检测页面是否出现验证码弹窗（文本+DOM双重检测）"""
    try:
        src = driver.page_source
        # 必须命中具体中文关键词
        text_match = any(kw in src for kw in _CAPTCHA_KEYWORDS)
        if not text_match:
            return False
        # 再确认弹窗 DOM 元素存在且可见
        for sel in _CAPTCHA_ELEM_SELECTORS:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            if els and els[0].is_displayed():
                return True
        # DOM未找到，但文本命中也认为是验证码（兼容旧版）
        return True
    except Exception:
        return False


def _wait_captcha_clear(driver, timeout: int = 12) -> bool:
    """等待验证码弹窗彻底消失，消失返回True，超时返回False"""
    for _ in range(timeout * 2):
        if not _has_captcha(driver):
            return True
        time.sleep(0.5)
    return False


# ===== 页面失效检测 =====
# 只保留抖音错误页独有的完整句式，避免误判正常页面中的评论/标题
_PAGE_INVALID_TEXTS = [
    "您要观看的视频不存在",
    "很抱歉，您访问的页面不存在",
    "该页面已失效",
    "该内容已下架",
    "账号已注销",
]
# 抖音错误页特有的 DOM 元素选择器（双重确认）
_PAGE_INVALID_SELECTORS = [
    '[class*="error-page"]',
    '[class*="page-not-found"]',
    '[class*="video-not-found"]',
    '[class*="empty-page"]',
]


def _check_page_invalid(driver) -> str:
    """检测页面是否失效/不存在，返回原因文字，正常返回空字符串"""
    try:
        cur = driver.current_url
        if cur.startswith("chrome-error://") or cur.startswith("data:"):
            return "浏览器无法加载页面"
        src = driver.page_source
        # 1. 文本命中特定句式
        matched_kw = next((kw for kw in _PAGE_INVALID_TEXTS if kw in src), None)
        if not matched_kw:
            return ""
        # 2. DOM 元素二次确认（防止正文/评论中偶然包含关键词）
        for sel in _PAGE_INVALID_SELECTORS:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            if els and els[0].is_displayed():
                return matched_kw
        # 3. 若未找到错误 DOM，但文本命中的是非常具体的句式也直接判定
        DEFINITE_TEXTS = ["您要观看的视频不存在", "很抱歉，您访问的页面不存在"]
        if matched_kw in DEFINITE_TEXTS:
            return matched_kw
    except Exception:
        pass
    return ""


# ===== PIL 免费识别缺口位置 =====
_BG_IMG_SELECTORS = [
    '[class*="captcha-verify-image"]',
    '[class*="captcha"][class*="bg"]',
    '[class*="secsdk-captcha"] img',
    'img[class*="captcha-bg"]',
    'img[class*="verify-img"]',
]

def _fetch_img_bytes(driver, src: str) -> bytes:
    """通过浏览器 XHR 下载图片（绝过 CORS限制）"""
    if src.startswith("data:image"):
        return base64.b64decode(src.split(",", 1)[1])
    result = driver.execute_script("""
        try {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', arguments[0], false);
            xhr.responseType = 'arraybuffer';
            xhr.send();
            var bytes = new Uint8Array(xhr.response);
            var bin = '';
            for (var i = 0; i < bytes.byteLength; i++) bin += String.fromCharCode(bytes[i]);
            return btoa(bin);
        } catch(e) { return null; }
    """, src)
    if result:
        return base64.b64decode(result)
    return None


def _find_gap_x(image_bytes: bytes, skip_ratio: float = 0.12) -> int:
    """
    PIL 分析背景图找到缺口的 X 坐标。
    方法优先级：透明度 > 灰色色块形状识别 > 饱和度内滑窗口。
    """
    try:
        img = _PILImage.open(io.BytesIO(image_bytes)).convert("RGB")
        w, h = img.size
        pixels = img.load()

        # ---方法 1: Alpha 透明度---
        img_rgba = _PILImage.open(io.BytesIO(image_bytes)).convert("RGBA")
        alpha_px = img_rgba.load()
        skip = int(w * skip_ratio)
        alpha_cols = [sum(alpha_px[x, y][3] for y in range(h)) for x in range(w)]
        min_alpha = min(alpha_cols[skip:])
        if min_alpha < h * 200:
            return alpha_cols.index(min_alpha, skip)

        # 计算每列饱和度（max-min），灰色形状饱和度接近0
        sat_cols = []
        for x in range(w):
            s = sum(max(pixels[x, y]) - min(pixels[x, y]) for y in range(h))
            sat_cols.append(s / h)

        avg_sat = sum(sat_cols) / w

        # ---方法 2: 灰色色块识别---
        # 找出所有“灰色列块”（连续的饱和度远低于平均的列）
        BLOB_THRESHOLD = avg_sat * 0.45  # 饱和度 < 平均*0.45 认为灰色
        MIN_BLOB_W = max(8, int(w * 0.04))  # 色块最小宽度

        blobs = []  # [(center_x, width), ...]
        in_blob = False
        blob_start = 0
        for x, s in enumerate(sat_cols):
            if s < BLOB_THRESHOLD and not in_blob:
                in_blob = True
                blob_start = x
            elif s >= BLOB_THRESHOLD and in_blob:
                in_blob = False
                bw = x - blob_start
                if bw >= MIN_BLOB_W:
                    blobs.append((blob_start + bw // 2, bw))
        if in_blob:
            bw = w - blob_start
            if bw >= MIN_BLOB_W:
                blobs.append((blob_start + bw // 2, bw))

        if len(blobs) >= 2:
            # 有两个及以上色块，最左的是拼图块，最右的是缺口
            blobs_sorted = sorted(blobs, key=lambda b: b[0])
            # 取最右边的大色块中心（宽度最大且在右侧的）
            right_blobs = [b for b in blobs_sorted if b[0] > w * 0.3]
            if right_blobs:
                # 选宽度最大的右侧色块
                gap_blob = max(right_blobs, key=lambda b: b[1])
                return gap_blob[0]

        if len(blobs) == 1:
            # 只找到一个色块且在右侧，直接用它
            if blobs[0][0] > w * 0.3:
                return blobs[0][0]

        # ---方法 3: 饱和度滑动窗口内最低区域（兼容公测---
        right_start = int(w * 0.25)
        window = max(15, int(w * 0.07))
        min_score = float('inf')
        best_x = int(w * 0.55)
        for x in range(right_start, w - window):
            score = sum(sat_cols[x: x + window]) / window
            if score < min_score:
                min_score = score
                best_x = x + window // 2
        return best_x

    except Exception:
        return -1


def _extract_gap_target(driver, track_width: int) -> int:
    """
    提取验证码背景图并识别缺口位置，返回滑块应拖动的像素数。
    方式A: 从 img src 提取图片分析。
    方式B: 直接截图验证码容器分析（备用）。
    识别失败时返回 -1。
    """
    _CONTAINER_SELECTORS = [
        '[class*="secsdk-captcha"]',
        '[class*="captcha-dialog"]',
        '[class*="verify-dialog"]',
        '[class*="captcha-container"]',
    ]

    def _calc_target(img_bytes, skip_ratio=0.12):
        img_w = _PILImage.open(io.BytesIO(img_bytes)).size[0]
        gap_x = _find_gap_x(img_bytes, skip_ratio=skip_ratio)
        if gap_x > 0 and img_w > 0:
            scale = track_width / img_w
            return max(0, int(gap_x * scale) - 20)
        return -1

    try:
        # 方式 A：从 img src 提取
        for sel in _BG_IMG_SELECTORS:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                src = el.get_attribute("src") or ""
                if not src:
                    continue
                img_bytes = _fetch_img_bytes(driver, src)
                if img_bytes:
                    result = _calc_target(img_bytes)
                    if result > 0:
                        return result

        # 方式 B：截图验证码容器
        for sel in _CONTAINER_SELECTORS:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            if not els:
                continue
            try:
                img_bytes = base64.b64decode(els[0].screenshot_as_base64)
                result = _calc_target(img_bytes, skip_ratio=0.05)
                if result > 0:
                    return result
            except Exception:
                pass

        return -1
    except Exception:
        return -1


def _try_slide_captcha(driver) -> bool:
    """
    尝试自动滑动验证码。
    优先用 PIL 识别缺口位置精准拖动，失败时退回随机拖动。
    返回 True 表示通过、False 表示失败。
    """
    _SLIDER_SELECTORS = [
        '[class*="secsdk-captcha-drag-icon"]',
        '[class*="captcha-slider-btn"]',
        '[class*="slider-btn"]',
        '[class*="drag-btn"]',
        '[class*="drag-icon"]',
        '[class*="sc-captcha"] button',
        'button[class*="slider"]',
    ]
    _TRACK_SELECTORS = [
        '[class*="captcha-slider-inner"]',
        '[class*="slider-track"]',
        '[class*="drag-track"]',
    ]

    try:
        time.sleep(1.2)

        # 查找滑块
        slider = None
        for sel in _SLIDER_SELECTORS:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                if els and els[0].is_displayed():
                    slider = els[0]
                    break
            except Exception:
                pass
        if not slider:
            return False

        # 获取轨道宽度
        track_width = 280
        for sel in _TRACK_SELECTORS:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                if els:
                    w = els[0].rect.get('width', 0)
                    if w > 50:
                        track_width = w
                        break
            except Exception:
                pass

        # 尝试 PIL 精准识别缺口位置
        target = _extract_gap_target(driver, track_width)
        if target <= 0:
            # PIL 识别失败，退回随机位置
            target = int(track_width * random.uniform(0.38, 0.68))

        ac = ActionChains(driver)
        ac.click_and_hold(slider)
        ac.pause(random.uniform(0.3, 0.7))

        # 分段移动：先快后慢，加随机微振
        moved = 0
        steps = random.randint(18, 28)
        for i in range(steps):
            remaining = target - moved
            if remaining <= 0:
                break
            progress = i / steps
            ratio = (1 - progress) if progress < 0.5 else (1 - progress) * 0.4
            step = max(1, int(remaining * ratio))
            step = min(step, remaining)
            dx = step + random.randint(-1, 1)
            dy = random.randint(-2, 2)
            ac.move_by_offset(dx, dy)
            ac.pause(random.uniform(0.01, 0.05))
            moved += dx

        ac.pause(random.uniform(0.1, 0.4))
        ac.release()
        ac.perform()

        time.sleep(1.8)
        return not _has_captcha(driver)
    except Exception:
        return False


def screenshot_worker(task: dict, excel_data: bytes, excel_name: str,
                      cookies: list, wait_time: int, sid: str,
                      hide_date: bool = False):
    tmp_root = "/tmp" if os.path.isdir("/tmp") else "."
    base_tmp = os.path.join(tmp_root, f"temp_imgs_{sid}_dy")
    os.makedirs(base_tmp, exist_ok=True)

    excel_basename = safe_filename(os.path.splitext(excel_name)[0])
    main_dir = os.path.join(base_tmp, excel_basename)
    os.makedirs(main_dir, exist_ok=True)

    driver = None
    try:
        task["current_info"] = "正在启动浏览器…"
        driver = init_driver()
        task["current_info"] = "正在加载 Cookie…"
        add_cookies_to_driver(driver, cookies, url=DY_URL)
        task["current_info"] = "Cookie 已导入，重新加载页面确认登录态…"
        # refresh 有时不能触发抖音登录验证，重新导航更可靠
        driver.get(DY_URL)
        time.sleep(4)
        task["current_info"] = "初始化完成，开始截图…"

        xls = pd.ExcelFile(io.BytesIO(excel_data))
        total = sum(
            len(pd.read_excel(io.BytesIO(excel_data), sheet_name=s))
            for s in xls.sheet_names
        )
        task["total"] = total

        all_stats = []
        total_ok = total_fail = processed = 0

        for sheet_name in xls.sheet_names:
            if task["stop_event"].is_set():
                break

            safe_sheet = safe_filename(sheet_name)
            sheet_dir = os.path.join(main_dir, safe_sheet)
            os.makedirs(sheet_dir, exist_ok=True)

            df = pd.read_excel(io.BytesIO(excel_data), sheet_name=sheet_name).fillna("")
            ok = fail = 0
            failed_list = []

            for idx, row in df.iterrows():
                if task["stop_event"].is_set():
                    break

                # 暂停等待
                while task["pause_event"].is_set():
                    if task["stop_event"].is_set():
                        break
                    time.sleep(0.3)

                if task["stop_event"].is_set():
                    break

                processed += 1
                task["processed"] = processed
                task["current_info"] = f"工作表「{sheet_name}」第 {idx + 1}/{len(df)} 行"

                try:
                    if len(row) < 3:
                        continue
                    serial = str(row.iloc[0]).strip()
                    nick = str(row.iloc[1]).strip()
                    url = str(row.iloc[2]).strip()
                    if not url.startswith(("http://", "https://")):
                        continue

                    driver.get(url)
                    time.sleep(wait_time)

                    # 检测页面是否失效/视频不存在
                    invalid_reason = _check_page_invalid(driver)
                    if invalid_reason:
                        raise RuntimeError(f"页面失效：{invalid_reason}")

                    # 检测验证码：先尝试自动滑动，最多重试 2 次
                    if _has_captcha(driver):
                        solved = False
                        for _attempt in range(2):
                            task["current_info"] = f"检测到验证码，正在尝试自动滑动（第 {_attempt+1} 次）…"
                            solved = _try_slide_captcha(driver)
                            if solved:
                                break
                            time.sleep(random.uniform(1, 2))
                        if not solved:
                            raise RuntimeError("验证码自动滑动未通过，已跳过。建议降低批量或增加等待时间")
                        # 等待验证码弹窗彻底消失后再截图
                        task["current_info"] = "验证码已处理，等待页面恢复…"
                        if not _wait_captcha_clear(driver):
                            raise RuntimeError("验证码弹窗未消失，截图已跳过")
                        time.sleep(random.uniform(1.5, 2.5))  # 页面渲染缓冲

                    # 截图前最终确认无验证码
                    if _has_captcha(driver):
                        raise RuntimeError("截图前仍检测到验证码弹窗，已跳过")

                    # 随机延迟 1–3s，降低风控触发频率
                    time.sleep(random.uniform(1, 3))
                    if hide_date:
                        try:
                            driver.execute_script(_JS_HIDE_DATE)
                            time.sleep(0.3)
                        except Exception:
                            pass
                        fname = f"{serial}_{safe_filename(nick)}_无日期.png"
                        path = os.path.join(sheet_dir, fname)
                        driver.save_screenshot(path)
                        try:
                            driver.execute_script(_JS_RESTORE_DATE)
                        except Exception:
                            pass
                    else:
                        fname = f"{serial}_{safe_filename(nick)}.png"
                        path = os.path.join(sheet_dir, fname)
                        driver.save_screenshot(path)

                    ok += 1
                    total_ok += 1
                except Exception as e:
                    fail += 1
                    total_fail += 1
                    failed_list.append({
                        "sheet": sheet_name,
                        "行": idx + 1,
                        "序号": serial if 'serial' in dir() else "",
                        "昵称": nick if 'nick' in dir() else "",
                        "链接": url if 'url' in dir() else "",
                        "错误": str(e),
                    })

                task["total_ok"] = total_ok
                task["total_fail"] = total_fail
                time.sleep(0.5)

            all_stats.append({"sheet": sheet_name, "success": ok,
                               "fail": fail, "failed": failed_list})

        # 打包 ZIP（即使中途停止也打包已完成的截图）
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(main_dir):
                for file in files:
                    fp = os.path.join(root, file)
                    zf.write(fp, os.path.relpath(fp, base_tmp))

        task["zip_data"] = zip_buf.getvalue()
        task["stats"] = {
            "total_success": total_ok,
            "total_fail": total_fail,
            "details": all_stats,
            "excel_basename": excel_basename,
            "was_stopped": task["stop_event"].is_set(),
        }

    except Exception as e:
        task["error"] = str(e)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        # 清理临时文件
        try:
            for root, dirs, files in os.walk(base_tmp, topdown=False):
                for f in files:
                    os.remove(os.path.join(root, f))
                for d in dirs:
                    try:
                        os.rmdir(os.path.join(root, d))
                    except Exception:
                        pass
            os.rmdir(base_tmp)
        except Exception:
            pass

        task["running"] = False
        task["done"] = True


# ===== 页面 UI =====
def main():
    init_session(("excel_data", None), ("excel_name", None))
    sid = st.session_state.sid
    task_sid = f"{sid}_dy"
    task = get_task(task_sid)

    st.title("🎵 抖音达人批量截图工具")
    st.markdown("上传 Excel + 提供抖音 Cookie，一键截图并按 **Excel名/Sheet名** 分类打包下载。")

    # ---- 状态一：运行中 ----
    if task["running"]:
        st.info("🔄 任务进行中，请勿关闭页面...")

        progress_placeholder = st.empty()
        info_placeholder = st.empty()

        st.markdown("---")
        is_paused = task["pause_event"].is_set()
        is_stopping = task["stop_event"].is_set()

        btn_col1, _, btn_col3 = st.columns([1, 1, 1])
        with btn_col1:
            label = "▶️ 继续任务" if is_paused else "⏸️ 暂停任务"

            def toggle_pause(t=task):
                if t["pause_event"].is_set():
                    t["pause_event"].clear()
                else:
                    t["pause_event"].set()

            st.button(label, on_click=toggle_pause,
                      use_container_width=True, disabled=is_stopping)
        with btn_col3:
            def do_stop(t=task):
                t["stop_event"].set()
                t["pause_event"].clear()

            st.button("🛑 结束任务", type="primary", on_click=do_stop,
                      use_container_width=True, disabled=is_stopping)

        if is_paused and not is_stopping:
            st.warning("⏸️ 任务已暂停，点击「继续任务」以恢复。")
        elif is_stopping:
            st.warning("🛑 正在停止任务，请稍候...")

        while task["running"]:
            total = task["total"]
            processed = task["processed"]
            pct = min(processed / total, 1.0) if total > 0 else 0
            progress_placeholder.progress(pct, text=f"已处理 {processed}/{total} 条")
            with info_placeholder.container():
                c1, c2, c3 = st.columns(3)
                c1.info(task["current_info"] or "正在初始化浏览器…")
                c2.metric("✅ 成功", task["total_ok"])
                c3.metric("❌ 失败", task["total_fail"])
            time.sleep(0.5)

        st.rerun()
        return

    # ---- 状态二：完成 ----
    if task["done"]:
        if task.get("error"):
            st.error(f"❌ 任务出错：{task['error']}")
        elif task.get("stats"):
            stats = task["stats"]
            hint = "（已提前终止）" if stats.get("was_stopped") else ""
            st.success(f"🎉 处理完成 {hint}")

            c1, c2, c3 = st.columns(3)
            c1.metric("✅ 成功", stats["total_success"])
            c2.metric("❌ 失败", stats["total_fail"])
            c3.metric("📊 工作表数", len(stats["details"]))

            with st.expander("📋 查看详细结果"):
                for item in stats["details"]:
                    st.subheader(f"工作表：{item['sheet']}")
                    st.write(f"成功 {item['success']} / 失败 {item['fail']}")
                    if item["fail"] > 0:
                        st.dataframe(pd.DataFrame(item["failed"]), use_container_width=True)

            btn_col_a, btn_col_b = st.columns(2)
            if task.get("zip_data"):
                with btn_col_a:
                    st.download_button(
                        "📥 下载全部截图（ZIP）",
                        data=task["zip_data"],
                        file_name=f"{stats['excel_basename']}_抖音截图.zip",
                        mime="application/zip",
                        use_container_width=True,
                        type="primary",
                    )

            if stats["total_fail"] > 0:
                # 汇总所有失败行并生成 Excel 供下载
                all_failed = []
                for item in stats["details"]:
                    all_failed.extend(item["failed"])
                if all_failed:
                    _fail_buf = io.BytesIO()
                    pd.DataFrame(all_failed).to_excel(_fail_buf, index=False, engine="openpyxl")
                    with btn_col_b:
                        st.download_button(
                            "📋 下载失败数据（Excel）",
                            data=_fail_buf.getvalue(),
                            file_name=f"{stats['excel_basename']}_失败数据.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                st.warning("部分失败，可下载失败数据查看原因（页面失效、Cookie过期或链接无效）。")

        st.markdown("---")
        with st.popover("🔄 重新开始（上传新文件）", use_container_width=False):
            st.warning("⚠️ 当前任务结果将被清除，确认要重新开始吗？")
            if st.button("✅ 确认重置", type="primary", use_container_width=True, key="confirm_reset_dy"):
                reset_task(task_sid)
                st.session_state.excel_data = None
                st.session_state.excel_name = None
                st.rerun()
        return

    # ---- 状态三：表单 ----
    cookie_section("🔑 第一步：提供抖音 Cookie", cookie_key="dy_cookies", platform="douyin")

    st.subheader("📂 第二步：上传 Excel 文件")
    st.download_button(
        "📥 下载示例 Excel 表格",
        data=create_example_excel(),
        file_name="示例_抖音达人链接.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Excel 要求：至少三列，依次为 序号、达人昵称、链接，支持多个 Sheet。")

    excel_file = st.file_uploader("拖拽或选择你的 Excel 文件", type=["xlsx", "xls"])
    if excel_file:
        st.session_state.excel_data = excel_file.getvalue()
        st.session_state.excel_name = excel_file.name

    # 上传后立即计算行数并提示风控建议
    if st.session_state.get("excel_data"):
        try:
            _xls = pd.ExcelFile(io.BytesIO(st.session_state.excel_data))
            _total = sum(
                len(pd.read_excel(io.BytesIO(st.session_state.excel_data), sheet_name=s))
                for s in _xls.sheet_names
            )
            if _total <= 100:
                st.success(f"✅ 共 **{_total}** 条数据，可以开始执行。")
            else:
                st.warning(
                    f"⚠️ 共 **{_total}** 条数据，建议分批处理！\n"
                    "抖音风控较严，一次请求过多容易触发限流，\n"
                    "建议每批不超过 **100 条**，将数据拆分成多个文件分批执行。"
                )
        except Exception:
            pass

    wait_time = st.slider("⏱ 页面加载等待时间（秒）", 5, 10, DEFAULT_WAIT)
    hide_date = st.checkbox(
        "🗓️ 去除日期截图（截图前自动隐藏页面中的日期元素）",
        help="将尝试隐藏页面中的日期显示，截图文件名会加 _无日期 后缀以区分。"
    )

    can_start = bool(st.session_state.get("dy_cookies") and st.session_state.get("excel_data"))
    if st.button("🚀 开始批量截图", type="primary", disabled=not can_start):
        reset_task(task_sid)
        task = get_task(task_sid)
        task["running"] = True
        threading.Thread(
            target=screenshot_worker,
            args=(task, st.session_state.excel_data, st.session_state.excel_name,
                  st.session_state.dy_cookies, wait_time, task_sid),
            kwargs={"hide_date": hide_date},
            daemon=True,
        ).start()
        st.rerun()


if __name__ == "__main__":
    main()
