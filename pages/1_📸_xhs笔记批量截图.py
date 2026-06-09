"""
小红书批量截图工具页面
"""
import io
import os
import time
import zipfile
import threading
import pandas as pd
import streamlit as st

from utils.helpers import init_session, safe_filename, create_example_excel
from utils.cookie import cookie_section
from utils.driver import init_driver, add_cookies_to_driver
from utils.task_store import get_task, reset_task

st.set_page_config(page_title="xhs笔记批量截图", page_icon="📸", layout="wide")

DEFAULT_WAIT = 5

# ===== 小红书登录弹窗检测 =====
# 登录弹窗特有的组合关键词：必须同时命中“登录后推荐更懂你的笔记”
# 并且命中其中至少一个内容可验证关键词，才判定 Cookie 失效
_XHS_LOGIN_ANCHOR = "登录后推荐更懂你的笔记"  # 只在登录弹窗主标题出现
_XHS_LOGIN_CONFIRM = ["新用户可直接登录", "手机号登录"]  # 必须同时命中其中一个


def _xhs_check_login_popup(driver) -> bool:
    """检测小红书是否弹出登录弹窗（Cookie 失效标志）
    需要标题 + 确认关键词同时命中，防止单一关键词导致误判
    """
    try:
        text = ""
        try:
            text = driver.execute_script("return document.body ? document.body.innerText : '';")
        except Exception:
            pass
        if not text:
            text = driver.page_source
        # 必须先命中标题长句
        if _XHS_LOGIN_ANCHOR not in text:
            return False
        # 再确认至少一个内容可验证关键词
        return any(kw in text for kw in _XHS_LOGIN_CONFIRM)
    except Exception:
        return False


# ===== 小红书页面失效检测 =====
# 只保留小红书错误页独有的完整句式，避免误判正文/评论内容
_XHS_INVALID_TEXTS = [
    "你访问的笔记不见了",
    "该笔记已被作者删除",
    "该内容已被删除",
    "很抱歉，此内容已无法显示",
    "该内容已下架",
    "当前笔记暂时无法浏览",
]
_XHS_INVALID_SELECTORS = [
    '[class*="error-page"]',
    '[class*="not-found"]',
    '[class*="note-not-found"]',
    '[class*="empty-page"]',
]


def _xhs_check_page_invalid(driver) -> str:
    """检测小红书页面是否失效，返回原因文字，正常返回空字符串"""
    from selenium.webdriver.common.by import By
    try:
        cur = driver.current_url
        if cur.startswith("chrome-error://") or cur.startswith("data:"):
            return "浏览器无法加载页面"
        src = driver.page_source
        matched_kw = next((kw for kw in _XHS_INVALID_TEXTS if kw in src), None)
        if not matched_kw:
            return ""
        # DOM 元素二次确认
        for sel in _XHS_INVALID_SELECTORS:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            if els and els[0].is_displayed():
                return matched_kw
        # 非常具体的大长句式直接判定
        DEFINITE_TEXTS = ["你访问的笔记不见了", "很抱歉，此内容已无法显示", "当前笔记暂时无法浏览"]
        if matched_kw in DEFINITE_TEXTS:
            return matched_kw
    except Exception:
        pass
    return ""


# ===== 小红书日期元素隐藏 JS =====
# 同时包含：按日期 class 隐藏 + 按文本匹配日期格式隐藏
_JS_HIDE_DATE = """
(function() {
    // 1. 小红书已知日期相关类名
    var dateSelectors = [
        '.date', '.note-date', '.publish-date', '.post-date',
        '.time', '.post-time', '.create-time', '.upload-time',
        '[class*="date"]', '[class*="time"]'
    ];
    dateSelectors.forEach(function(sel) {
        document.querySelectorAll(sel).forEach(function(el) {
            el.dataset._hiddenByTool = el.style.visibility;
            el.style.visibility = 'hidden';
        });
    });
    // 2. 匹配日期文本的纯文本节点
    var walker = document.createTreeWalker(
        document.body, NodeFilter.SHOW_ELEMENT, null, false
    );
    var dateRe = /^\d{4}[-年]\d{1,2}[-月]\d{1,2}[日]?$|^\d{2}-\d{2}-\d{2}$|^\d{4}\/\d{2}\/\d{2}$/;
    while (walker.nextNode()) {
        var node = walker.currentNode;
        if (node.childElementCount === 0) {
            var txt = node.textContent.trim();
            if (dateRe.test(txt)) {
                node.dataset._hiddenByTool = node.style.visibility;
                node.style.visibility = 'hidden';
            }
        }
    }
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


# ===== 后台截图线程 =====
def screenshot_worker(task: dict, excel_data: bytes, excel_name: str,
                      cookies: list, wait_time: int, sid: str,
                      hide_date: bool = False):
    tmp_root = "/tmp" if os.path.isdir("/tmp") else "."
    base_tmp = os.path.join(tmp_root, f"temp_imgs_{sid}")
    os.makedirs(base_tmp, exist_ok=True)

    excel_basename = safe_filename(os.path.splitext(excel_name)[0])
    main_dir = os.path.join(base_tmp, excel_basename)
    os.makedirs(main_dir, exist_ok=True)

    driver = None
    try:
        task["current_info"] = "正在启动浏览器…"
        driver = init_driver()
        task["current_info"] = "正在加载 Cookie…"
        add_cookies_to_driver(driver, cookies)
        task["current_info"] = "正在刷新页面…"
        try:
            driver.refresh()
        except Exception:
            driver.get("https://www.xiaohongshu.com/")
        time.sleep(2)
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

                    # 检测 Cookie 是否失效（登录弹窗）——一旦发现立即终止整个任务
                    if _xhs_check_login_popup(driver):
                        raise RuntimeError(
                            "❗Cookie 已过期，小红书登录弹窗出现。"
                            "请重新获取小红书 Cookie 后再运行。",
                            "COOKIE_EXPIRED"  # 特殊标识，上层捕获后终止任务
                        )

                    # 检测页面是否失效，首次检测到则额外等 3 秒复查
                    invalid_reason = _xhs_check_page_invalid(driver)
                    if invalid_reason:
                        task["current_info"] = f"疑似页面失效（{invalid_reason}），复查中…"
                        time.sleep(3)
                        invalid_reason = _xhs_check_page_invalid(driver)
                    if invalid_reason:
                        raise RuntimeError(f"页面失效：{invalid_reason}")

                    if hide_date:
                        # 隐藏日期元素 → 截图 → 恢复
                        try:
                            driver.execute_script(_JS_HIDE_DATE)
                            time.sleep(0.3)  # 等待渲染生效
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
                    err_msg = str(e)
                    # Cookie 过期——终止整个任务
                    if "COOKIE_EXPIRED" in err_msg or "Cookie 已过期" in err_msg:
                        all_stats.append({"sheet": sheet_name, "success": ok,
                                          "fail": fail, "failed": failed_list})
                        raise RuntimeError(
                            "❗ Cookie 已过期，小红书登录弹窗出现。"
                            "请重新获取小红书 Cookie 后再运行。"
                        )
                    fail += 1
                    total_fail += 1
                    failed_list.append({
                        "sheet": sheet_name,
                        "行": idx + 1,
                        "序号": serial if 'serial' in dir() else "",
                        "昵称": nick if 'nick' in dir() else "",
                        "链接": url if 'url' in dir() else "",
                        "错误": err_msg,
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
    task = get_task(sid)

    st.title("📸 小红书网页批量截图工具")
    st.markdown("上传 Excel + 提供小红书 Cookie，一键截图并按 **Excel名/Sheet名** 分类打包下载。")

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
                c3.metric("❌ 失败或链接失效", task["total_fail"])
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
            c2.metric("❌ 失败或链接失效", stats["total_fail"])
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
                        file_name=f"{stats['excel_basename']}_截图.zip",
                        mime="application/zip",
                        use_container_width=True,
                        type="primary",
                    )

            if stats["total_fail"] > 0:
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
            if st.button("✅ 确认重置", type="primary", use_container_width=True, key="confirm_reset_screenshot"):
                reset_task(sid)
                st.session_state.excel_data = None
                st.session_state.excel_name = None
                st.rerun()
        return

    # ---- 状态三：表单 ----
    cookie_section("🔑 第一步：提供小红书 Cookie")

    st.subheader("📂 第二步：上传 Excel 文件")
    st.download_button(
        "📥 下载示例 Excel 表格",
        data=create_example_excel(),
        file_name="示例_小红书链接.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Excel 要求：至少三列，依次为 序号、昵称、链接，支持多个 Sheet。")

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
                    "小红书风控较严，一次请求过多容易触发限流或封号，\n"
                    "建议每批不超过 **100 条**，将数据拆分成多个文件分批执行。"
                )
        except Exception:
            pass

    wait_time = st.slider("⏱ 页面加载等待时间（秒）", 5, 10, DEFAULT_WAIT)
    hide_date = st.checkbox(
        "🗓️ 去除日期截图（截图前自动隐藏页面中的日期元素）",
        help="将尝试隐藏页面中的日期显示，截图文件名会加 _无日期 后缀以区分。"
    )

    can_start = bool(st.session_state.get("cookies") and st.session_state.get("excel_data"))
    if st.button("🚀 开始批量截图", type="primary", disabled=not can_start):
        reset_task(sid)
        task = get_task(sid)
        task["running"] = True
        threading.Thread(
            target=screenshot_worker,
            args=(task, st.session_state.excel_data, st.session_state.excel_name,
                  st.session_state.cookies, wait_time, sid),
            kwargs={"hide_date": hide_date},
            daemon=True,
        ).start()
        st.rerun()


if __name__ == "__main__":
    main()
