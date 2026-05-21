"""
小红书用户 ID 批量获取工具
上传含博主主页链接的 Excel，自动打开页面提取 XHS 用户 ID，输出为 Excel 下载。
"""
import io
import re
import time
import threading
import pandas as pd
import streamlit as st

from utils.helpers import init_session, safe_filename, create_example_excel
from utils.cookie import cookie_section
from utils.driver import init_driver, add_cookies_to_driver
from utils.task_store import get_task, reset_task

st.set_page_config(page_title="获取用户ID", page_icon="🔍", layout="wide")

DEFAULT_WAIT = 3

# 匹配页面元素中「小红书号：XXXXX」格式
_XHS_NO_RE = re.compile(r"小红书号[\uff1a:\s]*([\w\d]+)")


def extract_xhs_id(driver, url: str, wait_time: int) -> str:
    """
    打开 URL，从页面元素「小红书号：XXXXX」中提取小红书号。
    返回 ID 字符串，失败返回空字符串。
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver.get(url)
    time.sleep(wait_time)

    # 方法一：通过元素定位（包含「小红书号」文本的节点）
    try:
        # XPath 查找所有包含「小红书号」的元素
        elements = driver.find_elements(
            By.XPATH,
            "//*[contains(text(), '小红书号')]"
        )
        for el in elements:
            text = el.text.strip()
            m = _XHS_NO_RE.search(text)
            if m:
                return m.group(1)
    except Exception:
        pass

    # 方法二：从页面源码匹配（元素定位失败时兆底）
    try:
        source = driver.page_source
        m = _XHS_NO_RE.search(source)
        if m:
            return m.group(1)
    except Exception:
        pass

    return ""


# ===== 后台工作线程 =====
def id_fetch_worker(task: dict, excel_data: bytes, excel_name: str,
                    cookies: list, wait_time: int, sid: str):
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
        task["current_info"] = "初始化完成，开始获取 ID…"

        xls = pd.ExcelFile(io.BytesIO(excel_data))
        total = sum(
            len(pd.read_excel(io.BytesIO(excel_data), sheet_name=s))
            for s in xls.sheet_names
        )
        task["total"] = total

        all_sheet_results = {}   # sheet_name -> DataFrame（含结果列）
        all_stats = []
        total_ok = total_fail = processed = 0

        for sheet_name in xls.sheet_names:
            if task["stop_event"].is_set():
                break

            df = pd.read_excel(io.BytesIO(excel_data), sheet_name=sheet_name).fillna("")
            # 在原有列基础上追加"小红书ID"列
            xhs_ids = [""] * len(df)
            ok = fail = 0
            failed_list = []

            for idx, row in df.iterrows():
                if task["stop_event"].is_set():
                    break

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
                    url = str(row.iloc[2]).strip()
                    if not url.startswith(("http://", "https://")):
                        continue

                    xhs_id = extract_xhs_id(driver, url, wait_time)
                    xhs_ids[idx] = xhs_id

                    if xhs_id:
                        ok += 1
                        total_ok += 1
                    else:
                        fail += 1
                        total_fail += 1
                        failed_list.append({
                            "sheet": sheet_name, "行": idx + 1,
                            "链接": url, "原因": "页面中未找到「小红书号」元素"
                        })
                except Exception as e:
                    fail += 1
                    total_fail += 1
                    failed_list.append({
                        "sheet": sheet_name, "行": idx + 1,
                        "链接": str(row.iloc[2]) if len(row) >= 3 else "",
                        "原因": str(e)
                    })

                task["total_ok"] = total_ok
                task["total_fail"] = total_fail
                time.sleep(0.3)

            df["小红书ID"] = xhs_ids
            all_sheet_results[sheet_name] = df
            all_stats.append({
                "sheet": sheet_name, "success": ok,
                "fail": fail, "failed": failed_list
            })

        # 生成结果 Excel
        excel_basename = safe_filename(
            __import__("os").path.splitext(excel_name)[0]
        )
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
            for sheet_name, df_result in all_sheet_results.items():
                safe_sheet = safe_filename(sheet_name)
                df_result.to_excel(writer, sheet_name=safe_sheet, index=False)

        task["zip_data"] = out_buf.getvalue()   # 复用 zip_data 字段存 Excel 二进制
        task["stats"] = {
            "total_success": total_ok,
            "total_fail": total_fail,
            "details": all_stats,
            "excel_basename": excel_basename,
            "was_stopped": task["stop_event"].is_set(),
            "output_filename": f"{excel_basename}_小红书ID.xlsx",
        }

    except Exception as e:
        task["error"] = str(e)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        task["running"] = False
        task["done"] = True


def create_id_example_excel() -> bytes:
    """生成示例 Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({
            "序号": [1, 2],
            "昵称": ["博主A", "博主B"],
            "主页链接": [
                "https://www.xiaohongshu.com/user/profile/xxxxxxxxxxxxxxxxxxxxxxxx",
                "https://www.xiaohongshu.com/user/profile/yyyyyyyyyyyyyyyyyyyyyyyy",
            ],
        }).to_excel(writer, sheet_name="博主列表", index=False)
    return output.getvalue()


# ===== 页面 UI =====
def main():
    init_session(("excel_data", None), ("excel_name", None))
    sid = st.session_state.sid
    # 用独立的任务 key 避免与截图工具任务冲突
    task_sid = f"{sid}_id_fetch"
    task = get_task(task_sid)

    st.title("🔍 小红书用户 ID 批量获取")
    st.markdown(
        "上传含博主主页链接的 Excel，自动打开页面提取 **小红书用户 ID**，"
        "结果追加到原表并下载。"
    )

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
                c2.metric("✅ 成功获取", task["total_ok"])
                c3.metric("❌ 未获取到", task["total_fail"])
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
            c1.metric("✅ 成功获取", stats["total_success"])
            c2.metric("❌ 未获取到", stats["total_fail"])
            c3.metric("📊 工作表数", len(stats["details"]))

            with st.expander("📋 查看详细结果"):
                for item in stats["details"]:
                    st.subheader(f"工作表：{item['sheet']}")
                    st.write(f"成功 {item['success']} / 失败 {item['fail']}")
                    if item["fail"] > 0:
                        st.json(item["failed"])

            if task.get("zip_data"):
                st.download_button(
                    "📥 下载结果 Excel（含小红书ID列）",
                    data=task["zip_data"],
                    file_name=stats.get("output_filename", "小红书ID结果.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )

            if stats["total_fail"] > 0:
                st.warning(
                    "部分链接未获取到小红书号，可能原因：\n"
                    "- 页面中确实没有「小红书号」显示（未登录、被封）\n"
                    "- Cookie 过期导致页面跳登录\n"
                    "- 链接不是用户主页\n"
                    "- 等待时间太短，尝试增加加载时间"
                )

        st.markdown("---")
        with st.popover("🔄 重新开始（上传新文件）", use_container_width=False):
            st.warning("⚠️ 当前任务结果将被清除，确认要重新开始吗？")
            if st.button("✅ 确认重置", type="primary", use_container_width=True, key="confirm_reset_id"):
                reset_task(task_sid)
                st.session_state.excel_data = None
                st.session_state.excel_name = None
                st.rerun()
        return

    # ---- 状态三：表单 ----
    cookie_section("🔑 第一步：提供小红书 Cookie")

    st.subheader("📂 第二步：上传 Excel 文件")
    st.download_button(
        "📥 下载示例 Excel 表格",
        data=create_id_example_excel(),
        file_name="示例_博主主页链接.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Excel 要求：至少三列，依次为 序号、昵称、主页链接，支持多个 Sheet。")

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

    wait_time = st.slider("⏱ 页面加载等待时间（秒）", 2, 8, DEFAULT_WAIT)

    st.info(
        "💡 **支持的链接格式**\n"
        "- 用户主页：`https://www.xiaohongshu.com/user/profile/xxxxxxxx`\n"
        "- 笔记链接会自动跳转到用户主页后提取 ID"
    )

    can_start = bool(st.session_state.get("cookies") and st.session_state.get("excel_data"))
    if st.button("🚀 开始批量获取 ID", type="primary", disabled=not can_start):
        reset_task(task_sid)
        task = get_task(task_sid)
        task["running"] = True
        threading.Thread(
            target=id_fetch_worker,
            args=(task, st.session_state.excel_data, st.session_state.excel_name,
                  st.session_state.cookies, wait_time, task_sid),
            daemon=True,
        ).start()
        st.rerun()


if __name__ == "__main__":
    main()
