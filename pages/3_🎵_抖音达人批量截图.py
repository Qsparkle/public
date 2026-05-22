"""
抖音达人批量截图工具页面
上传含达人主页链接的 Excel，自动打开页面截图，按 Excel名/Sheet名 分类打包下载。
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

st.set_page_config(page_title="抖音达人批量截图", page_icon="🎵", layout="wide")

DEFAULT_WAIT = 5
DY_URL = "https://www.douyin.com/"


# ===== 后台截图线程 =====
def screenshot_worker(task: dict, excel_data: bytes, excel_name: str,
                      cookies: list, wait_time: int, sid: str):
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
        task["current_info"] = "正在刷新页面…"
        try:
            driver.refresh()
        except Exception:
            driver.get(DY_URL)
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
                    fname = f"{serial}_{safe_filename(nick)}.png"
                    path = os.path.join(sheet_dir, fname)
                    driver.get(url)
                    time.sleep(wait_time)
                    driver.save_screenshot(path)
                    ok += 1
                    total_ok += 1
                except Exception as e:
                    fail += 1
                    total_fail += 1
                    failed_list.append({"sheet": sheet_name, "行": idx + 1, "错误": str(e)})

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
                        st.json(item["failed"])

            if task.get("zip_data"):
                st.download_button(
                    "📥 下载全部截图（ZIP）",
                    data=task["zip_data"],
                    file_name=f"{stats['excel_basename']}_抖音截图.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary",
                )

            if stats["total_fail"] > 0:
                st.warning("部分失败，可能是 Cookie 过期或链接无效，建议重新获取抖音 Cookie。")

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
    cookie_section("🔑 第一步：提供抖音 Cookie")

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

    can_start = bool(st.session_state.get("cookies") and st.session_state.get("excel_data"))
    if st.button("🚀 开始批量截图", type="primary", disabled=not can_start):
        reset_task(task_sid)
        task = get_task(task_sid)
        task["running"] = True
        threading.Thread(
            target=screenshot_worker,
            args=(task, st.session_state.excel_data, st.session_state.excel_name,
                  st.session_state.cookies, wait_time, task_sid),
            daemon=True,
        ).start()
        st.rerun()


if __name__ == "__main__":
    main()
