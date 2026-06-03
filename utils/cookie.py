"""
Cookie 解析与 UI 渲染，供各页面复用。
支持通过 cookie_key 参数隔离不同平台的 Cookie，默认为 "cookies"。
"""
import json
import pickle
import streamlit as st

# 各平台登录引导文案
_PLATFORM_GUIDE = {
    "xhs": """
**如何获取小红书 Cookie：**
1. 用 Chrome 打开 [小红书](https://www.xiaohongshu.com) 并登录
2. 按 `F12` → `Application` → 左侧 `Cookies` → 点击域名
3. 右侧表格 **全选 (Ctrl+A)** → **复制 (Ctrl+C)**
4. 粘贴到下方文本框
    """,
    "douyin": """
**如何获取抖音 Cookie：**
1. 用 Chrome 打开 [抖音](https://www.douyin.com) 并登录
2. 按 `F12` → `Application` → 左侧 `Cookies` → 点击 **`https://www.douyin.com`**
3. 在右侧表格中任意右击一条 Cookie → 选择 **Show requests with this cookie**
4. 自动跳转到 `Network` 页签，随便点击一条请求
5. 右侧 `Headers` → 向下找到 `Request Headers` → 找到 **`cookie:`** 这一行
6. 右击 `cookie:` → **Copy value** ，粘贴到下方文本框

> ⚠️ 该方式包含所有登录字段（包括 HttpOnly Cookie），拷贝表格可能丢失部分字段
    """,
}


def parse_cookie_input(cookie_text: str):
    """将多种格式的 Cookie 文本解析为列表，返回 (cookies, error)"""
    # 尝试 JSON
    try:
        data = json.loads(cookie_text)
        if isinstance(data, dict) and "cookies" in data:
            return data["cookies"], None
        if isinstance(data, list):
            return data, None
    except Exception:
        pass

    # 制表符分隔（Chrome DevTools 复制格式）
    lines = cookie_text.strip().splitlines()
    tab_cookies = []
    for line in lines:
        if not line.strip():
            continue
        parts = line.split("\t")
        if len(parts) >= 4:
            tab_cookies.append({
                "name": parts[0].strip(),
                "value": parts[1].strip(),
                "domain": parts[2].strip(),
                "path": parts[3].strip(),
            })
    if tab_cookies:
        return tab_cookies, None

    # name=value; 格式
    nv = []
    for item in cookie_text.split(";"):
        item = item.strip()
        if "=" in item:
            k, v = item.split("=", 1)
            nv.append({"name": k.strip(), "value": v.strip()})
    if nv:
        return nv, None

    return None, "无法识别格式。请从 Chrome DevTools → Cookies 表格全选复制后粘贴。"


def render_cookie_input(cookie_key: str = "cookies", platform: str = "xhs"):
    """Cookie 输入 UI，解析成功后存入 session_state[cookie_key]"""
    guide = _PLATFORM_GUIDE.get(platform, _PLATFORM_GUIDE["xhs"])
    tab1, tab2 = st.tabs(["📋 粘贴 Cookie", "📁 上传 .pkl 文件"])

    with tab1:
        st.markdown(guide)
        cookie_text = st.text_area("在此粘贴 Cookie 内容", height=150,
                                   placeholder="从 Chrome 直接复制的表格内容",
                                   key=f"{cookie_key}_textarea")
        if cookie_text.strip():
            cookies, err = parse_cookie_input(cookie_text)
            if cookies is None:
                st.error(f"❌ {err}")
            else:
                st.success(f"✅ 成功解析 {len(cookies)} 个 Cookie，将自动保留")
                st.session_state[cookie_key] = cookies

    with tab2:
        cookie_file = st.file_uploader("选择 cookies.pkl", type=["pkl"],
                                       key=f"{cookie_key}_uploader")
        if cookie_file:
            try:
                st.session_state[cookie_key] = pickle.loads(cookie_file.getvalue())
                st.success(f"✅ 从 .pkl 加载了 {len(st.session_state[cookie_key])} 个 Cookie，将自动保留")
            except Exception:
                st.error("❌ .pkl 文件损坏或格式不正确")


def cookie_section(title: str = "🔑 第一步：提供 Cookie",
                   cookie_key: str = "cookies",
                   platform: str = "xhs"):
    """
    完整的 Cookie 区域：
    - 已有 Cookie → 显示状态 + 清除按钮 + 折叠更换入口
    - 无 Cookie   → 展开输入区域
    cookie_key 用于隔离不同平台的 Cookie（默认 "cookies" 兼容小红书工具）。
    """
    st.subheader(title)

    if st.session_state.get(cookie_key):
        col_info, col_clear = st.columns([4, 1])
        with col_info:
            st.success(
                f"✅ Cookie 已加载（{len(st.session_state[cookie_key])} 个），"
                "任务结束后自动保留，无需重新粘贴。"
            )
        with col_clear:
            if st.button("🗑️ 清除", use_container_width=True, key=f"{cookie_key}_clear"):
                st.session_state[cookie_key] = None
                st.rerun()
        with st.expander("🔄 更换 Cookie（可选）"):
            render_cookie_input(cookie_key=cookie_key, platform=platform)
    else:
        render_cookie_input(cookie_key=cookie_key, platform=platform)
