"""
通用工具函数，供各页面复用。
"""
import re
import io
import uuid
import pandas as pd
import streamlit as st


def safe_filename(name: str) -> str:
    """将字符串转为安全的文件名"""
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    name = name.strip(". ")
    return name or "未命名"


def init_session(*extra_keys: tuple):
    """
    初始化 session state。
    - 固定初始化 sid（用户唯一标识）和 cookies
    - extra_keys: 额外需要初始化的 (key, default_value) 元组
    """
    if "sid" not in st.session_state:
        st.session_state.sid = str(uuid.uuid4())
    if "cookies" not in st.session_state:
        st.session_state.cookies = None
    for key, default in extra_keys:
        if key not in st.session_state:
            st.session_state[key] = default


def create_example_excel() -> bytes:
    """生成示例 Excel 文件，供用户下载参考格式"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({
            "序号": [1, 2],
            "昵称": ["小红书官方", "示例用户"],
            "链接": ["https://www.xiaohongshu.com/explore", "https://www.xiaohongshu.com"],
        }).to_excel(writer, sheet_name="小红书链接", index=False)
        pd.DataFrame(columns=["序号", "昵称", "链接"]).to_excel(
            writer, sheet_name="其他页面", index=False
        )
    return output.getvalue()
