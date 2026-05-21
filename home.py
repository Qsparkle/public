import streamlit as st
from utils.helpers import init_session

st.set_page_config(page_title="工具箱", page_icon="🧰", layout="wide")

init_session()

st.title("🧰 内容运营工具箱")
st.markdown("选择左侧导航栏中的工具开始使用。")

st.markdown("---")

# 工具卡片展示（新增工具时在这里加卡片）
cols = st.columns(3)

with cols[0]:
    st.markdown("""
    ### 📸 xhs笔记批量截图
    上传 Excel 链接表，自动截图并按工作表分类打包下载。

    支持暂停、继续、提前结束。
    """)
    if st.button("进入工具", key="goto_screenshot", use_container_width=True):
        st.switch_page("pages/1_📸_xhs笔记批量截图.py")

# 展占位用，新工具加在此处
with cols[1]:
    st.markdown("""
    ### 🔍 用户 ID 批量获取
    上传博主主页链接 Excel，自动提取小红书用户 ID 并回写入表格。

    支持多表、暂停、提前结束。
    """)
    if st.button("进入工具", key="goto_id_fetch", use_container_width=True):
        st.switch_page("pages/2_🔍_达人主页xhsID获取.py")

# st.markdown("---")
# st.caption("💡 工具箱基于 Streamlit 构建，所有数据在会话内隔离，互不影响。")
