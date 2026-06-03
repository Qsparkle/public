import streamlit as st
from utils.helpers import init_session


def _home_page():
    """首页工具卡片"""
    st.set_page_config(page_title="内容运营工具箱", page_icon="🧰", layout="wide")
    init_session()

    st.title("🧰 内容运营工具箱")
    st.markdown("选择左侧导航栏中的工具开始使用。")
    st.markdown("---")

    cols = st.columns(2)

    with cols[0]:
        st.markdown("""
        ### 📸 xhs笔记批量截图
        上传 Excel 链接表，自动截图并按工作表分类打包下载。

        支持暂停、继续、提前结束。
        """)
        if st.button("进入工具", key="goto_screenshot", use_container_width=True):
            st.switch_page("pages/1_📸_xhs笔记批量截图.py")

    with cols[1]:
        st.markdown("""
        ### 🎵 抖音达人批量截图
        上传 Excel 链接表，自动截图并按工作表分类打包下载。

        支持暂停、继续、提前结束。
        """)
        if st.button("进入工具", key="goto_dy_screenshot", use_container_width=True):
            st.switch_page("pages/3_🎵_抖音达人批量截图.py")


# ===== 侧边栏导航配置 =====
# title 参数控制侧边栏显示的名称
# 不列入此处的页面不会出现在侧边栏
pg = st.navigation([
    st.Page(_home_page,                             title="🏠 首页"),
    st.Page("pages/1_📸_xhs笔记批量截图.py",        title="📸 xhs笔记批量截图"),
    st.Page("pages/3_🎵_抖音达人批量截图.py",        title="🎵 抖音达人批量截图"),
    # 暂不开放，不列入导航即可隐藏：
    # st.Page("pages/2_🔍_达人主页xhsID获取.py",   title="🔍 达人ID获取"),
])
pg.run()
