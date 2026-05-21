"""
任务注册表：基于 @st.cache_resource，跨所有 rerun 和用户持久化。
Streamlit 每次交互都重新执行脚本，模块级变量会被重置，
cache_resource 是共享有状态资源的正确方式。
"""
import threading
import streamlit as st


@st.cache_resource
def get_task_store() -> dict:
    """返回全局任务仓库，整个应用生命周期只创建一次"""
    return {"registry": {}, "lock": threading.Lock()}


def _new_task() -> dict:
    return {
        "running": False,
        "done": False,
        "processed": 0,
        "total": 0,
        "current_info": "",
        "total_ok": 0,
        "total_fail": 0,
        "zip_data": None,
        "stats": None,
        "error": None,
        "stop_event": threading.Event(),
        "pause_event": threading.Event(),   # is_set() 表示已暂停
    }


def get_task(sid: str) -> dict:
    store = get_task_store()
    with store["lock"]:
        if sid not in store["registry"]:
            store["registry"][sid] = _new_task()
        return store["registry"][sid]


def reset_task(sid: str):
    """重置任务，若有正在运行的线程先发送停止信号"""
    store = get_task_store()
    with store["lock"]:
        old = store["registry"].get(sid)
        if old:
            old["stop_event"].set()
            old["pause_event"].clear()
        store["registry"][sid] = _new_task()
