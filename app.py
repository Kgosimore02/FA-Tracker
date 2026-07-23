"""FA Tracker — entry point. No auth for beta."""
import traceback, importlib, streamlit as st

st.set_page_config(page_title="BLIL FA Tracker", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

from ui import inject_css, sidebar_nav

inject_css()

if "page" not in st.session_state:
    st.session_state.page = "daily_entry"

PAGE_MAP = {
    "daily_entry":         "views.daily_entry",
    "budget_setup":        "views.budget_setup",
    "agent_mgmt":          "views.agent_mgmt",
    "export_daily":        "views.export_daily",
    "export_weekly":       "views.export_weekly",
    "export_monthly":      "views.export_monthly",
    "export_quarterly":    "views.export_quarterly",
    "export_appointments": "views.export_appointments",
}

try:
    branch = sidebar_nav(st.session_state.page)
except Exception:
    st.error("⚠️  Cannot reach database — check db.py credentials.")
    st.code(traceback.format_exc())
    st.stop()

mod_path = PAGE_MAP.get(st.session_state.page)
if mod_path:
    try:
        mod = importlib.import_module(mod_path)
        importlib.reload(mod)
        mod.render(branch)
    except Exception:
        st.error(f"Error in `{st.session_state.page}`:")
        st.code(traceback.format_exc())
