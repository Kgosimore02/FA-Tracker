"""Shared UI — styles, logo, sidebar nav, page header."""
import base64, pathlib, streamlit as st
from db import run_query

TEAL = "#29B6D2"
NAVY = "#1C2833"

def _logo_b64() -> str | None:
    p = pathlib.Path(__file__).parent / "assets" / "blil_logo.svg"
    if p.exists():
        return base64.b64encode(p.read_bytes()).decode()
    return None

def inject_css():
    st.markdown(f"""<style>
    .stApp{{background:{NAVY};color:#fff}}
    section[data-testid="stSidebar"]{{background:#0f1c24!important}}
    section[data-testid="stSidebar"] *{{color:#fff!important}}
    .nav-lbl{{font-size:.6rem;font-weight:700;letter-spacing:.12em;
              color:rgba(255,255,255,.35)!important;text-transform:uppercase;
              padding:.8rem 0 .2rem;margin:0}}
    section[data-testid="stSidebar"] .stButton>button{{
        width:100%;background:rgba(41,182,210,.08);color:#fff!important;
        border:1px solid rgba(41,182,210,.2);border-radius:6px;
        padding:.4rem .7rem;margin-bottom:3px;text-align:left;font-size:.86rem}}
    section[data-testid="stSidebar"] .stButton>button:hover{{
        background:rgba(41,182,210,.22)!important;border-color:{TEAL}!important}}
    .nav-active>div>button{{
        background:{TEAL}!important;color:{NAVY}!important;
        font-weight:700!important;border-color:{TEAL}!important}}
    .stSelectbox>div>div{{background:#1e303d;color:#fff}}
    div[data-testid="metric-container"]{{
        background:rgba(41,182,210,.08);
        border:1px solid rgba(41,182,210,.18);border-radius:8px;padding:.6rem 1rem}}
    .stButton>button{{background:{TEAL};color:{NAVY};font-weight:700;border:none;border-radius:6px}}
    .stButton>button:hover{{background:#1da0bc!important;color:#fff!important}}
    h1{{color:{TEAL}!important;border-bottom:1px solid rgba(41,182,210,.25);padding-bottom:.35rem}}
    h2,h3{{color:{TEAL}!important}}
    .stDataFrame{{border-radius:8px;overflow:hidden}}
    .footer{{font-size:.6rem;color:rgba(255,255,255,.2);text-align:center;padding:1.5rem 0 .5rem}}
    </style>""", unsafe_allow_html=True)


def sidebar_nav(current: str) -> dict:
    """Render sidebar. Returns selected branch dict."""
    with st.sidebar:
        # ── Logo ──────────────────────────────────────────────────────────────
        b64 = _logo_b64()
        if b64:
            st.markdown(
                f'<div style="text-align:center;padding:.8rem 0 .4rem">'
                f'<img src="data:image/svg+xml;base64,{b64}" '
                f'style="max-width:180px;width:100%"></div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div style="text-align:center;padding:.8rem 0 .4rem">'
                f'<span style="font-size:1.3rem;font-weight:800;color:{TEAL}">BLIL</span>'
                f'<span style="font-size:.7rem;color:rgba(255,255,255,.4);'
                f'display:block;letter-spacing:.08em">FA PRODUCTIVITY TRACKER</span></div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")

        # ── Branch selector ───────────────────────────────────────────────────
        branches = run_query(
            "SELECT branch_id, branch_name, channel_code "
            "FROM branches WHERE is_active=TRUE ORDER BY channel_code"
        )
        labels = [f"{b['channel_code']} — {b['branch_name']}" for b in branches]
        bmap   = {l: b for l, b in zip(labels, branches)}

        if "branch_lbl" not in st.session_state or \
                st.session_state.branch_lbl not in bmap:
            st.session_state.branch_lbl = labels[0] if labels else ""

        sel = st.selectbox("Branch", labels,
                           index=labels.index(st.session_state.branch_lbl)
                                 if st.session_state.branch_lbl in labels else 0,
                           key="branch_sel")
        st.session_state.branch_lbl = sel
        branch = bmap.get(sel, branches[0] if branches else {})

        st.markdown("---")

        # ── DATA ENTRY ────────────────────────────────────────────────────────
        st.markdown('<p class="nav-lbl">Data Entry</p>', unsafe_allow_html=True)
        _btn("daily_entry",  "📋  Daily NOC Entry",    current)
        _btn("agent_mgmt",   "👤  Agent Management",    current)

        # ── REPORTS & EXPORTS ─────────────────────────────────────────────────
        st.markdown('<p class="nav-lbl">Reports & Exports</p>', unsafe_allow_html=True)
        _btn("export_daily",        "📥  Daily Report",           current)
        _btn("export_weekly",       "📥  Weekly Report",          current)
        _btn("export_monthly",      "📥  Monthly Report",         current)
        _btn("export_quarterly",    "📥  Quarterly Report",       current)
        _btn("export_appointments", "📥  Weekly Appointments",    current)

        st.markdown("---")
        st.markdown('<p class="footer">Botswana Life Insurance Ltd © 2026</p>',
                    unsafe_allow_html=True)

    return branch


def _btn(key, label, current):
    active = key == current
    if active:
        st.markdown('<div class="nav-active">', unsafe_allow_html=True)
    if st.button(label, key=f"nav_{key}"):
        st.session_state.page = key
        st.rerun()
    if active:
        st.markdown('</div>', unsafe_allow_html=True)


def page_header(title: str, sub: str = ""):
    st.markdown(f"# {title}")
    if sub:
        st.markdown(f'<p style="color:rgba(255,255,255,.45);margin-top:-.5rem">'
                    f'{sub}</p>', unsafe_allow_html=True)
