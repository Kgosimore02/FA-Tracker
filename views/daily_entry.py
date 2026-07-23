"""Daily NOC Entry — NOC count + monthly premium per FA per day."""
import datetime
import streamlit as st
from db import run_query, run_write_many, refresh_views
from ui import page_header

def render(branch: dict):
    page_header("Daily NOC Entry",
                f"{branch.get('channel_code','')} — {branch.get('branch_name','')}")

    branch_id = branch["branch_id"]

    col_date, _ = st.columns([2, 5])
    with col_date:
        entry_date = st.date_input("Date", datetime.date.today(),
                                   max_value=datetime.date.today())

    agents = run_query(
        "SELECT agent_id, full_name FROM field_agents "
        "WHERE branch_id=:bid AND is_active=TRUE ORDER BY full_name",
        {"bid": branch_id},
    )
    if not agents:
        st.warning("No active FAs. Add agents in Agent Management.")
        return

    # ANY with a list needs ARRAY cast — plain :ids binding won't work with SQLAlchemy text()
    agent_ids = [a["agent_id"] for a in agents]
    existing = run_query(
        "SELECT agent_id, noc_count, monthly_premium FROM daily_production "
        "WHERE agent_id = ANY(ARRAY[:ids]::int[]) AND production_date = :dt",
        {"ids": agent_ids, "dt": entry_date},
    ) if agent_ids else []
    exist_map = {r["agent_id"]: r for r in existing}

    st.markdown(
        f'<p style="color:rgba(255,255,255,.5);font-size:.85rem">'
        f'{entry_date.strftime("%A, %d %B %Y")} &nbsp;·&nbsp; {len(agents)} active FAs</p>',
        unsafe_allow_html=True,
    )
    st.markdown("---")

    h1, h2, h3 = st.columns([4, 2, 2.5])
    h1.markdown("**FA Name**"); h2.markdown("**NOC Cases**")
    h3.markdown("**Monthly Premium (BWP)**")

    rows = []
    for ag in agents:
        aid  = ag["agent_id"]
        prev = exist_map.get(aid, {})
        c1, c2, c3 = st.columns([4, 2, 2.5])
        with c1:
            st.markdown(f'<div style="padding:.4rem 0">{ag["full_name"]}</div>',
                        unsafe_allow_html=True)
        with c2:
            noc = st.number_input("noc", 0, 999, int(prev.get("noc_count", 0)),
                                  key=f"n{aid}", label_visibility="collapsed")
        with c3:
            prem = st.number_input("prem", 0.0, step=0.01,
                                   value=float(prev.get("monthly_premium", 0.0)),
                                   key=f"p{aid}", label_visibility="collapsed",
                                   format="%.2f")
        rows.append((aid, noc, prem))

    st.markdown("---")
    if st.button("💾  Save"):
        # entered_by is INTEGER FK — omit it (NULL is fine, column allows NULL per DDL)
        stmts = [
            ("""INSERT INTO daily_production
                    (agent_id, production_date, noc_count, monthly_premium,
                     entered_at, updated_at)
                VALUES (:aid, :dt, :noc, :prem, NOW(), NOW())
                ON CONFLICT (agent_id, production_date) DO UPDATE SET
                    noc_count       = EXCLUDED.noc_count,
                    monthly_premium = EXCLUDED.monthly_premium,
                    updated_at      = NOW()""",
             {"aid": aid, "dt": entry_date, "noc": noc, "prem": prem})
            for aid, noc, prem in rows if noc > 0 or prem > 0
        ]
        if stmts:
            run_write_many(stmts)
            refresh_views()
            st.success(f"✅  {len(stmts)} record(s) saved.")
            st.rerun()
        else:
            st.info("Nothing to save — all values are zero.")
