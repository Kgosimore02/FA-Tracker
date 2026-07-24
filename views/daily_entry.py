"""Daily NOC Entry — NOC count + monthly premium + weekly appointments per FA per day."""
import datetime
import streamlit as st
from db import run_query, run_query_live, run_write, run_write_many, refresh_views
from ui import page_header

def _week_of_month(d: datetime.date) -> int:
    """Return week number (1-4) for a date within its month."""
    return min(4, (d.day - 1) // 7 + 1)

def render(branch: dict):
    page_header("Daily NOC Entry",
                f"{branch.get('channel_code','')} — {branch.get('branch_name','')}")
    branch_id = branch["branch_id"]

    col_date, _ = st.columns([2, 5])
    with col_date:
        entry_date = st.date_input("Date", datetime.date.today(),
                                   max_value=datetime.date.today())

    week_num  = _week_of_month(entry_date)
    year      = entry_date.year
    month     = entry_date.month

    st.markdown(
        f'<p style="color:rgba(255,255,255,.5);font-size:.85rem">'
        f'{entry_date.strftime("%A, %d %B %Y")} &nbsp;·&nbsp; Week {week_num} of {entry_date.strftime("%B")}</p>',
        unsafe_allow_html=True,
    )

    agents = run_query(
        "SELECT agent_id, full_name FROM field_agents "
        "WHERE branch_id=:bid AND is_active=TRUE ORDER BY full_name",
        {"bid": branch_id},
    )
    if not agents:
        st.warning("No active FAs. Add agents in Agent Management.")
        return

    agent_ids = [a["agent_id"] for a in agents]

    # Load existing daily production for selected date
    existing = run_query(
        "SELECT agent_id, noc_count, monthly_premium FROM daily_production "
        "WHERE agent_id = ANY(ARRAY[:ids]::int[]) AND production_date = :dt",
        {"ids": agent_ids, "dt": entry_date},
    ) if agent_ids else []
    exist_map = {r["agent_id"]: r for r in existing}

    # Load existing weekly appointments for current week
    appt_rows = run_query_live(
        "SELECT agent_id, appointments FROM fa_weekly_appointments "
        "WHERE agent_id = ANY(ARRAY[:ids]::int[]) "
        "AND year=:yr AND month=:mo AND week_number=:wk",
        {"ids": agent_ids, "yr": year, "mo": month, "wk": week_num},
    ) if agent_ids else []
    appt_map = {r["agent_id"]: r["appointments"] for r in appt_rows}

    st.markdown("---")

    # Headers
    h1, h2, h3, h4 = st.columns([4, 2, 2.5, 2])
    h1.markdown("**FA Name**")
    h2.markdown("**NOC Cases**")
    h3.markdown("**Monthly Premium (BWP)**")
    h4.markdown(f"**Appointments (Wk {week_num})**")

    rows = []
    for ag in agents:
        aid  = ag["agent_id"]
        prev = exist_map.get(aid, {})
        c1, c2, c3, c4 = st.columns([4, 2, 2.5, 2])
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
        with c4:
            appt = st.number_input("appt", 0, 9999,
                                   int(appt_map.get(aid, 0)),
                                   key=f"a{aid}", label_visibility="collapsed")
        rows.append((aid, noc, prem, appt))

    st.markdown("---")

    if st.button("💾  Save"):
        # 1. Save daily production (NOC + premium) — only non-zero rows
        prod_stmts = [
            ("""INSERT INTO daily_production
                    (agent_id, production_date, noc_count, monthly_premium,
                     entered_at, updated_at)
                VALUES (:aid, :dt, :noc, :prem, NOW(), NOW())
                ON CONFLICT (agent_id, production_date) DO UPDATE SET
                    noc_count       = EXCLUDED.noc_count,
                    monthly_premium = EXCLUDED.monthly_premium,
                    updated_at      = NOW()""",
             {"aid": aid, "dt": entry_date, "noc": noc, "prem": prem})
            for aid, noc, prem, _ in rows if noc > 0 or prem > 0
        ]

        # 2. Save weekly appointments per FA (all rows — 0 is valid)
        appt_stmts = [
            ("""INSERT INTO fa_weekly_appointments
                    (agent_id, year, month, week_number, appointments, entered_at)
                VALUES (:aid, :yr, :mo, :wk, :appt, NOW())
                ON CONFLICT (agent_id, year, month, week_number) DO UPDATE SET
                    appointments = EXCLUDED.appointments,
                    entered_at   = NOW()""",
             {"aid": aid, "yr": year, "mo": month, "wk": week_num, "appt": appt})
            for aid, _, _, appt in rows
        ]

        all_stmts = prod_stmts + appt_stmts

        if all_stmts:
            run_write_many(all_stmts)

            # 3. Sum FA appointments for this branch + week → upsert weekly_appointments
            total_appts = sum(appt for _, _, _, appt in rows)
            run_write(
                """INSERT INTO weekly_appointments
                       (branch_id, year, month, week_number, appointments_attained, entered_at)
                   VALUES (:bid, :yr, :mo, :wk, :att, NOW())
                   ON CONFLICT (branch_id, year, month, week_number) DO UPDATE SET
                       appointments_attained = EXCLUDED.appointments_attained,
                       entered_at            = NOW()""",
                {"bid": branch_id, "yr": year, "mo": month,
                 "wk": week_num, "att": total_appts}
            )

            refresh_views()
            st.success(f"✅  Saved. {len(prod_stmts)} production record(s) · "
                       f"Total appointments Week {week_num}: {total_appts}")
            st.rerun()
        else:
            st.info("Nothing to save — all values are zero.")
