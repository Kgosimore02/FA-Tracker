"""Budget Setup — monthly NOC + Annual Premium Income per branch."""
import streamlit as st, pandas as pd
from db import run_query, run_write_many, refresh_views
from ui import page_header

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def render(branch: dict):
    page_header("Budget Setup",
                f"{branch.get('channel_code','')} — {branch.get('branch_name','')}")
    bid  = branch["branch_id"]
    year = st.selectbox("Year", [2026, 2027])

    existing = run_query(
        "SELECT budget_month, noc_budget, annual_premium_budget "
        "FROM monthly_budgets WHERE branch_id=:bid AND budget_year=:yr ORDER BY budget_month",
        {"bid": bid, "yr": year},
    )
    bmap = {r["budget_month"]: r for r in existing}

    st.markdown("NOC Budget = cases. Annual Premium = annualised figure "
                "(app divides by 12 for monthly comparison).")
    st.markdown("---")

    h0, h1, h2 = st.columns([1.5, 2, 2.5])
    h0.markdown("**Month**"); h1.markdown("**NOC Budget**")
    h2.markdown("**Annual Premium (BWP)**")

    entries = []
    for i, mon in enumerate(MONTHS, 1):
        prev = bmap.get(i, {})
        c0, c1, c2 = st.columns([1.5, 2, 2.5])
        with c0:
            st.markdown(f'<div style="padding:.4rem 0">{mon}</div>', unsafe_allow_html=True)
        with c1:
            noc = st.number_input("n", 0, 9999, int(prev.get("noc_budget", 0)),
                                  key=f"bn{i}", label_visibility="collapsed")
        with c2:
            api = st.number_input("a", 0.0, step=100.0,
                                  value=float(prev.get("annual_premium_budget", 0.0)),
                                  key=f"ba{i}", label_visibility="collapsed", format="%.2f")
        entries.append((i, noc, api))

    st.markdown("---")
    if st.button("💾  Save Budgets"):
        # entered_by is INTEGER FK — omit it (NULL allowed per DDL)
        stmts = [
            ("""INSERT INTO monthly_budgets
                    (branch_id, budget_year, budget_month, noc_budget, annual_premium_budget,
                     entered_at)
                VALUES (:bid, :yr, :mo, :noc, :api, NOW())
                ON CONFLICT (branch_id, budget_year, budget_month) DO UPDATE SET
                    noc_budget            = EXCLUDED.noc_budget,
                    annual_premium_budget = EXCLUDED.annual_premium_budget,
                    entered_at            = NOW()""",
             {"bid": bid, "yr": year, "mo": m, "noc": n, "api": a})
            for m, n, a in entries if n > 0 or a > 0
        ]
        if stmts:
            run_write_many(stmts)
            refresh_views()
            st.success("✅  Budgets saved.")
            st.rerun()

    if existing:
        st.markdown("### Saved Budgets")
        df = pd.DataFrame(existing)
        df["budget_month"]       = df["budget_month"].apply(lambda x: MONTHS[x-1])
        df["monthly_api"]        = (df["annual_premium_budget"] / 12).map("{:,.2f}".format)
        df["annual_premium_budget"] = df["annual_premium_budget"].map("{:,.2f}".format)
        df.columns = ["Month", "NOC Budget", "Annual Premium (BWP)", "Monthly API (BWP)"]
        st.dataframe(df, use_container_width=True, hide_index=True)
