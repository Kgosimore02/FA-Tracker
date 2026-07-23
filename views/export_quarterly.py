"""Export — Quarterly Summary.

MV columns (from actual DDL):
  year, quarter, noc_actual, api_actual_annualised,
  noc_budget, quarterly_premium_budget,
  noc_attainment, api_attainment  (ratios 0–1)
"""
import io, streamlit as st, openpyxl
from db import run_query
from ui import page_header
from views._xlsx import title_row, hdr, cell, col_widths

def render(branch: dict):
    page_header("Quarterly Report", "Quarterly NOC and API rollup")
    bid = branch["branch_id"]
    ch  = branch.get("channel_code",""); bn = branch.get("branch_name","")

    c1, c2 = st.columns(2)
    year = c1.selectbox("Year", [2026, 2027])
    qtr  = c2.selectbox("Quarter", [1,2,3,4], format_func=lambda x: f"Q{x}")

    if st.button("📥  Generate"):
        rows = run_query(
            """SELECT noc_actual, noc_budget, noc_attainment,
                      api_actual_annualised, quarterly_premium_budget, api_attainment
               FROM mv_quarterly_summary
               WHERE branch_id=:bid AND year=:yr AND quarter=:q""",
            {"bid": bid, "yr": year, "q": qtr},
        )

        wb = openpyxl.Workbook(); ws = wb.active; ws.title = f"Q{qtr}"
        title_row(ws, f"Quarterly Report — {ch} {bn}  |  Q{qtr} {year}", "A1:F1")
        for ci, h in enumerate(["NOC Actual","NOC Budget","NOC Att %",
                                 "API Actual (BWP)","Qtrly API Budget (BWP)","API Att %"], 1):
            hdr(ws, 2, ci, h)

        if rows:
            r = rows[0]
            cell(ws, 3, 1, int(r["noc_actual"] or 0))
            cell(ws, 3, 2, float(r["noc_budget"] or 0), "0.0")
            cell(ws, 3, 3, float(r["noc_attainment"] or 0), "0.0%")
            cell(ws, 3, 4, float(r["api_actual_annualised"] or 0), "#,##0.00")
            cell(ws, 3, 5, float(r["quarterly_premium_budget"] or 0), "#,##0.00")
            cell(ws, 3, 6, float(r["api_attainment"] or 0), "0.0%")

        col_widths(ws, [("A",14),("B",12),("C",10),("D",18),("E",22),("F",10)])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        st.download_button("⬇️  Download", buf, f"quarterly_{ch}_{year}_Q{qtr}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if rows:
            r = rows[0]
            m1,m2,m3,m4 = st.columns(4)
            m1.metric("NOC Actual",     int(r["noc_actual"] or 0))
            m2.metric("NOC Attainment", f"{float(r['noc_attainment'] or 0)*100:.1f}%")
            m3.metric("API Actual",     f"BWP {float(r['api_actual_annualised'] or 0):,.0f}")
            m4.metric("API Attainment", f"{float(r['api_attainment'] or 0)*100:.1f}%")
        else:
            st.warning("No data for this quarter.")
