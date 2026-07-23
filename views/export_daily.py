"""Export — Daily NOC + Premium Report."""
import io, datetime, streamlit as st, openpyxl
from db import run_query
from ui import page_header
from views._xlsx import title_row, hdr, cell, total_cell, col_widths, TEAL_FILL

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

def render(branch: dict):
    page_header("Daily Report", "NOC cases and monthly premium for any date range")
    bid = branch["branch_id"]
    ch  = branch.get("channel_code",""); bn = branch.get("branch_name","")

    c1, c2 = st.columns(2)
    sd = c1.date_input("From", datetime.date.today().replace(day=1))
    ed = c2.date_input("To",   datetime.date.today())

    if st.button("📥  Generate"):
        rows = run_query(
            """SELECT fa.full_name, dp.production_date, dp.noc_count, dp.monthly_premium
               FROM daily_production dp
               JOIN field_agents fa ON fa.agent_id=dp.agent_id
               WHERE fa.branch_id=:bid AND dp.production_date BETWEEN :sd AND :ed
               ORDER BY dp.production_date, fa.full_name""",
            {"bid": bid, "sd": sd, "ed": ed},
        )
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Daily"
        title_row(ws, f"Daily Report — {ch} {bn}  |  {sd} to {ed}", "A1:D1")
        for ci, h in enumerate(["FA Name","Date","NOC Cases","Monthly Premium (BWP)"], 1):
            hdr(ws, 2, ci, h)
        tnoc = 0; tprem = 0.0
        for ri, r in enumerate(rows, 3):
            cell(ws, ri, 1, r["full_name"])
            cell(ws, ri, 2, r["production_date"])
            cell(ws, ri, 3, int(r["noc_count"] or 0), align=None)
            cell(ws, ri, 4, float(r["monthly_premium"] or 0), "#,##0.00")
            tnoc += r["noc_count"] or 0; tprem += float(r["monthly_premium"] or 0)
        tr = len(rows)+3
        total_cell(ws, tr, 1, "TOTAL"); total_cell(ws, tr, 2, "")
        total_cell(ws, tr, 3, tnoc); total_cell(ws, tr, 4, tprem, "#,##0.00")
        col_widths(ws, [("A",28),("B",14),("C",12),("D",22)])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        st.download_button("⬇️  Download",buf,f"daily_{ch}_{sd}_{ed}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if rows:
            m1,m2,m3 = st.columns(3)
            m1.metric("Records",   len(rows))
            m2.metric("Total NOC", tnoc)
            m3.metric("Total Premium", f"BWP {tprem:,.2f}")
        else:
            st.warning("No data for selected range.")
