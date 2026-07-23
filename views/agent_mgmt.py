"""Agent Management — add / deactivate / reactivate. Rank is always FA."""
import datetime, streamlit as st
from db import run_query, run_write
from ui import page_header

def render(branch: dict):
    page_header("Agent Management",
                f"{branch.get('channel_code','')} — {branch.get('branch_name','')}")
    bid = branch["branch_id"]

    active   = run_query("SELECT agent_id, full_name, start_date FROM field_agents "
                         "WHERE branch_id=:bid AND is_active=TRUE ORDER BY full_name", {"bid": bid})
    inactive = run_query("SELECT agent_id, full_name, start_date FROM field_agents "
                         "WHERE branch_id=:bid AND is_active=FALSE ORDER BY full_name", {"bid": bid})

    st.markdown("### Add New FA")
    with st.form("add_fa", clear_on_submit=True):
        name  = st.text_input("Full Name")
        start = st.date_input("Start Date", datetime.date.today())
        submitted = st.form_submit_button("➕  Add FA")

    # Handle outside form block so st.rerun() works correctly
    if submitted:
        if not name.strip():
            st.error("Enter a name.")
        else:
            run_write(
                "INSERT INTO field_agents (full_name, rank, branch_id, start_date, is_active) "
                "VALUES (:n, 'FA', :bid, :s, TRUE)",
                {"n": name.strip(), "bid": bid, "s": start},
            )
            st.success(f"✅  {name.strip()} added.")
            st.rerun()

    st.markdown("---")
    st.markdown(f"### Active FAs ({len(active)})")
    if not active:
        st.info("No active FAs.")
    else:
        c1h, c2h, c3h = st.columns([4, 2, 1.5])
        c1h.markdown("**Name**"); c2h.markdown("**Start Date**"); c3h.markdown("**Action**")
        for ag in active:
            c1, c2, c3 = st.columns([4, 2, 1.5])
            c1.write(ag["full_name"])
            c2.write(str(ag["start_date"]) if ag["start_date"] else "—")
            with c3:
                if st.button("Deactivate", key=f"d{ag['agent_id']}"):
                    run_write("UPDATE field_agents SET is_active=FALSE WHERE agent_id=:aid",
                              {"aid": ag["agent_id"]})
                    st.rerun()

    if inactive:
        st.markdown("---")
        st.markdown(f"### Inactive FAs ({len(inactive)})")
        for ag in inactive:
            c1, c2, c3 = st.columns([4, 2, 1.5])
            c1.markdown(f'<span style="color:rgba(255,255,255,.35)">{ag["full_name"]}</span>',
                        unsafe_allow_html=True)
            c2.write(str(ag["start_date"]) if ag["start_date"] else "—")
            with c3:
                if st.button("Reactivate", key=f"r{ag['agent_id']}"):
                    run_write("UPDATE field_agents SET is_active=TRUE WHERE agent_id=:aid",
                              {"aid": ag["agent_id"]})
                    st.rerun()
