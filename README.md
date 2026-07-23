# BLIL FA Productivity Tracker — v6

## Start
```bash
pip install -r requirements.txt
export DB_PASS=your_password   # or edit db.py
streamlit run app.py
```

## Structure
```
app.py            Entry point — no auth (beta)
db.py             SQLAlchemy pool + cached queries + view refresh
ui.py             BLIL styles, SVG logo, sidebar nav
views/
  daily_entry.py          NOC + premium entry per FA per day
  budget_setup.py         Monthly NOC + Annual Premium budget (upsert)
  agent_mgmt.py           Add / deactivate / reactivate FAs (rank=FA)
  export_daily.py         Date-range daily report → xlsx
  export_weekly.py        Weekly NOC vs weighted targets → xlsx
  export_monthly.py       Monthly NOC + API vs budget → xlsx
  export_quarterly.py     Quarterly rollup → xlsx
  export_appointments.py  FA×3 appointment summary (calculated) → xlsx
  _xlsx.py                Shared openpyxl helpers
assets/blil_logo.svg      BLIL logo (embedded in sidebar)
concurrent_indexes.sql    Run once on DB for non-blocking view refresh
```

## DB (db.py defaults)
```
Host : 10.1.112.109:5432
Name : FA_tracker
User : postgres
Pass : set via DB_PASS env var
```

## Key rules
- Weekly targets: W1=15% W2=20% W3=30% W4=35% (weekly_weight_config table)
- API actual = SUM(monthly_premium) × 12; monthly budget = annual_premium_budget / 12
- Appointment target = active_fa_count × 3 (never stored, always computed)
- Auth: login removed for beta — add auth.py back when ready
