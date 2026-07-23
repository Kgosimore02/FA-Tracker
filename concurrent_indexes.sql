-- Run this once on the PostgreSQL server BEFORE deploying the multi-user build.
-- These unique indexes are required for REFRESH MATERIALIZED VIEW CONCURRENTLY
-- which allows reads to continue unblocked during view refresh.

CREATE UNIQUE INDEX IF NOT EXISTS uidx_mv_daily_summary
    ON mv_daily_summary (branch_id, production_date);

CREATE UNIQUE INDEX IF NOT EXISTS uidx_mv_weekly_appt
    ON mv_weekly_appointment_summary (branch_id, year, month, week_number);

CREATE UNIQUE INDEX IF NOT EXISTS uidx_mv_monthly_branch
    ON mv_monthly_branch_summary (branch_id, budget_year, budget_month);

CREATE UNIQUE INDEX IF NOT EXISTS uidx_mv_quarterly
    ON mv_quarterly_summary (branch_id, budget_year, quarter);
