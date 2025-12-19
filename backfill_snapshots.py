import argparse
import sys
from datetime import date, timedelta

import pandas as pd

# Reuse logic from the main app.
from web_app import (
    authenticate_app_only,
    fetch_calendar_events,
    summarize_vacation,
    save_week_snapshot,
)


def week_bounds(iso_year: int, iso_week: int):
    """Return start (Mon) and end (Sun) dates for the ISO week."""
    start = date.fromisocalendar(iso_year, iso_week, 1)
    end = start + timedelta(days=6)
    return start, end


def last_iso_week_of_year(iso_year: int) -> int:
    """ISO weeks: Dec 28 is always in the last ISO week of the year."""
    return date(iso_year, 12, 28).isocalendar().week


def run_backfill(year: int, through_today: bool = True, target_user: str | None = None):
    if not target_user:
        from web_app import TARGET_MAILBOX

        target_user = TARGET_MAILBOX

    headers = authenticate_app_only()
    if not headers:
        print("ERROR: App-only authentication failed. Ensure TENANT_ID, CLIENT_ID, CLIENT_SECRET, and SCOPE are set.")
        sys.exit(1)
    if not target_user:
        print("ERROR: TARGET_MAILBOX (or --target-user) is required for app-only calendar access.")
        sys.exit(1)

    max_week = last_iso_week_of_year(year)
    today = date.today()
    if through_today and today.isocalendar().year == year:
        max_week = min(max_week, today.isocalendar().week)

    for week in range(1, max_week + 1):
        start_date, end_date = week_bounds(year, week)
        print(f"Fetching ISO week {year}-W{week:02d} ({start_date} to {end_date})...")
        events_df, stats = fetch_calendar_events(headers, start_date, end_date, include_all=False, target_user=target_user)
        print(f"  Graph returned {stats.get('total_events', 0)} events; matched {stats.get('matched_events', 0)} GO days.")
        summary_df, updated_events_df = summarize_vacation(events_df, start_date, end_date)
        save_week_snapshot(start_date, end_date, summary_df, updated_events_df)
        print(f"  Saved snapshot for {year}-W{week:02d}.")

    print("Backfill complete.")


def main():
    parser = argparse.ArgumentParser(description="Backfill weekly vacation snapshots for a given ISO year.")
    parser.add_argument("year", type=int, help="ISO year to backfill (e.g., 2025)")
    parser.add_argument(
        "--through-today",
        action="store_true",
        default=False,
        help="Stop at the current week if the year is the current year (default: backfill entire year).",
    )
    parser.add_argument(
        "--target-user",
        type=str,
        default=None,
        help="Mailbox UPN to read (defaults to TARGET_MAILBOX/TARGET_USER/USER_UPN env vars).",
    )
    args = parser.parse_args()
    run_backfill(args.year, through_today=args.through_today, target_user=args.target_user)


if __name__ == "__main__":
    main()
