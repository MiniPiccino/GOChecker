from datetime import date, timedelta
import sys

from web_app import authenticate_with_msal, fetch_calendar_events, summarize_vacation, save_week_snapshot


def current_week_bounds(today):
    iso_year, iso_week, _ = today.isocalendar()
    start = date.fromisocalendar(iso_year, iso_week, 1)
    end = start + timedelta(days=6)
    return start, end, iso_year, iso_week


def main():
    headers = authenticate_with_msal(silent_only=True)
    if not headers:
        print("ERROR: No cached token. Run the web app once to sign in, or run this script with interactive login.")
        sys.exit(1)

    today = date.today()
    start_date, end_date, iso_year, iso_week = current_week_bounds(today)
    events_df, stats = fetch_calendar_events(headers, start_date, end_date, include_all=False, target_user=None)
    if stats.get("error"):
        print(f"ERROR: {stats['error']}")
        sys.exit(1)
    print(f"Graph returned {stats.get('total_events', 0)} events; matched {stats.get('matched_events', 0)} GO days.")

    summary_df, updated_events_df = summarize_vacation(events_df, start_date, end_date)
    save_week_snapshot(start_date, end_date, summary_df, updated_events_df)
    print(f"Saved snapshot for {iso_year}-W{iso_week:02d}.")


if __name__ == "__main__":
    main()
