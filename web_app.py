import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from azure.identity import DeviceCodeCredential, ClientSecretCredential
#from msgraph.core import GraphClient
import requests
import re
import openpyxl
import io
import threading
import time
import holidays
from dotenv import load_dotenv
import os
from datetime import date
from pathlib import Path
from urllib.parse import quote
from collections import defaultdict

#nest_asyncio.apply()
load_dotenv() 

# TENANT_ID = os.getenv("TENANT_ID")
# CLIENT_ID = os.getenv("CLIENT_ID")
# SCOPE = os.getenv("SCOPE", "https://graph.microsoft.com/.default")

# DEVICE_CODE_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/devicecode"
# TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

def _get_conf(key, default=None):
    try:
        return st.secrets.get(key, None) or os.getenv(key, default)
    except Exception:
        return os.getenv(key, default)

TENANT_ID = _get_conf("TENANT_ID")
CLIENT_ID = _get_conf("CLIENT_ID")
# For Graph calendar calls you usually want delegated scopes:
# e.g. "User.Read Calendars.Read"
SCOPE = _get_conf("SCOPE", "https://graph.microsoft.com/.default")
CLIENT_SECRET = _get_conf("CLIENT_SECRET")
TARGET_MAILBOX = _get_conf("TARGET_MAILBOX") or _get_conf("TARGET_USER") or _get_conf("USER_UPN")

DEVICE_CODE_URL = f"https://login.microsoftonline.com/{TENANT_ID or 'common'}/oauth2/v2.0/devicecode"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID or 'common'}/oauth2/v2.0/token"

@st.cache_resource
def authenticate_device_flow():
    # Step 1: Request device code
    response = requests.post(DEVICE_CODE_URL, data={
        "client_id": CLIENT_ID,
        "scope": SCOPE
    })

    if response.status_code != 200:
        st.error("Failed to initiate device code flow.")
        return None

    data = response.json()

    st.info("Microsoft Login Required")
    st.markdown(f" [Click here to log in]({data['verification_uri']})", unsafe_allow_html=True)
    st.code(f"Enter this code: {data['user_code']}", language="text")

    device_code = data["device_code"]
    interval = data["interval"]

    # Step 2: Poll for access token
    with st.spinner("Waiting for authentication..."):
        for _ in range(60):  # Wait up to ~5 minutes (60 * interval)
            time.sleep(interval)
            token_response = requests.post(TOKEN_URL, data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": CLIENT_ID,
                "device_code": device_code
            })

            if token_response.status_code == 200:
                token_data = token_response.json()
                access_token = token_data["access_token"]
                st.success("Authentication successful.")
                return {
                    "Authorization": f"Bearer {access_token}",
                    "Accept": "application/json"
                }

            elif token_response.status_code == 400:
                error = token_response.json().get("error")
                if error in ["authorization_pending", "slow_down"]:
                    continue
                else:
                    st.error(f" Authentication failed: {error}")
                    return None

        st.error(" Authentication timed out.")
        return None



    # Return headers to be used in Graph API calls
    headers = {
        "Authorization": f"Bearer {login_info['access_token']}",
        "Accept": "application/json"
    }
    st.success("✅ Authentication successful.")
    return headers


def authenticate_app_only():
    """Authenticate with client credentials (no user prompt) when CLIENT_SECRET is configured."""
    if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
        return None
    try:
        credential = ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET,
        )
        token = credential.get_token(SCOPE)
        return {
            "Authorization": f"Bearer {token.token}",
            "Accept": "application/json"
        }
    except Exception as exc:
        st.warning(f"App-only auth failed, falling back to device code: {exc}")
        return None


hr_holidays = holidays.country_holidays("HR")
def is_working_day(date):
    return date.weekday() < 5 and date not in hr_holidays


def load_allowances():
    try:
        allowance_df = pd.read_csv("vacation_allowances.csv")
        return {
            row["Name"]: int(row["Allowance"])
            for _, row in allowance_df.iterrows()
            if pd.notna(row.get("Name")) and pd.notna(row.get("Allowance"))
        }
    except FileNotFoundError:
        st.warning("'vacation_allowances.csv' not found. Defaulting to 25 days for everyone.")
        return {}
    except Exception as exc:
        st.error(f"Failed to read 'vacation_allowances.csv': {exc}")
        return {}


def load_carryover():
    try:
        carryover_df = pd.read_csv("vacation_carryover.csv")
    except FileNotFoundError:
        st.info("No carryover file found. Add 'vacation_carryover.csv' to track previous-year days.")
        return {}
    except Exception as exc:
        st.error(f"Failed to read 'vacation_carryover.csv': {exc}")
        return {}

    carryover_map = {}
    for _, row in carryover_df.iterrows():
        try:
            name = row["Name"]
            year = int(row["Year"])
            carry = int(row["Carryover"])
            carryover_map[(name, year)] = carry
        except Exception:
            continue
    return carryover_map


def get_week_key(end_date):
    iso_year, iso_week, _ = end_date.isocalendar()
    return iso_year, iso_week


def snapshot_paths():
    base = Path("snapshots")
    return base, base / "summary_weekly.csv", base / "events_weekly.csv"


def load_week_snapshot(end_date):
    _, summary_path, events_path = snapshot_paths()
    iso_year, iso_week = get_week_key(end_date)
    if not summary_path.exists():
        return None, None
    try:
        summary_df = pd.read_csv(summary_path)
        week_summary = summary_df[(summary_df["Year"] == iso_year) & (summary_df["Week"] == iso_week)]
        if week_summary.empty:
            return None, None
        events_df = pd.DataFrame()
        if events_path.exists():
            events_df = pd.read_csv(events_path)
            events_df = events_df[(events_df["Year"] == iso_year) & (events_df["Week"] == iso_week)]
        return week_summary.reset_index(drop=True), events_df.reset_index(drop=True)
    except Exception as exc:
        st.warning(f"Failed to load cached snapshot: {exc}")
        return None, None


def save_week_snapshot(start_date, end_date, summary_df, events_df):
    base, summary_path, events_path = snapshot_paths()
    base.mkdir(exist_ok=True)
    iso_year, iso_week = get_week_key(end_date)
    start_str = pd.to_datetime(start_date).date().isoformat()
    end_str = pd.to_datetime(end_date).date().isoformat()
    fetched_at = datetime.utcnow().isoformat()

    def persist(path, df):
        df_copy = df.copy()
        df_copy["Year"] = iso_year
        df_copy["Week"] = iso_week
        df_copy["Period Start"] = start_str
        df_copy["Period End"] = end_str
        df_copy["Fetched At"] = fetched_at

        if path.exists():
            existing = pd.read_csv(path)
            # Drop existing rows for the same week to avoid duplicates.
            existing = existing[(existing["Year"] != iso_year) | (existing["Week"] != iso_week)]
            combined = pd.concat([existing, df_copy], ignore_index=True)
        else:
            combined = df_copy

        combined.to_csv(path, index=False)

    persist(summary_path, summary_df)
    persist(events_path, events_df)


def snapshot_fetched_at(end_date):
    _, summary_path, _ = snapshot_paths()
    if not summary_path.exists():
        return None
    iso_year, iso_week = get_week_key(end_date)
    try:
        meta_df = pd.read_csv(summary_path, usecols=["Year", "Week", "Fetched At"])
        row = meta_df[(meta_df["Year"] == iso_year) & (meta_df["Week"] == iso_week)]
        if row.empty:
            return None
        return pd.to_datetime(row.iloc[0]["Fetched At"], errors="coerce")
    except Exception:
        return None


def is_snapshot_stale(end_date, max_age_days=7):
    fetched_at = snapshot_fetched_at(end_date)
    if fetched_at is None or pd.isna(fetched_at):
        return True
    age = datetime.utcnow() - fetched_at.to_pydatetime()
    return age > timedelta(days=max_age_days)

def fetch_calendar_events(headers, start_datetime, end_datetime, include_all=False, target_user=None):
    start = datetime.combine(start_datetime, datetime.min.time()).astimezone(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
    end = datetime.combine(end_datetime, datetime.max.time()).astimezone(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')

    if target_user:
        user_part = quote(target_user)
        url = f"https://graph.microsoft.com/v1.0/users/{user_part}/calendar/calendarView?startDateTime={start}&endDateTime={end}"
    else:
        url = f"https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime={start}&endDateTime={end}"
    all_events = []

    error_msg = None
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            error_msg = f"Graph API error {response.status_code}: {response.text}"
            st.error(error_msg)
            break

        result = response.json()
        events = result.get("value", [])
        all_events.extend(events)
        url = result.get("@odata.nextLink", None)

    vacation_rows = []

    def is_go_event(subject, categories):
        subject = subject or ""
        categories = categories or []
        # Match GO/go/g.o/g-o and TBC variants with flexible spacing/punctuation.
        go_pattern = r'(?i)\bg[\s\.\-]*o\b(?:\s*[\-:]*\s*\(?tbc\)?)?'
        subject_match = bool(re.search(go_pattern, subject))
        category_match = any(cat.lower() == "go" for cat in categories if isinstance(cat, str))
        return subject_match or category_match

    for ev in all_events:
        subject = ev.get("subject", "")
        categories = ev.get("categories", [])
        if (include_all or is_go_event(subject, categories)) and not re.search(r'\bcanceled\b|\botkazano\b', subject, re.IGNORECASE):
            organizer = ev.get("organizer", {}).get("emailAddress", {}).get("name", "Unknown")
            start_dt = pd.to_datetime(ev.get("start", {}).get("dateTime"))
            end_dt = pd.to_datetime(ev.get("end", {}).get("dateTime"))

            # Treat end as exclusive only when it is exactly at midnight the next day (all-day events).
            if end_dt.time() == datetime.min.time() and end_dt.date() > start_dt.date():
                end_dt -= timedelta(days=1)

            start_day = start_dt.normalize()
            end_day = end_dt.normalize()
            all_dates = pd.date_range(start=start_day, end=end_day, freq="D")

            for d in all_dates:
                date_only = d.date()
                if is_working_day(date_only) and (start_datetime <= date_only <= end_datetime):
                    vacation_rows.append({
                        "Name": organizer,
                        "Date": date_only,
                        "Weekday": date_only.strftime("%A"),
                        "Subject": subject,
                        "Categories": ", ".join(categories) if categories else ""
                    })
    # Ensure consistent columns even when no events match.
    if error_msg:
        stats = {
            "total_events": len(all_events),
            "matched_events": 0,
            "sample_events": [],
            "error": error_msg,
        }
        return None, stats

    matched_df = pd.DataFrame(vacation_rows, columns=["Name", "Date", "Weekday", "Subject", "Categories"])
    sample_events = []
    for ev in all_events[:5]:
        sample_events.append({
            "Subject": ev.get("subject", ""),
            "Categories": ", ".join(ev.get("categories", [])) if ev.get("categories") else "",
            "Start": ev.get("start", {}).get("dateTime", ""),
            "End": ev.get("end", {}).get("dateTime", ""),
        })

    stats = {
        "total_events": len(all_events),
        "matched_events": len(vacation_rows),
        "sample_events": sample_events,
        "error": None,
    }
    return matched_df, stats

# def get_calendar_events(graph_client, start_datetime, end_datetime, include_all=False, target_user=None):
#     return asyncio.get_event_loop().run_until_complete(
#         fetch_calendar_events(graph_client, start_datetime, end_datetime, include_all=include_all, target_user=target_user)
#     )
def get_calendar_events(graph_client, start_datetime, end_datetime, include_all=False, target_user=None):
    df, stats = fetch_calendar_events(graph_client, start_datetime, end_datetime, include_all=include_all, target_user=target_user)
    if df is None:
        return pd.DataFrame(columns=["Name", "Date", "Weekday", "Subject", "Categories"]), stats or {"total_events": 0, "matched_events": 0, "error": "Unknown error"}
    return df, stats


def summarize_vacation(events_df, start_date, end_date):
    if events_df is None or "Date" not in events_df.columns:
        st.warning("No events with dates found to summarize.")
        return pd.DataFrame(), pd.DataFrame(columns=["Name", "Date", "Weekday", "Subject", "Categories"])

    events_df = events_df[(events_df["Date"] >= start_date) & (events_df["Date"] <= end_date)]

    allowances = load_allowances()
    carryover = load_carryover()

    # Determine which names to include in the report.
    if allowances:
        allowed_names = set(allowances.keys())
        before_count = len(events_df)
        names_before = set(events_df["Name"].unique())
        events_df = events_df[events_df["Name"].isin(allowed_names)]
        removed = before_count - len(events_df)
        if removed > 0:
            dropped = sorted(names_before - allowed_names)
            st.info(f"Filtered out {removed} event rows for people not in vacation_allowances.csv.")
            if dropped:
                st.caption(f"Skipped names: {', '.join(dropped)}")
        report_names = allowed_names
    else:
        report_names = set(events_df["Name"].unique())

    if events_df.empty and not report_names:
        return pd.DataFrame(), events_df

    enriched_rows = []
    summary_rows = []
    target_year = end_date.year
    carryover_year = target_year - 1
    carryover_expiry = date(target_year, 6, 30)

    for name in sorted(report_names):
        group = events_df[events_df["Name"] == name].sort_values("Date").copy()
        base_allowance = allowances.get(name, 25)
        if not group.empty:
            group["Year"] = group["Date"].apply(lambda d: d.year)
            group["Allowance Year"] = group["Year"]
            group["Used Status"] = None
            group["Carryover Window"] = ""
        else:
            # Ensure expected columns even with no events.
            group = pd.DataFrame(columns=["Name", "Date", "Weekday", "Subject", "Categories", "Year", "Allowance Year", "Used Status", "Carryover Window"])

        usage_tracker = defaultdict(int)
        usage_by_year = defaultdict(int)

        for idx, row in group.iterrows():
            year = row["Year"]
            usage_tracker[year] += 1
            usage_by_year[year] += 1
            status = "Within Base Allowance" if usage_tracker[year] <= base_allowance else "Over Base Allowance"
            group.at[idx, "Used Status"] = status

            if year == target_year and row["Date"] <= carryover_expiry:
                group.at[idx, "Carryover Window"] = f"Eligible to use carryover until {carryover_expiry.isoformat()}"

        used_total = len(group)
        over_base = any(used > base_allowance for used in usage_by_year.values())
        carryover_days = carryover.get((name, carryover_year), 0)
        if carryover_days > 0:
            if end_date <= carryover_expiry:
                carryover_status = f"Active until {carryover_expiry.isoformat()}"
                carryover_usable = "Yes"
            else:
                carryover_status = f"Expired after {carryover_expiry.isoformat()} (not usable)"
                carryover_usable = "No (expired)"
        else:
            carryover_status = "No carryover recorded"
            carryover_usable = "No"

        summary = {
            "Name": name,
            "Base Allowance": base_allowance,
            f"Carryover from {carryover_year}": carryover_days,
            "Carryover Status": carryover_status,
            "Carryover Usable?": carryover_usable,
            f"Used {target_year}": usage_by_year.get(target_year, 0),
            "Used Total": used_total,
            "⚠️ Over Base Limit?": "Yes" if over_base else "No"
        }

        for y in sorted(usage_by_year.keys()):
            summary[f"Used {y}"] = usage_by_year[y]
            summary[f"Remaining {y} (base only)"] = max(base_allowance - usage_by_year[y], 0)

        summary_rows.append(summary)
        enriched_rows.append(group)

    summary_df = pd.DataFrame(summary_rows)
    updated_events_df = pd.concat(enriched_rows, ignore_index=True)

    return summary_df, updated_events_df


# --- Streamlit UI ---
st.title("Vacation Tracker - GO Events Summary")

#graph_client = authenticate_graph()
# headers = authenticate_graph()

def get_auth_headers():
    """Persist headers and auth mode in session so the login prompts disappear after success."""
    if "auth_headers" in st.session_state and "auth_mode" in st.session_state:
        return st.session_state["auth_headers"], st.session_state["auth_mode"]

    # Prefer app-only auth when client secret is provided.
    headers = authenticate_app_only()
    if headers:
        st.success("Authenticated with app credentials (no prompt).")
        auth_mode = "app"
    else:
        headers = authenticate_device_flow()
        auth_mode = "delegated" if headers else None

    if headers:
        st.session_state["auth_headers"] = headers
        st.session_state["auth_mode"] = auth_mode
    return headers, auth_mode

headers, auth_mode = get_auth_headers()
if headers is None or auth_mode is None:
    st.stop()

with st.sidebar:
    st.header("Filter Settings")
    start_date = st.date_input("Start Date", datetime.now() - timedelta(days=30))
    end_date = st.date_input("End Date", datetime.now() + timedelta(days=30))
    fetch = st.button("Fetch and Calculate")

if fetch:
    with st.spinner("Fetching events and calculating..."):
        # For app-only auth you must target a specific mailbox.
        target_user = None
        if auth_mode == "app":
            target_user = TARGET_MAILBOX
            if not target_user:
                st.error("App-only auth requires TARGET_MAILBOX (UPN or user id) to fetch calendar data.")
                st.stop()

        events_df, stats = get_calendar_events(headers, start_date, end_date, include_all=False, target_user=target_user)

        if stats.get("error"):
            st.error("Calendar fetch failed.")
            st.stop()

        st.info(f"Graph returned {stats.get('total_events', 0)} events; matched {stats.get('matched_events', 0)} GO days.")
        if stats.get("matched_events", 0) == 0 and stats.get("total_events", 0) > 0:
            st.caption("First few returned events (subjects/categories) to diagnose GO matching:")
            st.dataframe(pd.DataFrame(stats.get("sample_events", [])))
        summary_df, updated_events_df = summarize_vacation(events_df, start_date, end_date)

        st.success("Done!")
        st.subheader("Vacation Summary")
        st.dataframe(summary_df)  # ✅ Only display the summary

        # ✅ Export both to Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            updated_events_df.to_excel(writer, sheet_name="Events", index=False)

        st.download_button("Download Excel Summary", buffer.getvalue(), file_name="vacation_summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
