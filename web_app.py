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
import msal
from collections import defaultdict
import calendar

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
AUTH_MODE = (_get_conf("AUTH_MODE", "delegated") or "delegated").lower()
DELEGATED_SCOPES = (_get_conf("DELEGATED_SCOPES", "User.Read Calendars.Read") or "").split()
MSAL_CACHE_PATH = _get_conf("MSAL_CACHE_PATH", ".msal_cache.bin")

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


def _load_msal_cache():
    cache = msal.SerializableTokenCache()
    cache_path = Path(MSAL_CACHE_PATH)
    if cache_path.exists():
        cache.deserialize(cache_path.read_text())
    return cache, cache_path


def _save_msal_cache(cache, cache_path):
    if cache.has_state_changed:
        cache_path.write_text(cache.serialize())


def authenticate_with_msal(silent_only=False):
    if not (TENANT_ID and CLIENT_ID):
        st.error("Missing TENANT_ID or CLIENT_ID for delegated auth.")
        return None

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    scopes = DELEGATED_SCOPES or ["User.Read", "Calendars.Read"]
    cache, cache_path = _load_msal_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=authority, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result and not silent_only:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            st.error("Failed to start device code flow.")
            return None
        st.info("Microsoft Login Required")
        st.markdown(f"[Click here to log in]({flow['verification_uri']})", unsafe_allow_html=True)
        st.code(f"Enter this code: {flow['user_code']}", language="text")
        result = app.acquire_token_by_device_flow(flow)

    _save_msal_cache(cache, cache_path)

    if result and "access_token" in result:
        return {
            "Authorization": f"Bearer {result['access_token']}",
            "Accept": "application/json",
        }

    if result and "error" in result and not silent_only:
        st.error(f"Authentication failed: {result.get('error_description', result['error'])}")
    return None


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


hr_holidays = holidays.country_holidays("HR", years=2025)
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


def load_range_snapshot(start_date, end_date):
    _, summary_path, events_path = snapshot_paths()
    if not summary_path.exists():
        return None, None

    try:
        summary_df = pd.read_csv(summary_path)
    except Exception as exc:
        st.warning(f"Failed to load cached summary snapshot: {exc}")
        summary_df = None

    if summary_df is not None and ("Period Start" not in summary_df.columns or "Period End" not in summary_df.columns):
        summary_df = None

    start_date = pd.to_datetime(start_date).date()
    end_date = pd.to_datetime(end_date).date()

    in_range = pd.DataFrame()
    if summary_df is not None:
        summary_df["Period Start"] = pd.to_datetime(summary_df["Period Start"], errors="coerce").dt.date
        summary_df["Period End"] = pd.to_datetime(summary_df["Period End"], errors="coerce").dt.date
        in_range = summary_df[
            (summary_df["Period End"] >= start_date) & (summary_df["Period Start"] <= end_date)
        ].copy()

    # Load and filter events within date range.
    events_df = pd.DataFrame()
    if events_path.exists():
        try:
            events_df = pd.read_csv(events_path)
            if "Date" in events_df.columns:
                events_df["Date"] = pd.to_datetime(events_df["Date"], errors="coerce").dt.date
                events_df = events_df[(events_df["Date"] >= start_date) & (events_df["Date"] <= end_date)]
                if "Name" in events_df.columns:
                    events_df = events_df.drop_duplicates(subset=["Name", "Date"])
        except Exception as exc:
            st.warning(f"Failed to load cached events snapshot: {exc}")

    summary_agg = in_range.reset_index(drop=True) if not in_range.empty else None
    return summary_agg, events_df.reset_index(drop=True)


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
        # Match GO/go/g.o/g-o and punctuation/spacing variants, plus optional TBC suffix.
        go_pattern = r'(?i)\bg[\s\.\-_/]*o\b[!?.]?(?:\s*[\-:]*\s*\(?tbc\)?)?'
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
    if not matched_df.empty:
        # Count at most one vacation day per person per date.
        matched_df = matched_df.drop_duplicates(subset=["Name", "Date"])
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
        "matched_events": len(matched_df),
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

        used_in_year = usage_by_year.get(target_year, 0)
        used_pre_june = int(((group["Year"] == target_year) & (group["Date"] <= carryover_expiry)).sum()) if not group.empty else 0
        carryover_used = min(carryover_days, used_pre_june)
        carryover_remaining = max(carryover_days - carryover_used, 0)
        base_used = max(used_in_year - carryover_used, 0)
        base_remaining = max(base_allowance - base_used, 0)
        remaining_total = base_remaining + (carryover_remaining if carryover_usable == "Yes" else 0)

        summary = {
            "Name": name,
            #"Base Allowance": base_allowance,
            f"Carryover from {carryover_year}": carryover_days,
            #"Carryover Status": carryover_status,
            #"Carryover Usable?": carryover_usable,
            #"Carryover Used": carryover_used,
            "Carryover Remaining": carryover_remaining,
            #"Base Remaining": base_remaining,
            "Remaining Total": remaining_total,
            f"Used {target_year}": usage_by_year.get(target_year, 0),
            #"Used Total": used_total,
            #"Over Base Limit?": "Yes" if over_base else "No"
        }

        #for y in sorted(usage_by_year.keys()):
        #    summary[f"Used {y}"] = usage_by_year[y]

        summary_rows.append(summary)
        enriched_rows.append(group)

    summary_df = pd.DataFrame(summary_rows)
    updated_events_df = pd.concat(enriched_rows, ignore_index=True)

    return summary_df, updated_events_df


def build_vacation_calendar(events_df, month_date):
    if events_df is None or events_df.empty:
        return pd.DataFrame()

    month_date = pd.to_datetime(month_date).date()
    year = month_date.year
    month = month_date.month
    events_df = events_df.copy()
    events_df["Date"] = pd.to_datetime(events_df["Date"]).dt.date
    month_events = events_df[(events_df["Date"] >= date(year, month, 1)) & (events_df["Date"] <= date(year, month, calendar.monthrange(year, month)[1]))]

    day_to_names = (
        month_events.groupby("Date")["Name"]
        .apply(lambda s: ", ".join(sorted(set(n for n in s if isinstance(n, str)))))
        .to_dict()
    )

    cal = calendar.Calendar(firstweekday=0)
    weeks = cal.monthdatescalendar(year, month)
    rows = []
    for week in weeks:
        row = []
        for d in week:
            if d.month != month:
                row.append("")
                continue
            names = day_to_names.get(d, "")
            cell = f"{d.day}\n{names}" if names else str(d.day)
            row.append(cell)
        rows.append(row)

    return pd.DataFrame(rows, columns=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"])


def build_vacation_calendar_html(events_df, month_date):
    if events_df is None or events_df.empty:
        return ""

    month_date = pd.to_datetime(month_date).date()
    year = month_date.year
    month = month_date.month
    events_df = events_df.copy()
    events_df["Date"] = pd.to_datetime(events_df["Date"]).dt.date
    month_events = events_df[(events_df["Date"] >= date(year, month, 1)) & (events_df["Date"] <= date(year, month, calendar.monthrange(year, month)[1]))]

    day_to_names = (
        month_events.groupby("Date")["Name"]
        .apply(lambda s: ", ".join(sorted(set(n for n in s if isinstance(n, str)))))
        .to_dict()
    )

    cal = calendar.Calendar(firstweekday=0)
    weeks = cal.monthdatescalendar(year, month)
    month_label = month_date.strftime("%B %Y")

    def cell_html(d):
        if d.month != month:
            return '<td class="day muted"></td>'
        names = day_to_names.get(d, "")
        names_html = f'<div class="names">{names}</div>' if names else ""
        has_vac = " has-vac" if names else ""
        return f'<td class="day{has_vac}"><div class="date">{d.day}</div>{names_html}</td>'

    rows_html = "\n".join(
        "<tr>" + "".join(cell_html(d) for d in week) + "</tr>"
        for week in weeks
    )

    return f"""
    <div class="vacation-calendar">
      <div class="calendar-header">{month_label}</div>
      <table>
        <thead>
          <tr>
            <th>Mon</th><th>Tue</th><th>Wed</th><th>Thu</th><th>Fri</th><th>Sat</th><th>Sun</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
        </tbody>
      </table>
    </div>
    """


# --- Streamlit UI ---
st.title("Vacation Tracker - GO Events Summary")

#graph_client = authenticate_graph()
# headers = authenticate_graph()

def get_auth_headers():
    """Persist headers and auth mode in session so the login prompts disappear after success."""
    if "auth_headers" in st.session_state and "auth_mode" in st.session_state:
        return st.session_state["auth_headers"], st.session_state["auth_mode"]

    if AUTH_MODE == "app":
        headers = authenticate_app_only()
        if headers:
            st.success("Authenticated with app credentials (no prompt).")
            auth_mode = "app"
        else:
            headers = authenticate_with_msal(silent_only=False)
            auth_mode = "delegated" if headers else None
    else:
        headers = authenticate_with_msal(silent_only=False)
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
    use_cached_week = st.checkbox("Use cached weekly snapshot if available", True)
    auto_refresh_stale = st.checkbox("Auto-refresh stale weekly snapshot", True)
    fetch = st.button("Fetch and Calculate")

if fetch:
    with st.spinner("Fetching events and calculating..."):
        iso_year, iso_week = get_week_key(end_date)
        cache_used = False
        summary_df = pd.DataFrame()
        updated_events_df = pd.DataFrame()

        if use_cached_week:
            stale = is_snapshot_stale(end_date)
            if stale and auto_refresh_stale:
                st.info(f"Cached snapshot for {iso_year}-W{iso_week} is stale; refreshing.")
            else:
                cached_summary, cached_events = load_range_snapshot(start_date, end_date)
                if cached_events is not None and not cached_events.empty:
                    summary_df, updated_events_df = summarize_vacation(cached_events, start_date, end_date)
                    cache_used = True
                    st.success("Loaded cached events for selected date range.")
                elif cached_summary is not None:
                    cache_used = True
                    summary_df = cached_summary.drop(columns=[c for c in cached_summary.columns if c.startswith("Remaining ")], errors="ignore")
                    if cached_events is not None:
                        updated_events_df = cached_events
                    st.success("Loaded cached snapshot for selected date range.")

        if not cache_used:
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
            save_week_snapshot(start_date, end_date, summary_df, updated_events_df)
            st.caption(f"Snapshot saved for {iso_year}-W{iso_week}.")

        st.success("Done!")
        st.subheader("Vacation Summary")
        st.dataframe(summary_df)  # ✅ Only display the summary

        st.subheader("Testing Zone")
        calendar_month = st.date_input("Calendar Month", end_date, key="calendar_month")
        calendar_html = build_vacation_calendar_html(updated_events_df, calendar_month)
        if not calendar_html:
            st.info("No vacations to display in the calendar for the selected month.")
        else:
            st.markdown(
                """
                <style>
                .vacation-calendar { font-family: "Trebuchet MS", "Segoe UI", sans-serif; }
                .vacation-calendar .calendar-header { font-size: 1.2rem; font-weight: 700; margin-bottom: 0.5rem; }
                .vacation-calendar table { width: 100%; border-collapse: collapse; table-layout: fixed; }
                .vacation-calendar th { text-align: left; font-size: 0.85rem; padding: 0.4rem; color: #333; border-bottom: 1px solid #ddd; }
                .vacation-calendar td { vertical-align: top; padding: 0.4rem; border: 1px solid #eee; height: 90px; background: #fafafa; }
                .vacation-calendar td.has-vac { background: #fff4d6; border-color: #f2d08b; }
                .vacation-calendar td.muted { background: #f5f5f5; color: #999; }
                .vacation-calendar .date { font-weight: 700; margin-bottom: 0.25rem; }
                .vacation-calendar .names { font-size: 0.78rem; line-height: 1.2; color: #333; }
                </style>
                """,
                unsafe_allow_html=True,
            )
            st.markdown(calendar_html, unsafe_allow_html=True)

        display_cols = [c for c in ["Date", "Name", "Weekday", "Subject", "Used Status", "Carryover Window"] if c in updated_events_df.columns]
        if display_cols:
            st.caption("Selected vacations in the current date range")
            st.dataframe(updated_events_df.sort_values(["Date", "Name"])[display_cols], use_container_width=True)

        # ✅ Export both to Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            updated_events_df.to_excel(writer, sheet_name="Events", index=False)

        st.download_button("Download Excel Summary", buffer.getvalue(), file_name="vacation_summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
