import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from azure.identity import DeviceCodeCredential
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


hr_holidays = holidays.country_holidays("HR")
def is_working_day(date):
    return date.weekday() < 5 and date not in hr_holidays

def fetch_calendar_events(headers, start_datetime, end_datetime, keyword):
    start = datetime.combine(start_datetime, datetime.min.time()).astimezone(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
    end = datetime.combine(end_datetime, datetime.max.time()).astimezone(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')

    url = f"https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime={start}&endDateTime={end}"
    all_events = []

    while url:
        response = requests.get(url, headers=headers)
        result = response.json()
        events = result.get("value", [])
        all_events.extend(events)
        url = result.get("@odata.nextLink", None)

    vacation_rows = []

    for ev in all_events:
        subject = ev.get("subject", "")
        if (
            re.search(r'\bg[\.\s]?o\b', subject, re.IGNORECASE) and
            not re.search(r'\bcanceled\b|\botkazano\b', subject, re.IGNORECASE)
        ):
            organizer = ev.get("organizer", {}).get("emailAddress", {}).get("name", "Unknown")
            start = pd.to_datetime(ev.get("start", {}).get("dateTime")).normalize()
            end = pd.to_datetime(ev.get("end", {}).get("dateTime")).normalize()
            all_dates = pd.date_range(start=start, end=end - timedelta(days=1))

            for d in all_dates:
                date_only = d.date()
                if is_working_day(date_only) and (start_datetime <= date_only <= end_datetime):
                    vacation_rows.append({
                        "Name": organizer,
                        "Date": date_only,
                        "Weekday": date_only.strftime("%A")
                    })
    return pd.DataFrame(vacation_rows)

# def get_calendar_events(graph_client, start_datetime, end_datetime, keyword):
#     return asyncio.get_event_loop().run_until_complete(
#         fetch_calendar_events(graph_client, start_datetime, end_datetime, keyword)
#     )
def get_calendar_events(graph_client, start_datetime, end_datetime, keyword):
    return fetch_calendar_events(graph_client, start_datetime, end_datetime, keyword)


def summarize_vacation(events_df, start_date, end_date):
    events_df = events_df[(events_df["Date"] >= start_date) & (events_df["Date"] <= end_date)]

    try:
        allowance_df = pd.read_csv("vacation_allowances.csv")
    except FileNotFoundError:
        st.warning("'vacation_allowances.csv' not found. Defaulting to 25 days for everyone.")
        allowance_df = pd.DataFrame(columns=["Name", "Allowance"])

    enriched_rows = []
    summary_rows = []

    for name, group in events_df.groupby("Name"):
        allowance_value = allowance_df.loc[allowance_df["Name"] == name, "Allowance"].values
        allowance = int(allowance_value[0]) if len(allowance_value) > 0 else 25

        group = group.sort_values("Date").copy()
        group["Allowance Year"] = None
        group["Used Status"] = None

        # Build allowance pool: {2023: {"remaining": 25, "valid_until": 2024-06-30}, ...}
        allowance_pool = {}
        for year in range(group["Date"].min().year - 1, group["Date"].max().year + 1):
            allowance_pool[year] = {
                "remaining": allowance,
                "valid_until": date(year + 1, 6, 30)
            }

        over_limit = 0
        usage_by_year = {}

        for idx, row in group.iterrows():
            used = False
            for y in sorted(allowance_pool.keys()):
                if row["Date"] <= allowance_pool[y]["valid_until"] and allowance_pool[y]["remaining"] > 0:
                    allowance_pool[y]["remaining"] -= 1
                    group.at[idx, "Allowance Year"] = y
                    group.at[idx, "Used Status"] = "Within Allowance"
                    usage_by_year[y] = usage_by_year.get(y, 0) + 1
                    used = True
                    break
            if not used:
                group.at[idx, "Used Status"] = "Over Allowance"
                over_limit += 1

        # Summary per person
        summary = {
            "Name": name,
            "Used Total": len(group),
            "Over Limit": over_limit,
            "⚠️ Over Limit?": "Yes" if over_limit > 0 else "No"
        }

        # Add per-year used and remaining
        for y in sorted(allowance_pool.keys()):
            summary[f"Used {y}"] = usage_by_year.get(y, 0)
            summary[f"Remaining {y}"] = allowance_pool[y]["remaining"]

        summary_rows.append(summary)
        enriched_rows.append(group)

    summary_df = pd.DataFrame(summary_rows)
    updated_events_df = pd.concat(enriched_rows, ignore_index=True)

    return summary_df, updated_events_df


# --- Streamlit UI ---
st.title("Vacation Tracker - GO Events Summary")

#graph_client = authenticate_graph()
# headers = authenticate_graph()
headers = authenticate_device_flow()

with st.sidebar:
    st.header("Filter Settings")
    start_date = st.date_input("Start Date", datetime.now() - timedelta(days=30))
    end_date = st.date_input("End Date", datetime.now() + timedelta(days=30))
    keyword = st.text_input("Keyword to Filter Events", "GO")
    fetch = st.button("Fetch and Calculate")

if fetch:
    with st.spinner("Fetching events and calculating..."):
        events_df = get_calendar_events(headers, start_date, end_date, keyword)

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
