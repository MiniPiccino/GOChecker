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

#nest_asyncio.apply()


TENANT_ID = "94aa9436-2653-434c-bd47-1124432cb7d7"
CLIENT_ID = "b9d8662f-5800-4498-a4f1-28b92bea4f39"
SCOPE = "https://graph.microsoft.com/.default"

DEVICE_CODE_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/devicecode"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
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

    st.info("🔐 Microsoft Login Required")
    st.markdown(f"👉 [Click here to log in]({data['verification_uri']})", unsafe_allow_html=True)
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
                st.success("✅ Authentication successful.")
                return {
                    "Authorization": f"Bearer {access_token}",
                    "Accept": "application/json"
                }

            elif token_response.status_code == 400:
                error = token_response.json().get("error")
                if error in ["authorization_pending", "slow_down"]:
                    continue
                else:
                    st.error(f"❌ Authentication failed: {error}")
                    return None

        st.error("⏰ Authentication timed out.")
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
    print(f"Filtered events with '{keyword}': {(vacation_rows)}")
    return pd.DataFrame(vacation_rows)

# def get_calendar_events(graph_client, start_datetime, end_datetime, keyword):
#     return asyncio.get_event_loop().run_until_complete(
#         fetch_calendar_events(graph_client, start_datetime, end_datetime, keyword)
#     )
def get_calendar_events(graph_client, start_datetime, end_datetime, keyword):
    return fetch_calendar_events(graph_client, start_datetime, end_datetime, keyword)
def summarize_vacation(events_df):
    # Load allowances from CSV
    try:
        allowance_df = pd.read_csv("vacation_allowances.csv")
    except FileNotFoundError:
        st.warning("⚠️ 'vacation_allowance.csv' not found. Defaulting to 25 days for everyone.")
        allowance_df = pd.DataFrame(columns=["Name", "Allowance"])

    # Count used vacation days
    used = events_df.groupby("Name")["Date"].count().reset_index().rename(columns={"Date": "Used"})

    # Merge with allowance
    merged = pd.merge(used, allowance_df, on="Name", how="left")
    merged["Allowance"] = merged["Allowance"].fillna(25).astype(int)
    merged["Remaining"] = merged["Allowance"] - merged["Used"]
    return merged

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
        summary_df = summarize_vacation(events_df)

        st.success("Done!")
        st.subheader("Vacation Summary")
        st.dataframe(summary_df)

        # Excel export
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            events_df.to_excel(writer, sheet_name="Events", index=False)
        st.download_button("Download Excel Summary", buffer.getvalue(), file_name="vacation_summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
