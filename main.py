# import asyncio
# import configparser
# from msgraph.generated.models.o_data_errors.o_data_error import ODataError
# from graph import Graph
# from msgraph import GraphServiceClient
# #from msgraph.generated.graph_client import GraphClient
# from azure.identity import DeviceCodeCredential
# from msgraph.graph_service_client import GraphServiceClient
# #from msgraph.generated.me.calendar.events.events_request_builder import EventsRequestBuilderGetRequestConfiguration
# from msgraph.generated.models.o_data_errors.o_data_error import ODataError
# from msgraph import GraphServiceClient
# from msgraph.generated.users.item.calendar.events.events_request_builder import EventsRequestBuilder
# from kiota_abstractions.base_request_configuration import RequestConfiguration
# import re
# import pandas as pd
# import openpyxl
# from openpyxl.utils import get_column_letter
# from openpyxl.styles import Font
# import os
# from openpyxl.styles import NamedStyle
# from datetime import datetime
# # Removed unused import that could not be resolved
# #from msgraph.core import GraphClient
# from datetime import datetime, timedelta, timezone
# # from msgraph.generated.me.calendar.calendar_view.calendar_view_request_builder import (
# #     CalendarViewRequestBuilderGetQueryParameters,
# #     CalendarViewRequestBuilderGetRequestConfiguration
# # )

# #result = await graph_client.me.calendar.events.get()
# async def main():
#     scopes = ['User.Read', 'Calendars.Read'] 

# # Multi-tenant apps can use "common",
# # single-tenant apps must use the tenant ID from the Azure portal

# # Values from app registration
#    '

# # azure.identity
#     credential = DeviceCodeCredential(
#         tenant_id=tenant_id,
#         client_id=client_id)
#     graph_client = GraphServiceClient(credential, scopes)


#     start_datetime = (datetime.now(timezone.utc) - timedelta(days=365)).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
#     end_datetime = (datetime.now(timezone.utc) + timedelta(days=365)).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
#     print(f"Querying from {start_datetime} to {end_datetime}")
#     url = f"https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime={start_datetime}&endDateTime={end_datetime}"

    
#     all_events = []
#     #page = graph_client.me.calendar.calendar_view.with_url(url)
#     calendar_view_builder = graph_client.me.calendar.calendar_view.with_url(url)
#     page = await calendar_view_builder.get()

#     while page:
#         all_events.extend(page.value)

#         # Get next page if available
#         if page.odata_next_link:
#             page = await graph_client.me.calendar.calendar_view.with_url(page.odata_next_link).get()
#         else:
#             break

#     excel_file = "GO.xlsx"
#     sheet_name = "Vacations"


#     filtered_events = []

#     print(f"Filtered events with 'GO': {len(filtered_events)}")

#     for event in all_events:
#         if re.search(r'\bGO\b', event.subject, re.IGNORECASE):
#             print(f"{event.subject} at {event.start.date_time} {event.end.date_time}")
#             filtered_events.append({
#                 'subject': event.subject,
#                 'start': event.start.date_time,
#                 'end': event.end.date_time
#             })
#     if os.path.exists(excel_file): 
#         wb = openpyxl.load_workbook(excel_file)
#         ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
#     else:
#         wb = openpyxl.Workbook()
#         ws = wb.active
#         ws.title = sheet_name
#         # Write header if creating new file
#         ws.append(["Subject", "Start", "End"])
#         # Optional: bold headers
#         for col in range(1, 4):
#             ws.cell(row=1, column=col).font = Font(bold=True)


#     for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=3):
#         for cell in row:
#             cell.value = None
#     # Find next empty row
#     next_row = 2
#     # Append new events
#     for i, event_data in enumerate(filtered_events, start=next_row):
#         start_time = datetime.fromisoformat(event_data['start'])
#         end_time = datetime.fromisoformat(event_data['end'])
    
#         ws.cell(row=i, column=1, value=event_data['subject'])
#         ws.cell(row=i, column=2, value=start_time)
#         ws.cell(row=i, column=3, value=end_time)

#     # Save
#     wb.save(excel_file)
#     print(f"Appended {len(filtered_events)} events to {excel_file}")


# asyncio.run(main())


