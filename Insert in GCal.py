# %%
import datetime as dt
import pickle


import pandas as pd
from googleapiclient.discovery import build
from pandas._libs.tslibs import Timestamp


# from google_auth_oauthlib.flow import InstalledAppFlow

PlanXlPn = 'week.xlsm'
SheetName = 'gcal'
CalName = "plan"

FirstIdenCNo = 4 - 1
DayColsLen = 7
EventNameCNo = 7
EventNoCNo = 7 - 1
StartRow = 1
RowsADay = 24 * 4
TimeCNo = 3 - 1

Cols = ["Name", "Start Date", "End Date", "Start Time", "End Time"]


def find_next_event_row(day, row, data_0):
    search_col = FirstIdenCNo + day * DayColsLen
    for i in range(row + 1, RowsADay + 1):
        if data_0[search_col][i] != 0:
            return i
    return find_next_event_row(day + 1, StartRow - 1, data_0)


def main():
    pass

    # %%

    data_0 = pd.read_excel(PlanXlPn, sheet_name=SheetName, engine="openpyxl", header=None)
    data_0

    # %%
    plan = pd.DataFrame(columns=Cols)

    for weekday in range(0, 7):
        event_start_row = StartRow - 1

        while True:
            event_start_row = find_next_event_row(weekday, event_start_row, data_0)
            event_start_time = data_0[TimeCNo][event_start_row]
            event_end_row = find_next_event_row(weekday, event_start_row, data_0)
            event_end_time = data_0[TimeCNo][event_end_row]

            if event_end_row < event_start_row:
                event_name = data_0[EventNameCNo + TimeCNo + weekday * DayColsLen][event_end_row]
                new_event = {Cols[0]: event_name,
                             Cols[1]: weekday,
                             Cols[2]: weekday + 1,
                             Cols[3]: event_start_time,
                             Cols[4]: event_end_time}
                plan = plan.append(new_event, ignore_index=True)
                break

            else:
                event_name = data_0[EventNameCNo + TimeCNo + weekday * DayColsLen][event_end_row]
                new_event = {Cols[0]: event_name,
                             Cols[1]: weekday,
                             Cols[2]: weekday,
                             Cols[3]: event_start_time,
                             Cols[4]: event_end_time}
                plan = plan.append(new_event, ignore_index=True)

    plan = plan.reset_index(drop=True)
    plan

    # %%
    plan = plan[plan['Name'].ne('Null')]

    # %%
    plan = plan.applymap(lambda x: x.time() if type(x) == Timestamp else x)
    plan

    # %%
    for weekday in range(0, 8):
        for start_or_end in Cols[1:3]:
            if weekday < dt.date.weekday(dt.date.today()):
                delta_days = 7 - (dt.date.weekday(dt.date.today()) - weekday)
            else:
                delta_days = weekday - dt.date.weekday(dt.date.today())

            plan.loc[plan[start_or_end] == weekday, [start_or_end]] = dt.date.today() + dt.timedelta(days=delta_days)
    plan

    # %%
    # For sunday (end of week) last event
    plan.loc[plan[Cols[2]] < plan[Cols[1]], Cols[2]] += dt.timedelta(days=7)
    plan

    # %%
    ## OAuth 2.0 Setup
    # scopes = ['https://www.googleapis.com/auth/calendar']
    # client_secret_file_name = "client_secret_mac1.json" #client secret file in working directory
    # flow = InstalledAppFlow.from_client_secrets_file(client_secret_file_name, scopes=scopes)
    # credentials = flow.run_console()
    # pickle.dump(credentials, open("token.pkl", "wb"))

    # %%
    credentials = pickle.load(open("token.pkl", "rb"))
    service = build("calendar", "v3", credentials=credentials)

    # %%
    all_calendars = service.calendarList().list().execute()
    interested_cal_data = next(item for item in all_calendars["items"] if item["summary"] == CalName)

    # %%
    today = dt.date.today()

    t = dt.datetime(dt.datetime.today().year, dt.datetime.today().month, dt.datetime.today().day, 0, 0, 0)
    time_min = t.astimezone().isoformat()

    # %%
    seven_days_ahead_ids = []
    seven_days_data = service.events().list(timeMin=time_min, calendarId=interested_cal_data["id"]).execute()
    seven_days_ahead_ids.extend([item["id"] for item in seven_days_data["items"]])

    while True:
        if not ("nextPageToken" in seven_days_data):
            break
        seven_days_data = service.events().list(pageToken=seven_days_data["nextPageToken"], timeMin=time_min,
                                                calendarId=interested_cal_data["id"], ).execute()
        seven_days_ahead_ids.extend([item["id"] for item in seven_days_data["items"]])

    # %%
    for event_id in seven_days_ahead_ids:
        service.events().delete(calendarId=interested_cal_data["id"], eventId=event_id).execute()

    print('Done Deleting!')

    # %%
    timezone = "Asia/Tehran"

    for row_num in range(len(plan)):
        if plan.iloc[row_num][Cols[2]] >= today:
            event_name = plan.iloc[row_num][Cols[0]]
            start_time = dt.datetime.combine(plan.iloc[row_num][Cols[1]], plan.iloc[row_num][Cols[3]])
            end_time = dt.datetime.combine(plan.iloc[row_num][Cols[2]], plan.iloc[row_num][Cols[4]])
            new_event = {
                "summary"    : event_name,
                "location"   : "",
                "description": "",
                "start"      : {
                    "dateTime": start_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": timezone,
                },
                "end"        : {
                    "dateTime": end_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "timeZone": timezone,
                },
                "reminders"  : {
                    "useDefault": True,
                },
            }
            service.events().insert(calendarId=interested_cal_data["id"], body=new_event).execute()

    print("Done!")


# %%


if __name__ == '__main__':
    main()


# %%
# pip freeze > reqs.txt
