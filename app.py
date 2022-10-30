import datetime as dt
import pandas as pd
import os
import smtplib
from dotenv import load_dotenv


def fetch_planner():
    TIME_PERIOD = int(os.environ["TIME_PERIOD"])
    ORG_PLANNER_PATH = os.environ["ORG_PLANNER_PATH"]
    ORG_PLANNER_SHEET_NAME = os.environ["ORG_PLANNER_SHEET_NAME"]
    ORG_PLANNER_HEADER_ROW = int(os.environ["ORG_PLANNER_HEADER_ROW"])
    ORG_PLANNER_COLUMNS = os.environ["ORG_PLANNER_COLUMNS"]


    planner = pd.read_excel(io=ORG_PLANNER_PATH, sheet_name=ORG_PLANNER_SHEET_NAME, header=ORG_PLANNER_HEADER_ROW, usecols=ORG_PLANNER_COLUMNS)

    planner = planner.loc[0:43]
    planner = planner.drop([0,1])
    planner.columns.values[0] = "Namn"
    planner.dropna(subset="Namn", inplace=True)
    planner.set_index("Namn", inplace=True)

    today = dt.datetime.today()
    start_date = today - dt.timedelta(days=today.weekday())
    
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)

    period_dates = []
    for i in range(TIME_PERIOD):
        date = start_date + dt.timedelta(days=i)
        period_dates.append(date)

    planner = planner.loc[:,period_dates]

    return planner


def mail_daily_digest(planner):
    EMAIL_SENDER = os.environ["EMAIL_SENDER"]
    EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
    EMAIL_RECIPIENT = os.environ["EMAIL_RECIPIANT"]
    SUBJECT = "Subject:Daily Digest\n\n"
    SIG = "Ha en bra dag\nAdrian"
    ME = os.environ["ME"] 
    
    
    msg = ""
    planner = planner.loc[ME]
    print(planner)


def main():
    load_dotenv()
    planner = fetch_planner()
    mail_daily_digest(planner)
        
if __name__ == "__main__":
    main()
