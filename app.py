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

    today = dt.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    period_dates = []
    for i in range(TIME_PERIOD):
        date = today + dt.timedelta(days=i)
        period_dates.append(date)

    planner = planner.loc[:,period_dates]
    return planner


def construct_message(planner):
    ME = os.environ["ME"]


    my_planner = planner.loc[ME]
    msg = []
    jobs = {}
    for index, value in my_planner.items():
        date = str(index).split()[0]
        jobs[date] = {}
        if pd.isna(value):
            jobs[date] = {"job" : "Oplanerad", "team" : []}
        else:
            team = []
            for name, job in planner[index].items():
                if job == value and name != ME:
                    team.append(name)
            jobs[date] = {"job" : value, "team" : team}

    for date in jobs.items():
        local_msg = f"{date[0]} är du planerad på; {date[1]['job']}.\nTilsammans med:\n"
        for name in date[1]["team"]:
            msg = msg + f"{name}\n"
        msg.append(local_msg)

    return msg


def mail_daily_digest(message):
    EMAIL_SENDER = os.environ["EMAIL_SENDER"]
    EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
    EMAIL_RECIPIENT = os.environ["EMAIL_RECIPIANT"]
    SUBJECT = "Subject:Daily Digest\n\n"
    SIG = "Ha en bra dag\nAdrian"
    HOST = os.environ["HOST"]
    PORT= os.environ["PORT"]

    mail_msg = SUBJECT
    for index in message:
        mail_msg = mail_msg + index
    mail_msg = mail_msg + SIG
    mail_msg.encode("UTF-8")

    with smtplig.SMTP(host=HOST, port=PORT) as connection:
        connection.starttls()
        connection.login(user=EMAIL_SENDER, to_addrs=EMAIL_RECIPIENT, msg=mail_msg)
    
    return

    
def main():
    load_dotenv()
    planner = fetch_planner()
    message = construct_message(planner)
    mail_daily_digest(message)
        
if __name__ == "__main__":
    main()
