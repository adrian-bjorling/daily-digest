import datetime as dt
import pandas as pd
import os
import smtplib
from json import load
from dotenv import load_dotenv


def fetch_planner():
    TIME_PERIOD = int(os.environ["TIME_PERIOD"])
    ORG_PLANNER_PATH = os.environ["ORG_PLANNER_PATH"]
    ORG_PLANNER_SHEET_NAME = os.environ["ORG_PLANNER_SHEET_NAME"]
    ORG_PLANNER_HEADER_ROW = int(os.environ["ORG_PLANNER_HEADER_ROW"])
    ORG_PLANNER_COLUMNS = os.environ["ORG_PLANNER_COLUMNS"]


    planner = pd.read_excel(io=ORG_PLANNER_PATH, sheet_name=ORG_PLANNER_SHEET_NAME, header=ORG_PLANNER_HEADER_ROW, usecols=ORG_PLANNER_COLUMNS)

    planner = planner.loc[0:78]
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


def construct_message(planner, name):
    ME = name

    my_planner = planner.loc[ME]
    msg = []
    jobs = {}
    for index, value in my_planner.items():
        if index.weekday() > 4:
            date = str(index).split()[0] + "(Helg)"
        else:
            date = str(index).split()[0]
        jobs[date] = {}
        if pd.isna(value):
            jobs[date] = {"job" : "-", "team" : []}
        else:
            team = []
            for name, job in planner[index].items():
                if job == value and name != ME:
                    team.append(name)
            jobs[date] = {"job" : value, "team" : team}

    for date in jobs.items():
        local_msg = f"{date[0]}: {date[1]['job']}\n"
        if date[1]["job"] != "Oplanerad" and date[1]["team"] != []:
            for name in date[1]["team"]:
                local_msg = local_msg + f"->{name}\n"
        msg.append(local_msg)

    return msg


def write_msg_to_file(msg):
    with open("message.txt", "w") as doc:
        for index in msg:
            doc.write(index)
    return

def mail_daily_digest(message, email_to):
    EMAIL_FROM = os.environ["EMAIL_FROM"]
    EMAIL_TO = email_to
    EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
    SUBJECT = "Subject:Daily Digest\n\n"
    SIG = "\nHa en bra dag!"
    HOST = "smtp.gmail.com"
    PORT= 587

    msg = ""
    for index in message:
        msg += index
    msg = SUBJECT + msg + SIG
    msg = msg.encode("utf-8")

    server = smtplib.SMTP(HOST, PORT)
    server.starttls()
    server.login(EMAIL_FROM, EMAIL_PASSWORD)

    server.sendmail(EMAIL_FROM, EMAIL_TO, msg)
    server.quit()
    return


def main():
    load_dotenv()
    with open(".maillist.json", "r") as data:
        maillist = load(data)

    planner = fetch_planner()
    for name, mail in maillist.items():
        message = construct_message(planner, name)
        mail_daily_digest(message, mail)


if __name__ == "__main__":
    main()
