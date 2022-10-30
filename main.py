import datetime as dt
import pandas as pd
import os
import smtplib
from dotenv import load_dotenv


def fetch_planner():
    TIME_PERIOD = os.environ["TIME_PERIOD"]
    ORG_PLANNER_PATH = os.environ["ORG_PLANNER_PATH"]
    ORG_PLANNER = os.environ["ORG_PLANNER_SHEET_NAME"]
    ORG_PLANNER_HEADER_ROW = os.environ["ORG_PLANNER_HEADER_ROW"]
    ORG_PLANNER_COLUMNS = os.environ["ORG_PLANNER_COLUMNS"]
    TEMP_TWO_WEEK_PLANNER_PATH = os.environ["TEMP_TWO_WEEK_PLANNER"]


    planner = pd.read_excel(io=ORG_PLANNER_PATH, sheet_name=ORG_PLANNER_SHEET_NAME, header=ORG_PLANNER_HEADER_ROW, usecols=ORG_PLANNER_COLUMNS)

    planner = planner.loc[0:43]
    planner = planner.drop([0,1])
    planner.columns.values[0] = "Namn"
    planner.dropna(subset="Namn", inplace=True)
    planner.set_index("Namn", inplace=True)

    today = dt.datetime.today()
    start_date = today - dt.timedelta(days=today.weekday())

   # start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)

    period_dates = []
    for i in range(TIME_PERIOD):
        date = start_date + dt.timedelta(days=i)
        period_dates.append(date)

    planner = planner.loc[:,period_dates]

    return planner

def two_week_planner(planner):
    ME = os.environ["ME"]

    
    my_planner = planner.loc[ME]

    for index, value in my_planner.items():
        outputstring = str(index).split()[0]
        date = dt.datetime.strptime(index, "%Y-%m-%d %H:%M:%S")
        if date.weekday() > 4:
            outputstring = outputstring + "(Helg): "
        else:
            outputstring = outputstring + ": "
            if pd.isna(value):
                outputstring = outputstring + "-"
            else:
                outputstring = outputstring + value + " ("
                for name, job in planner[str(index)].items():
                    if job == value and name != ME:
                        outputstring = outputstring + name + ", "
                outputstring = outputstring + ")"
    return 

def mail_digest():
    EMAIL = 'raccoon.bjorling@gmail.com'
    PASSWORD = 'Tallskog_123'

    SUBJECT = 'Subject:Mondaymotivation!\n\n'
    MSG = 'Today is Monday. Then it is nice to start the week with some motivation that can carry you trough the week!\n'
    SIG = 'Best regards\nAdrian'

    try:
        with open('./temp/quotes.txt', 'r') as data_file:
            quotes = data_file.readlines()
    except FileNotFoundError:
        quote = '"It is good to be alive!" - Raccoon'
    else:
        quote = random.choice(quotes)
    finally:
        mail_msg = SUBJECT + MSG + quote + SIG
        mail_msg.encode("utf-8")

        with smtplib.SMTP(host='64.233.184.109', port=687) as con:
            con.starttls()
            con.login(user=EMAIL, password=PASSWORD)
            con.sendmail(from_addr=EMAIL, to_addrs='adrian.bjorling@gmail.com', msg=mail_msg)
    return

def main():
    load_dotenv()
    planner = fetch_planner()
    two_week_planner(planner)
        
if __name__ == "__main__":
    main()
