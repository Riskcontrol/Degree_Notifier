import os
import pandas as pd
import requests
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 1. download the workbook ----------------------------------------------------
EXCEL_URL = (
    "https://riskcontrolacademy-my.sharepoint.com/:x:/g/personal/"
    "officeadmin_riskcontrolnigeria_com/"
    "EZQwO5Z3ShpNpN4TyjGXj-8BpPruKpHn6ZMolDPyUh5-4w?download=1"
)

local_file = "tat_data.xlsx"
resp = requests.get(EXCEL_URL, timeout=30)
resp.raise_for_status()
with open(local_file, "wb") as f:
    f.write(resp.content)
print("Excel downloaded")

# 2. load and enrich ----------------------------------------------------------
df = pd.read_excel(local_file, engine="openpyxl")

# normalise column names just once
df.columns = df.columns.str.strip().str.upper()

if "DATE RECEIVED" not in df.columns:
    raise KeyError("DATE RECEIVED column not found")

# parse and build TAT due date
df["DATE RECEIVED"] = pd.to_datetime(df["DATE RECEIVED"], dayfirst=True, errors="coerce")
df["TAT DUE"] = df["DATE RECEIVED"] + pd.DateOffset(months=3)

today = datetime.now()

# 3. mail setup ---------------------------------------------------------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
SENDER = os.environ.get("USER_EMAIL")   # same pattern you used earlier
PASSWORD = os.environ.get("USER_PASSWORD")
TARGET = "officeadmin@riskcontrolnigeria.com"

if not SENDER or not PASSWORD:
    raise EnvironmentError("USER_EMAIL or USER_PASSWORD not set")

def send_reminder(row, days_left):
    name = row["NAMES"]
    client = row.get("CLIENTS", "Unspecified")
    received = row["DATE RECEIVED"].strftime("%d %b %Y")
    due = row["TAT DUE"].strftime("%d %b %Y")

    subject = f"Turnaround time alert – {days_left} day(s) left for {name}"
    body = f"""
    Dear Admin,

    Verification for {name} (client: {client}) was logged on {received}.
    The three-month turnaround deadline is {due}, which is in {days_left} day(s).

    Kindly ensure completion and documentation before the deadline.

    Regards
    Nigeria Risk Index Automation Bot
    """

    msg = MIMEMultipart()
    msg["From"] = SENDER
    msg["To"] = TARGET
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SENDER, PASSWORD)
        server.sendmail(SENDER, TARGET, msg.as_string())
    print(f"Mail sent for {name} ({days_left} days left)")

# 4. iterate and alert --------------------------------------------------------
for _, r in df.iterrows():
    if pd.isna(r["TAT DUE"]):
        continue
    days_remaining = (r["TAT DUE"] - today).days
    if days_remaining in {30, 14, 1}:
        send_reminder(r, days_remaining)

print("TAT check complete")
