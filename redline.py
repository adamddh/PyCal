"""File to detect changes in the red line"""

from datetime import datetime
from os.path import exists
from pickle import dump, load
from smtplib import SMTP_SSL
from socket import gaierror
from ssl import create_default_context

from pygsheets import authorize
from pygsheets.spreadsheet import Spreadsheet

from credentials.creds import PASSWORD
from pythoncalendar_v3 import check_connection

DIRECTORY = "/Users/adamdenhaan/Documents/PyCal/credentials/"
SHEET_SECRET_NAME = "SCS.json"


def sh_creds() -> Spreadsheet:
    """Get google sheet"""
    gcal = authorize(
        client_secret=f'{DIRECTORY}bricreds/{SHEET_SECRET_NAME}',
        credentials_directory=DIRECTORY,
        local=True
    )
    return gcal.open_by_key("1f3G7XkZtH4vJJ0qoPjjsOnRHcnWurceP83_cKqI_8jE")


def main() -> None:
    """Main"""
    date = get_redline_date()

    if exists(f"{DIRECTORY}/row.pkl"):
        with open(f"{DIRECTORY}/row.pkl", "rb") as fin:
            old_date = load(fin)
    else:
        with open(f"{DIRECTORY}/row.pkl", "wb") as fout:
            dump(date, fout)

    if date != old_date:
        with open(f"{DIRECTORY}/row.pkl", "wb") as fout:
            dump(date, fout)
        send_email(old_date, date)


def get_redline_date() -> datetime:
    """Get the date of the row of the redline"""
    sheets = sh_creds()

    events_sheet = sheets[0]

    try:
        row1 = events_sheet.range('A1:A2000')
    except gaierror:
        exit()
    for cell in row1:
        cell = cell[0]
        row = cell.row if cell.color == (1, 0, 0, 0) else -1
        if row != -1:
            if cell.value == "":
                found_date = False
                while not found_date:
                    row -= 1
                    cell = events_sheet.cell(f"A{row}")
                    found_date = cell.value != ""
            # return date object of parsed date, eg Saturday-Mar-05-22
            return datetime.strptime(cell.value, "%A-%b-%d-%y")


def send_email(old_date: datetime, new_date: datetime) -> None:
    """Send email to Adam"""
    # format now as HH:MM AM/PM, MM/DD/YYYY
    now = datetime.now().strftime("%I:%M %p, %m/%d/%Y")
    port = 465  # For SSL

    smtp_server = "smtp.gmail.com"
    sender_email = "adamdh00@gmail.com"  # Enter your address
    receiver_email = "add22@students.calvin.edu"  # Enter receiver address
    message = f"""\
Subject: REDLINE MOVED

Detected at {now}. Moved from {old_date} to {new_date}."""

    context = create_default_context()
    with SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, PASSWORD)
        server.sendmail(sender_email, receiver_email, message)


if __name__ == "__main__" and check_connection():
    main()
