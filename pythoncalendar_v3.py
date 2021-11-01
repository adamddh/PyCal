'''a functionified version of v2
By Adam DenHaan
Jun 15, 2020

calendar_colors = {
    "Lavender"  : 1
    "Sage"      : 2
    "Grape"     : 3
    "Flamingo"  : 4
    "Banana"    : 5
    "Tangerine" : 6
    "Peacock"   : 7
    "Grapite"   : 8
    "Blueberry" : 9
    "Basil"     : 10
    "Tomato"    : 11
}

'''


from datetime import datetime
from os.path import exists
from pickle import dump, load
from sys import argv
from threading import Thread
from time import sleep
from typing import Union

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import Resource, build
from pygsheets import authorize
from pygsheets.spreadsheet import Spreadsheet
from pygsheets.worksheet import Worksheet
from requests import Timeout, get
from termcolor import colored

ARG = argv[1] if len(argv) > 1 else None


def hasnumbers(inputstring: str) -> bool:
    """Check if string has numbers"""
    return any(char.isdigit() for char in inputstring)


def calhelp(
    initials, calendar_id, directory, sheet_secret_name,
    cal_secret_name, param, color_id=11
):  # sourcery no-metrics
    """designed to facilitate the use of v2

    Args:
        initials (string): intials to search for on the sheet
        calendar_id (string): calendar ID of the calendar to write events to
        directory (string): location of where to store credential files
        sheet_secret_name (string): filename of the sheet secret name
        cal_secret_name (string): filename of the calendar secret name
        color_id (int, optional): color of the events on the calendar. Defaults to 11.
    """

    now: datetime = datetime.now()

    # makes available the google sheet
    sheets = sh_creds(directory, sheet_secret_name)

    # makes available the google calendar
    service = gc_creds(directory, cal_secret_name)

    t_del = Thread(target=del_events, args=(now, service, calendar_id))

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Deleting events...", "white", attrs=['bold']))

    t_del.start()

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Fetching events...", "white", attrs=['bold']))

    # use only the event sheet within the workbook
    for sheet in sheets:
        if sheet.title == 'Tech Schedule':
            events_sheet: Worksheet = sheet

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Finding my events...", "white", attrs=['bold']))

    my_events_rows = get_event_rows(events_sheet, initials)

    dates = events_sheet.get_col(1)
    titles = events_sheet.get_col(2)
    calls = events_sheet.get_col(3)
    starts = events_sheet.get_col(4)
    ends = events_sheet.get_col(5)
    locations = events_sheet.get_col(6)
    workers = events_sheet.get_col(7)
    records = events_sheet.get_col(8)
    event_coords = events_sheet.get_col(11)

    t_del.join()

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Adding events to calendar...", "white", attrs=['bold']))

    for i in my_events_rows:

        date_obj = datetime.strptime(dates[i], "%A-%b-%d-%y")
        # year/month/date ints
        month = date_obj.month
        year = date_obj.year
        date = date_obj.day

        # event title string
        name = titles[i]

        # call hour and minute ints
        start_hour, start_minute = get_event_time(calls[i])

        # end hour and minute ints
        end_hour, end_minute = get_event_time(ends[i])
        end_string = ends[i]

        # location string
        location = locations[i]

        # record string
        record = "Yes" if records[i] == "Yes" else "No"
        # start string
        start = starts[i]

        # Event coordinator string
        event_coord = event_coords[i]

        # Description string
        descripion = ('Automatic creation\nEvent Start Time: ' + start +
                      '\nEvent Coordinator: ' + event_coord +
                      '\nRecord: ' + record +
                      '\nWorkers: ' + workers[i] +
                      '\nRuntime: ' + str(datetime.now()))

        # create the calendar event
        start_time = datetime(year, month, date, start_hour, start_minute, 0)
        end_time = datetime(year, month, date, end_hour, end_minute, 0)

        if start_time > end_time:
            start_hour, start_minute = get_event_time(starts[i])

        start_time = datetime(year, month, date, start_hour, start_minute, 0)
        calevent = {
            'summary': name,
            'location': location,
            'colorId': color_id,          # where you can select the color of the event
            'description': descripion,
            'start': {
                'dateTime': start_time.strftime("%Y-%m-%dT%H:%M:%S"),
                'timeZone': "America/New_York",
            },
            'end': {
                'dateTime': end_time.strftime("%Y-%m-%dT%H:%M:%S"),
                'timeZone': "America/New_York",
            },
            'reminders': {
                'useDefault': True,
            },
        }
        # adds calendar event
        add_cal_event(end_string, now, start_time, param,
                      name, service, calendar_id, calevent)


def get_event_time(time_string: str) -> Union[int, int]:
    """Get event time info"""
    if hasnumbers(time_string):
        hour = int(time_string.split(':')[0])
        minute = int(time_string.split(':')[1][0:2])
        if time_string.split(':')[1][-2:] == "PM" and hour != 12:
            hour += 12
    else:
        hour = 23
        minute = 59
    return hour, minute


def add_cal_event(
    end_string, now, end_time, param, name, service, calendar_id, calevent
) -> None:
    """Add event to calendar"""
    if hasnumbers(end_string) and now < end_time:
        if param == "v":
            print('  ', name)
        try:
            add_event(service, calendar_id, calevent)
        except Exception:
            sleep(1)
            add_event(service, calendar_id, calevent)


def add_event(service, calendar_id, calevent) -> None:
    """Add event to service"""
    service.events().insert(calendarId=calendar_id,
                            body=calevent).execute()


def get_event_rows(events_sheet, initials) -> list:
    """Get events that {initals} are working"""
    my_events = events_sheet.find(initials)
    my_events_rows = [event.row - 1 for event in my_events]
    everyone_events = events_sheet.find(
        'ALL', cols=(1, 7), matchEntireCell=True)
    for event in everyone_events:
        my_events_rows.append(event.row - 1)
    return sorted(my_events_rows)


def del_events(now, service, calendar_id) -> None:
    """Delete future events"""
    time_min = now.isoformat('T') + "-05:00"
    events = service.events()

    while True:
        delete_events_id = []
        result = events.list(calendarId=calendar_id,
                             timeMin=time_min, singleEvents=True,
                             orderBy="startTime").execute()
        for i in range(len(result['items'])):
            try:
                if result['items'][i]['description']\
                        .startswith('Automatic creation') and \
                        datetime.strptime(
                            result["items"][i]["start"]["dateTime"][:-6],
                            "%Y-%m-%dT%H:%M:%S"
                ) > now:
                    delete_events_id.append(result['items'][i]['id'])
            except KeyError:  # not all event have a feild 'description', and raises
                pass  # an error if there is no feild. Addressed by skipping
                # event as its not ours
        for i in delete_events_id:
            try:
                events.delete(calendarId=calendar_id,
                              eventId=i).execute()
            except Exception:
                sleep(1)
                events.delete(calendarId=calendar_id,
                              eventId=i).execute()
        if not delete_events_id:
            break


def gc_creds(directory, cal_secret_name) -> Resource:
    """Authorize access to google calendar"""
    pklstr = directory + '/token.pkl'
    scopes = ['https://www.googleapis.com/auth/calendar']
    if exists(pklstr):
        with open(pklstr, "rb") as token:
            credentials = load(token)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            directory + "/" + cal_secret_name,
            scopes=scopes)
        credentials = flow.run_local_server(port=0)
        with open(pklstr, "wb") as token:
            dump(credentials, token)
    return build("calendar", "v3", credentials=credentials)


def sh_creds(directory, sheet_secret_name) -> Spreadsheet:
    """Get google sheet"""
    gcal = authorize(
        client_secret=directory + "/" + sheet_secret_name,
        credentials_directory=directory,
        local=True)
    return gcal.open_by_key("1cEZ6P6EXSaBuu00gEbg2mmqbAci1CECGKorfmdMGtfQ")


def check_connection(param: str) -> bool:
    """
    Check if computer has internet connection

    Returns:
        bool: True if internet connection.
    """
    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Checking connection...", "white", attrs=['bold']))
    url = "https://www.google.com/"
    timeout = 15
    try:
        get(url, timeout=timeout)
        if param == "v":
            print('\t', "Connection established")
        return True
    except (ConnectionError, Timeout):
        if param == "v":
            print('\t', "Connection failed, exiting.")
        return False


def run_adam():
    """run_adam"""
    calhelp(
        initials="ADH",
        calendar_id="adamdh00@gmail.com",
        directory="/Users/adamdenhaan/Documents/PyCal/credentials",
        sheet_secret_name="SCS.json",
        cal_secret_name="PCS.json",
        color_id=11,
        param=ARG)


def run_bri():
    """run_bri"""
    calhelp(
        initials="BJ",
        calendar_id="brijans19@gmail.com",
        directory="/Users/adamdenhaan/Documents/PyCal/credentials/bricreds",
        sheet_secret_name="SCS.json",
        cal_secret_name="PCS.json",
        color_id=7,
        param=None
    )


def main():
    """Main()"""
    Thread(target=run_adam).start()
    Thread(target=run_bri).start()


if __name__ == '__main__' and check_connection(ARG):
    main()
