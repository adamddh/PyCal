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
from googleapiclient.errors import HttpError
from pygsheets import authorize
from pygsheets.spreadsheet import Spreadsheet
from pygsheets.worksheet import Worksheet
from requests import Timeout, get
from requests.exceptions import ConnectionError as ConnError
from termcolor import colored

ARG = argv[1] if len(argv) > 1 else None
BASEPATH = "/Users/adamdenhaan/Documents/PyCal"


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
    service: Resource = gc_creds(directory, cal_secret_name)

    t_del = Thread(target=del_events, args=(now, service, calendar_id, param))

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Deleting events...", "white", attrs=['bold']))

    t_del.start()

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Fetching events...", "white", attrs=['bold']))

    # use only the event sheet within the workbook
    events_sheet: Worksheet = sheets[0]
    es_conact_sheet: Worksheet = sheets[1]
    cpt_contact_sheet: Worksheet = sheets[2]

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
    department = events_sheet.get_col(7)
    sound = events_sheet.get_col(8)
    lights = events_sheet.get_col(9)
    records = events_sheet.get_col(9)
    stage = events_sheet.get_col(10)
    b_cast = events_sheet.get_col(11)
    video = events_sheet.get_col(12)
    temp = events_sheet.get_col(13)
    event_coords = events_sheet.get_col(14)

    event_coord_names = es_conact_sheet.get_col(1)
    event_coord_nums = es_conact_sheet.get_col(4)

    cpt_names = cpt_contact_sheet.get_col(1)
    cpt_nums = cpt_contact_sheet.get_col(3)

    t_del.join()

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Adding events to calendar...", "white", attrs=['bold']))

    for i in my_events_rows:

        if dates[i] == "":
            continue

        try:
            date_obj = datetime.strptime(dates[i], "%A-%b-%d-%y")
        except ValueError:
            continue
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
        try:
            event_coord_num = event_coord_nums[event_coord_names.index(
                event_coords[i])]
        except ValueError:
            event_coord_num = ""

        descripion = ('Automatic creation' +
                      '\nEvent Start Time: ' + start +
                      '\nEvent Coordinator: ' + event_coord + " " +
                      event_coord_num +
                      '\nDepartment: ' + department[i] +
                      '\nRecord: ' + record +
                      '\nSound: ' + sound[i] +
                      '\nBroadcast: ' + b_cast[i] +
                      '\nTemp: ' + temp[i] +
                      '\nLights: ' + lights[i] +
                      '\nStage: ' + stage[i] +
                      '\nVideo: ' + video[i] +
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
        minute = int(time_string.split(':')[1][:2])
        if time_string.split(':')[1][-2:] == "PM" and hour != 12:
            hour += 12
    else:
        hour = 23
        minute = 59
    return hour, minute


def add_cal_event(
    end_string, now, end_time, param, name, service: Resource, calendar_id, calevent
) -> None:
    """Add event to calendar"""
    if hasnumbers(end_string) and now < end_time:
        if param == "v":
            print('  ', name)
        try:
            add_event(service, calendar_id, calevent)
        except HttpError:
            sleep(1)
            add_event(service, calendar_id, calevent)


def add_event(service: Resource, calendar_id, calevent) -> None:
    """Add event to service"""
    service.events().insert(calendarId=calendar_id,
                            body=calevent).execute()


def get_event_rows(events_sheet, initials) -> list:
    """Get events that {initals} are working"""
    if initials == "ANY":
        return list(range(1, events_sheet.rows))
    my_events = events_sheet.find(initials)
    my_events_rows = [event.row - 1 for event in my_events]
    everyone_events = events_sheet.find(
        'ALL', cols=(1, 8, 9), matchEntireCell=True)
    my_events_rows.extend(event.row - 1 for event in everyone_events)
    return sorted(list(set(my_events_rows)))


def del_events(now, service: Resource, calendar_id, param) -> None:
    """Delete future events"""
    time_min = now.isoformat('T') + "-05:00"
    events = service.events()

    while True:
        result = events.list(calendarId=calendar_id,
                             timeMin=time_min, singleEvents=True,
                             orderBy="startTime").execute()
        delete_events_id = [
            result['items'][i]['id']
            for i in range(len(result['items']))
            if "description" in result["items"][i]
            and result["items"][i]["description"].startswith(
                'Automatic creation'
            )
            and datetime.strptime(
                result["items"][i]["start"]["dateTime"][:-6],
                "%Y-%m-%dT%H:%M:%S",
            )
            > now
        ]

        for i in delete_events_id:
            try:
                events.delete(calendarId=calendar_id,
                              eventId=i).execute()
            except (HttpError, TimeoutError):
                try:
                    sleep(1)
                    events.delete(calendarId=calendar_id,
                                  eventId=i).execute()
                except (HttpError, TimeoutError):
                    pass
        if not delete_events_id:
            break

        if param == "v":
            print('\t', "Events Deleted")


def gc_creds(directory, cal_secret_name) -> Resource:
    """Authorize access to google calendar"""
    pklstr = f'{directory}/token.pkl'
    scopes = ['https://www.googleapis.com/auth/calendar']
    if exists(pklstr):
        with open(pklstr, "rb") as token:
            credentials = load(token)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            f'{directory}/{cal_secret_name}', scopes=scopes
        )

        credentials = flow.run_local_server(port=0)
        with open(pklstr, "wb") as token:
            dump(credentials, token)
    return build("calendar", "v3", credentials=credentials)


def sh_creds(directory, sheet_secret_name) -> Spreadsheet:
    """Get google sheet"""
    gcal = authorize(
        client_secret=f'{directory}/{sheet_secret_name}',
        credentials_directory=directory,
        local=True
    )

    return gcal.open_by_key("1f3G7XkZtH4vJJ0qoPjjsOnRHcnWurceP83_cKqI_8jE")


def check_connection(param: str = None) -> bool:
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
    except (ConnError, Timeout):
        if param == "v":
            print('\t', "Connection failed, exiting.")
        return False


def run_adam():
    """run_adam"""
    calhelp(
        initials="ADH",
        calendar_id="adamdh00@gmail.com",
        directory=f"{BASEPATH}/credentials",
        sheet_secret_name="SCS.json",
        cal_secret_name="PCS.json",
        color_id=11,
        param=ARG)


def run_bri():
    """run_bri"""
    calhelp(
        initials="BJ",
        calendar_id="brijans19@gmail.com",
        directory=f"{BASEPATH}/credentials/bricreds",
        sheet_secret_name="SCS.json",
        cal_secret_name="PCS.json",
        color_id=7,
        param=None
    )


def run_brooks():
    """run brooks"""
    calhelp(
        initials="BT",
        calendar_id="Xobr21037@gmail.com",
        directory=f"{BASEPATH}/credentials/brookscreds",
        sheet_secret_name="SCS.json",
        cal_secret_name="PCS.json",
        color_id=1,
        param=None
    )


def run_calvin():
    """run calvin"""
    calhelp(
        initials="ANY",
        calendar_id="rdfvrbd06kr0fsa49g4h771l8c@group.calendar.google.com",
        directory=f"{BASEPATH}/credentials",
        sheet_secret_name="SCS.json",
        cal_secret_name="PCS.json",
        color_id=3,
        param=ARG)


def main():
    """Main()"""
    Thread(target=run_adam).start()
    Thread(target=run_bri).start()
    Thread(target=run_brooks).start()


if __name__ == '__main__' and check_connection(ARG):
    main()
