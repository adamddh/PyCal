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

import sys
from datetime import datetime
from io import BytesIO
from os.path import exists
from pickle import dump, load
from time import sleep

from google.auth.exceptions import TransportError
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import Resource, build
from googleapiclient.http import MediaIoBaseDownload
from pandas import read_excel
from termcolor import colored

arg = sys.argv[1] if len(sys.argv) > 1 else None


def hasnumbers(inputstring: str) -> bool:
    """
    Check if string has numbers in it

    Args:
        inputstring (str): string to check for numbers

    Returns:
        bool: Wether or not the string has numbers
    """
    return any(char.isdigit() for char in inputstring)


def auth_service(pkl_path: str, scopes: list[str], serv_type: str) -> Resource:
    """
    Authenticate a Google API

    Args:
        pkl_path (str): Path to pickle file if it exists, else create one there
        scopes (list[str]): Scopes to authenticate the API for
        serv_type (str): Service type to authenticate API for

    Returns:
        (Resource): Google API Resource
    """
    if exists(pkl_path):
        with open(pkl_path, "rb") as token:
            credentials = load(token)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            f"{'/'.join(pkl_path.split('/')[:-1])}/PCS.json",
            scopes=scopes)
        credentials = flow.run_local_server(port=0)
        with open(pkl_path, "wb") as token:
            dump(credentials, token)
    return build(serv_type, "v3", credentials=credentials)


def add_event(
    calendar_id, param, now, service1, end_time, name, calevent
):
    """
    Add event to google calendar, print param if given.

    Args:
        calendar_id (str): gmail address of associated google calendar.
        param (str): verbosity option
        now (datetime): datetime.now() from earlier
        service1 (Resource): google API resource.
        ends (list[str]): list of endtimes from excel sheet
        name (str): event name
        calevent (dict): dictionary of details for calendar event
    """
    if hasnumbers(str(end_time.time())) and now < end_time:
        if param == "v":
            print('  ', name)
        try:
            service1.events().insert(calendarId=calendar_id,
                                     body=calevent).execute()
        except Exception:
            sleep(1)
            service1.events().insert(calendarId=calendar_id,
                                     body=calevent).execute()


def authenticate_services(directory, param):
    """
    Authenticate both services

    Args:
        directory (str): Location of access files
        param (str): Verbosity parameter

    Returns:
        tuple[Resource]: Both authenticated services
    """
    SCOPES0 = [  # pylint: disable-msg=C0103
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/drive.readonly'
    ]
    service0 = auth_service(f"{directory}/token0.pkl", SCOPES0, "drive")

    SCOPES1 = [  # pylint: disable-msg=C0103
        'https://www.googleapis.com/auth/calendar']
    service1 = auth_service(f"{directory}/token1.pkl", SCOPES1, "calendar")

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Deleting events...", "white", attrs=['bold']))

    return service0, service1


def get_file(service0):
    """
    Get excel sheet from google drive and store it in /tmp

    Args:
        service0 (Resource): Google Api Resource.
    """
    FILE_ID = "1Rb0JxnQKpM55uYis4TaldQK5PM-yhG4S"  # pylint: disable-msg=C0103
    req = service0.files().get_media(fileId=FILE_ID)
    file_download = BytesIO()
    downloader = MediaIoBaseDownload(fd=file_download, request=req)
    done = False

    while not done:
        done = downloader.next_chunk()[1]

    file_download.seek(0)

    with open("/tmp/data.xlsx", "wb") as fout:
        fout.write(file_download.read())


def delete_events(calendar_id, now, service1):
    """
    Delete future events with string in description.

    Iterate over all events, then reiterate until no events were deleted.

    Args:
        calendarId (str): gmail address associated with google calendar account.
        now (datetime): datetime.now() from earlier
        service1 (Resouce): Google Api Resource.
    """
    while True:
        j = 0
        time_min = now.isoformat('T') + "-05:00"
        events = service1.events()
        try:
            result = events.list(
                calendarId=calendar_id,
                timeMin=time_min,
                singleEvents=True,
                orderBy="startTime"
            ).execute()
        except TransportError:
            sys.exit(1)
        delete_events_id = []
        for i in range(len(result['items'])):
            try:
                if result['items'][i]['description']\
                        .startswith('Automatic creation',):
                    delete_events_id.append(result['items'][i]['id'])
            except Exception:   # not all event have a feild 'description', and
                pass            # raises an error if there is no feild.
                # Addressed by skipping event as its not ours

        for i in delete_events_id:
            try:
                events.delete(calendarId=calendar_id,
                              eventId=i).execute()
                j += 1
            except:
                sleep(1)
                events.delete(calendarId=calendar_id,
                              eventId=i).execute()
                j += 1
        if j == 0:
            break


def calhelp(initials, calendar_id, directory, param, color_id=11):
    """designed to facilitate the use of v2

    Args:
        initials (string): intials to search for on the sheet
        calendarId (string): calendar ID of the calendar to write events to
        directory (string): location of where to store credential files
        param (str): Verbosity parameter
        colorID (int, optional): color of the events on the calendar. Defaults to 11.
    """

    now = datetime.now()

    service0, service1 = authenticate_services(directory, param)

    delete_events(calendar_id, now, service1)

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Fetching events...", "white", attrs=['bold']))

    get_file(service0)

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Finding my events...", "white", attrs=['bold']))

    events_sheet = read_excel(io="/tmp/data.xlsx", sheet_name="Schedule")

    dates = events_sheet["DATE"].tolist()
    titles = events_sheet["EVENT"].tolist()
    calls = events_sheet["CALL"].tolist()
    starts = events_sheet["START"].tolist()
    ends = events_sheet["END"].tolist()
    locations = events_sheet["LOCATION"].tolist()
    records = events_sheet["Record/ Livestream"].tolist()
    event_coords = events_sheet["Event Coordinator"].tolist()
    sounds = events_sheet["SOUND"].tolist()

    my_events_rows = [row for row, data in enumerate(
        sounds) if initials in str(data) or "all" in str(data).lower()]

    if param == "v":
        print(colored("==>", "green", attrs=['bold']), colored(
            "Adding events to calendar...", "white", attrs=['bold']))

    for i in my_events_rows:

        # year/month/date ints
        try:
            start_time = datetime.combine(dates[i].date(), calls[i])
            end_time = datetime.combine(dates[i].date(), ends[i])
        except Exception:
            continue

        # record string
        record = "Yes" if records[i] == "Yes" else "No"
        # start string

        # Event coordinator string
        event_coord = event_coords[i]

        calevent = {
            'summary': titles[i],
            'location': locations[i],
            'colorId': color_id,          # where you can select the color of the event
            'description': (f'Automatic creation: {datetime.now()}\n' +
                            f'Event Start Time: {starts[i]}\n' +
                            f'Event Coordinator: {event_coord}\n' +
                            f'Record: {record}'),

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
        add_event(calendar_id, param, now, service1,
                  end_time, titles[i], calevent)


def main():
    """ Main() """
    calhelp(initials="ADH",
            calendar_id="adamdh00@gmail.com",
            directory="/Users/adamdenhaan/Documents/PyCal/credentials",
            color_id=11,
            param=arg)


if __name__ == '__main__':
    main()
