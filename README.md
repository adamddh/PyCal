# PyCal

This is a file & script which contains the necessary code to take events listed on a master sheet of events and add event event which I am associated with to my personal calendar.

## Expansion

This was later expanded to include functionality for my coworkers calendars, as well as a calendar that would contain all events.

## Files

### - py_all.py

This imports functionality from pythoncalendar_v3.py to be automated by cron with a different schedule

### - pycalv4.py _no longer used/maintained_

This was built to do the same job as v3, but instead of working with a google sheet stored online,used the excel sheet stored in the same location. Functionality of accessing the sheet and the data had to change, but the accessing the calendar was the same.

### - pythoncalendar_v3.py

The workhorse. Contains functionality for getting events and putting them in the calendar, as well as running this directly from the command line.

### - redline.py

Detects changes in the redline, so one can schedule themselves for more events. \*_This is slightly devious but it's public code._
