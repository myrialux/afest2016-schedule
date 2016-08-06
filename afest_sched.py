#!/usr/bin/env python

"""
Code and classes to handle keeping the AnimeFest 2016 app's schedule in-sync with the version from the web site.
"""

import sys
import csv
import re
from datetime import datetime
from openpyxl import *


RETURN_VALUE_SUCCESS               = 0
RETURN_VALUE_INVALID_PARAMETER     = 1
RETURN_VALUE_WORKBOOK_ERROR        = 2
RETURN_VALUE_EXCESSIVE_ID_USE      = 3
RETURN_VALUE_SPLIT_EVENTS_SAME_DAY = 4

ATTENDIFY_SCHEDULE_SHEET_NAME = "Schedule"

ATTENDIFY_DESC_COL_INDEX = 4
ATTENDIFY_ID_COL_INDEX = 7

ATTENDIFY_DATE_FORMAT = "%m/%d/%Y"


class AFestEvent:
    """Model object for events, whether from the AFest schedule or Attendify."""

    def load_from_attendify(self, row):
        self.title = row[0].value.strip()
        self.date = row[1].value.strip()
        self.start_time = row[2].value.strip()
        self.end_time = row[3].value.strip()
        self.desc = row[ATTENDIFY_DESC_COL_INDEX].value.strip()
        self.location = row[5].value.strip()
        self.track = (row[6].value or "").strip()
        self.attendify_id = row[ATTENDIFY_ID_COL_INDEX].value.strip()

        regex = re.compile(r".*\[afestid:(.+)\]$", re.IGNORECASE)
        match = regex.search(self.desc)
        if match:
            self.afest_id = match.group(1)
        else:
            self.afest_id = None

    def load_from_afest(self, row):
        self.title = row["Session Title"].strip()
        self.date = row["Date"].strip()
        self.start_time = row["Start Time"].strip()
        self.end_time = row["End Time"].strip()
        self.desc = row["Description"].strip()
        self.location = row["Location"].strip()
        self.track = row["Track Title"].strip()
        self.attendify_id = row["UID"].strip()
        self.afest_id = row["id_schedule_block"].strip()

    def is_match(self, other):
        """Returns whether the other event matches this one for purposes of copying the AFest ID over.
        Not the same as an equality function, because we don't care about the track or description.
        """

        # Common checks first
        if (self.title != other.title) or (self.location != other.location):
            return False

        if (self.date == other.date):
            # If the days match then the start time needs to, and the end time either needs to match, or needs to be the end of the day. This helps with the first portion of events split over midnight.
            return (self.start_time == other.start_time) and ((self.end_time == other.end_time) or (self.end_time == "23:59"))
        else:
            # If this event is on the day after the other one, starts at midnight, and its end time matches, then this is the second part of a split-over-midnight event.
            thisDate = datetime.strptime(self.date, ATTENDIFY_DATE_FORMAT)
            otherDate = datetime.strptime(other.date, ATTENDIFY_DATE_FORMAT)
            return ((thisDate - otherDate).days == 1) and (self.start_time == "00:00") and (self.end_time == other.end_time)


def open_attendify_schedule(file_name):
    wb = load_workbook(file_name)

    schedule_sheet = wb[ATTENDIFY_SCHEDULE_SHEET_NAME]
    if not schedule_sheet:
        print("ERROR - Schedule sheet not found in workbook")
        sys.exit(RETURN_VALUE_WORKBOOK_ERROR)
    
    return wb


def iter_attendify_schedule_rows(sheet):
    """Returns an iterator for the schedule rows in the given sheet."""

    events_range = "A6:H" + str(sheet.max_row)
    return sheet.iter_rows(range_string=events_range)


def load_attendify_events(file_name):
    wb = open_attendify_schedule(file_name)
    schedule_sheet = wb[ATTENDIFY_SCHEDULE_SHEET_NAME]

    events = []

    for row in iter_attendify_schedule_rows(schedule_sheet):
        event = AFestEvent()
        event.load_from_attendify(row)
        events.append(event)

    return events


def load_afest_events(file_name):
    events = []

    with open(file_name) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            event = AFestEvent()
            event.load_from_afest(row)
            events.append(event)

    return events


def add_afest_id_to_attendify(workbook, attendify_id, afest_id):
    schedule_sheet = workbook[ATTENDIFY_SCHEDULE_SHEET_NAME]

    for row in iter_attendify_schedule_rows(schedule_sheet):
        if row[ATTENDIFY_ID_COL_INDEX].value.strip() == attendify_id:
            row[ATTENDIFY_DESC_COL_INDEX].value += "\n\n[afestid:{0}]".format(afest_id)
            break


def merge_events(event1, event2):
    if event1.date == event2.date:
        print("ERROR - Events to merge share the same date ({0}) - \"{1}\" and \"{2}\"".format(event1.date, event1.title, event2.title))
        sys.exit(RETURN_VALUE_SPLIT_EVENTS_SAME_DAY)

    date1 = datetime.strptime(event1.date, ATTENDIFY_DATE_FORMAT)
    date2 = datetime.strptime(event2.date, ATTENDIFY_DATE_FORMAT)

    if date1 < date2:
        startEvent = event1
        endEvent = event2
    else:
        startEvent = event2
        endEvent = event1

    result = AFestEvent()
    result.title = startEvent.title
    result.date = startEvent.date
    result.start_time = startEvent.start_time
    result.end_time = endEvent.end_time
    result.desc = startEvent.desc
    result.location = startEvent.location
    result.track = startEvent.track
    result.attendify_id = startEvent.attendify_id
    result.afest_id = startEvent.afest_id
    return result


def merge_split_events(events):
    """Takes the given list of events and returns a new list with all events that were split across midnight merged into a single event with the proper start and end time.
    The date field of merged events is the date of the start of the event.
    """

    by_afest_id = {}
    for event in events:
        id_list = by_afest_id[event.afest_id]
        if not id_list:
            id_list = [event]
        else:
            id_list.append(event)
        by_afest_id[event.afest_id] = id_list

    result = []
    for afest_id in by_afest_id.keys:
        id_list = by_afest_id[afest_id]
        if len(id_list) == 1:
            result.append(id_list[0])
        else if len(id_list) == 2:
            result.append(merge_events(id_list[0], id_list[1]))
        else:
            print("ERROR - {0} event(s) for AFest ID {1}".format(len(id_list), afest_id))
            sys.exit(RETURN_VALUE_EXCESSIVE_ID_USE)
    
    return result


def add_ids_to_attendify(args):
    afest_events = load_afest_events(args.afest_file)
    attendify_events = load_attendify_events(args.attendify_file)

    print("AFest: {0}  Attendify: {1}".format(len(afest_events), len(attendify_events)))

    workbook = open_attendify_schedule(args.attendify_file)

    exact_matches = 0
    title_matches = 0
    for at_event in attendify_events:
        if at_event.afest_id:
            continue

        # Exact matches first
        matched = False
        for af_event in afest_events:
            if at_event.is_match(af_event):
                add_afest_id_to_attendify(workbook, at_event.attendify_id, af_event.afest_id)

                matched = True
                exact_matches += 1
                break
        if matched:
            continue
        
        # If the event has exactly one match against the titles in the AFest schedule, use that one.
        matched_afest_ids = []
        for af_event in afest_events:
            if at_event.title == af_event.title:
                matched_afest_ids.append(af_event.afest_id)
        if len(matched_afest_ids) == 1:
            add_afest_id_to_attendify(workbook, at_event.attendify_id, matched_afest_ids[0])
            title_matches += 1

    workbook.save(args.attendify_file + ".new")
    
    print("Exact Matches: {0}  Title Matches: {1}".format(exact_matches, title_matches))


def check_afest_ids_in_attendify(args):
    attendify_events = load_attendify_events(args.attendify_file)

    missing_ids = 0
    for event in attendify_events:
        if not event.afest_id:
            print("{0}: {1} {2}-{3} {4}".format(event.title, event.date, event.start_time, event.end_time, event.location))
            missing_ids += 1

    print("Total events: {0}  Missing IDs: {1}".format(len(attendify_events), missing_ids))


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Show schedule data")
    subparsers = parser.add_subparsers()

    add_ids_parser = subparsers.add_parser("add_ids", help="Attempt to match AFest and Attendify events. Copy AFest IDs to matched Attendify events")
    add_ids_parser.add_argument("afest_file", help="AFest .csv schedule")
    add_ids_parser.add_argument("attendify_file", help="Attendify .xlsx schedule")
    add_ids_parser.set_defaults(func=add_ids_to_attendify)

    check_ids_parser = subparsers.add_parser("check_ids", help="Check the given Attendify schedule for missing AFest IDs.")
    check_ids_parser.add_argument("attendify_file", help="Attendify .xlsx schedule")
    check_ids_parser.set_defaults(func=check_afest_ids_in_attendify)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
