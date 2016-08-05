#!/usr/bin/env python


import sys
import csv
import re
from openpyxl import *


RETURN_VALUE_SUCCESS           = 0
RETURN_VALUE_INVALID_PARAMETER = 1
RETURN_VALUE_WORKBOOK_ERROR    = 2

ATTENDIFY_SCHEDULE_SHEET_NAME = "Schedule"

ATTENDIFY_DESC_COL_INDEX = 4
ATTENDIFY_ID_COL_INDEX = 7


class AFestEvent:
    """Model object for events, whether from the AFest schedule or Attendify."""

    title = ""
    date = ""
    start_time = ""
    end_time = ""
    desc = ""
    location = ""
    track = ""
    attendify_id = None
    afest_id = None

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

    def load_from_afest(self, row):
        self.title = row["Session Title"]
        self.date = row["Date"]
        self.start_time = row["Start Time"]
        self.end_time = row["End Time"]
        self.desc = row["Description"]
        self.location = row["Location"]
        self.track = row["Track Title"]
        self.attendify_id = row["UID"]
        self.afest_id = row["id_schedule_block"]

    def is_match(self, other):
        """Returns whether the other event matches this one for purposes of copying the AFest ID over.
        Not the same as an equality function, because we don't care about the track or description.
        """

        return (self.date == other.date) and (self.start_time == other.start_time) and (self.end_time == other.end_time) and (self.location == other.location) and (self.title == other.title)


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
