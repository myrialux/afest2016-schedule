#!/usr/bin/env python


import csv
import re
from openpyxl import *


class AFestEvent:
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
        self.desc = row[4].value.strip()
        self.location = row[5].value.strip()
        self.track = row[6].value.strip()
        self.attendify_id = row[7].value.strip()

        regex = re.compile(r".*{afestid:(.+)}$", re.IGNORECASE)
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


def load_attendify_events(file_name):
    wb = load_workbook(file_name)

    schedule_sheet = wb["Schedule"]
    if not schedule_sheet:
        print("ERROR - Schedule sheet not found in workbook")
        sys.exit(RETURN_VALUE_WORKBOOK_ERROR)

    events = []

    events_range = "A6:H" + str(schedule_sheet.max_row)
    for row in schedule_sheet.iter_rows(range_string=events_range):
        event = AFestEvent()
        event.load_from_attendify(row)
        # print event.title
        events.append(event)

    print("Total Attendify Events: " + str(len(events)))
    return events


def load_afest_events(file_name):
    events = []

    with open(file_name) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            event = AFestEvent()
            event.load_from_afest(row)
            events.append(event)

    print("Total A-Fest Events: " + str(len(events)))
    return events


def main():
    RETURN_VALUE_SUCCESS           = 0
    RETURN_VALUE_INVALID_PARAMETER = 1
    RETURN_VALUE_WORKBOOK_ERROR    = 2

    import sys
    import argparse

    parser = argparse.ArgumentParser(description="Show schedule data")
    parser.add_argument("attendify_file", help="Attendify .xlsx schedule")
    parser.add_argument("afest_file", help="A-Fest .csv schedule")
    args = parser.parse_args()

    attendify_events = load_attendify_events(args.attendify_file)
    afest_events = load_afest_events(args.afest_file)


if __name__ == "__main__":
    main()
