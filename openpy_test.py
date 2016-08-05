#!/usr/bin/env python


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


def main():
    RETURN_VALUE_SUCCESS           = 0
    RETURN_VALUE_INVALID_PARAMETER = 1
    RETURN_VALUE_WORKBOOK_ERROR    = 2

    import sys
    import argparse

    parser = argparse.ArgumentParser(description="Show schedule data")
    parser.add_argument("input", help="Attendify .xlsx schedule file")
    args = parser.parse_args()

    wb = load_workbook(args.input)

    schedule_sheet = wb["Schedule"]
    if not schedule_sheet:
        print("ERROR - Schedule sheet not found in workbook")
        sys.exit(RETURN_VALUE_WORKBOOK_ERROR)

    events = {}

    events_range = "A6:H" + str(schedule_sheet.max_row)
    event_count = 0
    for event in schedule_sheet.iter_rows(range_string=events_range):
        event_count += 1

        afe = AFestEvent()
        afe.load_from_attendify(event)
        print afe.title

    print("Total Events: " + str(event_count))


if __name__ == "__main__":
    main()
