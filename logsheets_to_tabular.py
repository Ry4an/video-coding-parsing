#!/usr/bin/env python

import sys
import os
from collections import defaultdict
from openpyxl import load_workbook

FILENAME_LABELS = ['id', 'wave', 'fm', 'initials']
CODE_ROW_START  = 2
CODE_ROW_END    = 11  # inclusive

VIDEO_CODE_RANGE = 'A2:C11'
INTERVALS_RANGE = 'B13:H16'

warnings = defaultdict(list)  # filename -> list(messages)


def warn_check(expected, cell, filename):
    """add a warning if cell values don't match expected string"""
    if expected == cell.value:
        return
    warnings[filename].append(f"'{expected}' expected, but '{cell.value}' "
        + f"found in cell {cell.column_letter}{cell.row}")


for filepath in sys.argv[1:]:
    filename = os.path.basename(filepath)

    outrow = dict(zip(FILENAME_LABELS, filename[:-5].split("_")))
    outrow['filename'] = filename

    try:
        wb = load_workbook(filename=filepath,
                data_only=True,
                read_only=True)
    except Exception as ex:
        warnings[filename].append(f"Unable to load. {ex} Skipping.")

    sheet = wb.active

    # check headers
    warn_check("Question", sheet['A1'], filename)
    warn_check("Code"    , sheet['B1'], filename)

    # get code and note for each of the whole-video questions
    for row in sheet[VIDEO_CODE_RANGE]:
        (question, code, note) = [ cell.value for cell in row ]
        question_num = question.split(".")[0]
        outrow[f"q{question_num}_code"] = code
        outrow[f"q{question_num}_note"] = note

    # get interval names, times, and questions
    interval_cells = sheet[INTERVALS_RANGE]
    warn_check("Interval 1", interval_cells[0][0], filename)

    names = [ cell.value for cell in interval_cells[0] ]
    times = [ cell.value for cell in interval_cells[1] ]
    q10s  = [ cell.value for cell in interval_cells[2] ]
    q11s  = [ cell.value for cell in interval_cells[3] ]
    for interval in zip(names, times, q10s, q11s):
        (name, time, q10, q11) = interval
        interval_num = name.split(" ")[-1]
        outrow[f"interval{interval_num}_times"] = time
        outrow[f"interval{interval_num}_q10"  ]  = q10
        outrow[f"interval{interval_num}_q11"  ]  = q11

    print(outrow)

print(warnings)
