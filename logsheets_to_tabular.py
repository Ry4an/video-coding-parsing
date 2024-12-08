#!/usr/bin/env python

import sys
import os
import re
import datetime
from collections import defaultdict
import csv
from openpyxl import load_workbook

VIDEO_CODE_RANGE = 'A2:C11'
INTERVALS_RANGE = 'B13:T16'  # up to 19 Intervals, but max seen is 14

OUTPUT_FIELDNAMES = [  # controls order and inclusion
        'filename',
        'id',
        'wave',
        'fm',
        'initials',
        'q1a_code',
        'q1a_note',
        'q1b_code',
        'q1b_note',
        'q2_code',
        'q2_note',
        'q3_code',
        'q3_note',
        'q4_code',
        'q4_note',
        'q5_code',
        'q5_note',
        'q6_code',
        'q6_note',
        'q7_code',
        'q7_note',
        'q8_code',
        'q8_note',
        'q9_code',
        'q9_note',
        'interval1_times',
        'interval1_q10',
        'interval1_q11',
        'interval2_times',
        'interval2_q10',
        'interval2_q11',
        'interval3_times',
        'interval3_q10',
        'interval3_q11',
        'interval4_times',
        'interval4_q10',
        'interval4_q11',
        'interval5_times',
        'interval5_q10',
        'interval5_q11',
        'interval6_times',
        'interval6_q10',
        'interval6_q11',
        'interval7_times',
        'interval7_q10',
        'interval7_q11',
        'interval8_times',
        'interval8_q10',
        'interval8_q11',
        'interval9_times',
        'interval9_q10',
        'interval9_q11',
        'interval10_times',
        'interval10_q10',
        'interval10_q11',
        'interval11_times',
        'interval11_q10',
        'interval11_q11',
        'interval12_times',
        'interval12_q10',
        'interval12_q11',
        'interval13_times',
        'interval13_q10',
        'interval13_q11',
        'interval14_times',
        'interval14_q10',
        'interval14_q11',
]

FILENAME_LABELS = ['id', 'wave', 'fm', 'initials']

DURATION_CODES = ['1a', '1b', '5', '6']  # values are hh:mm:ss

out_csv = csv.DictWriter(sys.stdout, OUTPUT_FIELDNAMES, extrasaction='ignore')
out_csv.writeheader()

warnings = defaultdict(list)  # filename -> list(messages)

def warn_check(expected, cell, filename):
    """add a warning if cell values don't match expected string"""
    if expected == cell.value:
        return
    warnings[filename].append(f"'{expected}' expected, but '{cell.value}' "
        + f"found in cell {cell.column_letter}{cell.row}")

DURATION_RE = re.compile(r'^(?:(?P<hours>\d+):)?(?P<minutes>\d?\d)?:(?P<seconds>\d\d)$')

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
        continue

    sheet = wb.active

    # check headers
    warn_check("Question", sheet['A1'], filename)
    warn_check("Code"    , sheet['B1'], filename)

    try:
        # get code and note for each of the whole-video questions
        for row in sheet[VIDEO_CODE_RANGE]:
            (question, code, note) = [ cell.value for cell in row ]
            question_num = question.split(".")[0]
            outrow[f"q{question_num}_code"] = code
            if question_num in DURATION_CODES:
                if isinstance(code, datetime.time):
                    code = f"{code:%H:%M}" # sheets have hours & mins!
                if (match := DURATION_RE.match(code.strip())) is not None:
                    outrow[f"q{question_num}_code"] = "{h}:{m:02d}:{s:02d}".format(
                            h=int(match.group('hours') or 0),
                            m=int(match.group('minutes') or 0),
                            s=int(match.group('seconds')))
                else:
                    warnings[filename].append(f"Unparsable duration '{code}'"
                        + f" found in question {question_num}")
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
            if name is None:
                break
            try:
                (label, interval_num) = name.strip().split(" ")
                if label != "Interval":
                    warnings[filename].append(f"Non interval label: '{label}'")
            except ValueError:
                warnings[filename].append(
                        f"Skipping unparsable interval label: '{name}'")
                next

            outrow[f"interval{interval_num}_times"] = time
            outrow[f"interval{interval_num}_q10"  ]  = q10
            outrow[f"interval{interval_num}_q11"  ]  = q11

        # SKIP warn_check(None, sheet['A17'], filename)
    except Exception as ex:
        warnings[filename].append(f"Failure to parse. {ex} Skipping.")

    out_csv.writerow(outrow)

if warnings:
    warnings_csv = csv.writer(sys.stderr)
    warnings_csv.writerow(['filename', 'warning'])
    for filename, messages in warnings.items():
        for message in messages:
            warnings_csv.writerow([filename, message])
