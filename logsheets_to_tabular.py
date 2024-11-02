#!/usr/bin/env python

import sys
import os
from collections import defaultdict
from openpyxl import load_workbook

FILENAME_LABELS = ['id', 'wave', 'fm', 'initials']
CODE_ROW_START  = 2
CODE_ROW_END    = 11  # inclusive

warnings = defaultdict(list)  # filename -> list(messages)

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

    # get code and note for each of the whole-video questions
    for row in sheet.iter_rows(min_row=CODE_ROW_START,
            max_col=3, max_row=CODE_ROW_END):
        (question, code, note) = [ cell.value for cell in row ]
        question_num = question.split(".")[0]
        outrow[f"question{question_num}_code"] = code
        outrow[f"question{question_num}_note"] = note

    print(outrow)

print(warnings)
