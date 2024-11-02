#!/usr/bin/env python

from openpyxl import load_workbook

FILENAME_LABELS = ['id', 'wave', 'fm', 'initials']
CODE_ROW_START  = 2
CODE_ROW_END    = 11  # inclusive

filename = '3961_Wave1_FMO2_JK.xlsx'

outrow = dict(zip(FILENAME_LABELS, filename[:-5].split("_")))

try:
    wb = load_workbook(filename=filename,
            data_only=True,
            read_only=True)
except Exception as ex:
    print(f"Unable to load {filename}", ex)

sheet = wb.active

for row in sheet.iter_rows(min_row=CODE_ROW_START,
        max_col=3, max_row=CODE_ROW_END):
    (question, code, note) = [ cell.value for cell in row ]
    question_num = question.split(".")[0]
    outrow[f"question{question_num}_code"] = code
    outrow[f"question{question_num}_note"] = note

print(outrow)
