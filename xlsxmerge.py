#!/usr/bin/env python3

import xlrd
import xlsxwriter
import sys


workbook = xlsxwriter.Workbook('MergeTest.xlsx')
worksheet = workbook.add_worksheet()

master_file = xlrd.open_workbook(sys.argv[1])
additional_files = sys.argv[1:]
headings = {x:y for y,x in enumerate(master_file.sheet_by_index(0).row_values(0))}


def write_rows(filename):
    for sheet in filename.sheet_names():
        print('Collating ' + sheet)
        temp_rows = range(0,filename.sheet_by_name(sheet).nrows)
        for rw in temp_rows:
            try: 
                row += 1
            except NameError:
                row = 0
            col_names = dict(enumerate(filename.sheet_by_name(sheet).row_values(0)))
            values = filename.sheet_by_name(sheet).row_values(rw)
            for x in range(len(col_names)):
                try:
                    col = headings[col_names[x]]
                except KeyError:
                    add_headings(col_names[x])
                    col = headings[col_names[x]]
                        
                worksheet.write(row, col, values[x])

def add_headings(col_name):
   headings[col_name] = len(headings)
   worksheet.write(0, headings[col_name], col_name)

for files in additional_files:
    print('Opening file ' + files)
    filename = xlrd.open_workbook(files) 
    write_rows(filename)

workbook.close()
