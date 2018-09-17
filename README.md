# xlsxmerge.py
Merges Excell Sheets and Workbooks into one Excel Sheet

This project was designed for a specific use case, but by design it is
flexible. Any other uses would require some tweaking of the code, but
not much.  

## Installing 

`xlsxmerge.py` requires Python 3+ and the `xlrd` and `xlsxwrite`
modules. They can be installed with `pip`.

## Using xlsxmerge.py

To merge `file-1.xlsx`, `file-2.xlsx`, etc. into `merged.xlsx` run:

    $ python xlsxmerge.py file-1.xlsx file-2.xlxs [... file-n.xlsx] merged.xlsx

This assumes that the first row of every sheet has a heading. The first
row, of the first sheet's file is the one that will be used as a "Master
Heading": all the other sheets will be organized according to these
headings. If one of the sheets has an additional heading (or differently
named heading!), the code will create a new column with this additional
heading name.
