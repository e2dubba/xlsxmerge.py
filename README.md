# xlsxmerge.py
Merges Excell Sheets and Workbooks into one Excell Sheet

This project was designed for a specific use case, but by design it is flexible. Any other uses would require some tweaking of the code, but not much. 

The syntax works by calling the command `xlsxmerge.py file1.xlsx file2.xlxs etc`. This assumes that the first row of every sheet has a heading. The first row, of the first sheet's file is the one that will be used as a "Master Heading": all the other sheets will be organized according to these headings. 

The output will be called `MergeTest.xlsx`. This could easily be adjusted in the code. 
