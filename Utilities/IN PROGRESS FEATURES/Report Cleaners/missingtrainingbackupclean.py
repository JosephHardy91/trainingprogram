# -*- coding: utf-8 -*-
# Welcome to the DataNitro Editor
# Use Cell(row,column).value, or Cell(name).value, to read or write to cells
# Cell(1,1) and Cell("A1") refer to the top-left cell in the spreadsheet
# 
# Note: To run this file, save it and run it from Excel (click "Run from File")
# If you have trouble saving, try using a different directory

yesses=['QAM R9']
for sheet in all_sheets():
    active_sheet(sheet)
    i=1
    while not Cell(i,"A").is_empty():
        if Cell(i,"I").is_empty():
            if ' '.join([Cell(i,"A").value,Cell(i,"B").value]) in yesses:
                Cell(i,"I").value='Yes'
            else:
                Cell(i,"I").value='No'
        i+=1
