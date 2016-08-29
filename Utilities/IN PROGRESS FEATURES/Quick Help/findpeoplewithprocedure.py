# -*- coding: utf-8 -*-
# Welcome to the DataNitro Editor
# Use Cell(row,column).value, or Cell(name).value, to read or write to cells
# Cell(1,1) and Cell("A1") refer to the top-left cell in the spreadsheet
# 
# Note: To run this file, save it and run it from Excel (click "Run from File")
# If you have trouble saving, try using a different directory

#meant to be used with training file index
import time
people=[]
procedure='10-7'
for sheet in all_sheets():
    active_sheet(sheet)
    if sheet!="Index":
        print sheet
        i=3
        while not Cell(i,1).is_empty():
            if sheet=="Jack Olsen":
                print str(Cell(i,1).value),procedure,str(Cell(i,1).value)==procedure
            if str(Cell(i,1).value)==procedure:
                people.append(sheet)

            i+=1

people=set(people)
print people
with open('people.txt','w') as p:
    p.write(procedure+"\n")
    for person in people:
        p.write(person+"\n")

time.sleep(500)

        
