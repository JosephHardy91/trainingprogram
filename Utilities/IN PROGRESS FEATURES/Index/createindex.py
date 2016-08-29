# -*- coding: utf-8 -*-
# Welcome to the DataNitro Editor
# Use Cell(row,column).value, or Cell(name).value, to read or write to cells
# Cell(1,1) and Cell("A1") refer to the top-left cell in the spreadsheet
# 
# Note: To run this file, save it and run it from Excel (click "Run from File")
# If you have trouble saving, try using a different directory
active_wkbk("bnpptraining.xlsx")
active_sheet("Version of Procedures")
#active_wkbk()
l1=2
trainingdict={}
while not Cell(l1,1).is_empty():
	val=Cell(l1,1).value
	if "QAM" in Cell(l1,1).value:
		try:
			val=str(Cell(l1,1).value).split(" ")[1]
		except:
			val="QAM"
	trainingdict[val]=[Cell(l1,2).value,Cell(l1,4).value]
	print val
	l1+=1

persondict={}

active_wkbk("auditors.xlsx")
for sheet in all_sheets():
	personlist=[]
	l1=2
	while not Cell(l1,1).is_empty():
		reference=Cell(l1,1).value
		if "QAM" in Cell(l1,1).value:
			reftype="QAM"
			try:
				reference=str(Cell(l1,1).value).split(" ")[1]
			except:
				reference="QAM"
		elif "SI" in str(Cell(l1,1).value):
			reftype="QAP-SI"
		elif "-" in Cell(l1,1).value:
			reftype="QAP"
		else:
			reftype="Misc"
		try:
			name=trainingdict[reference][1]
			currentrev=trainingdict[reference][0]
		except:
			name=""
			currentrev=""
		trainedrev=Cell(l1,2).value
		personlist.append([reftype,"'"+reference,name,currentrev,trainedrev])
		l1+=1
	persondict[sheet]=personlist
        
        
labels=["Type","Reference","Name","Current Rev","Trained to Rev"]
nb=new_wkbk()
active_wkbk(nb)

for person in persondict:
    new_sheet(person)
    active_sheet(person)
    l1=1
    c=1
    for entry in persondict[person]:
        c=1
        if l1==1:
            while c<len(labels):
                Cell(l1,c).value=labels[c-1]
                c+=1
            c=1
        else:
            while c<len(entry):
                Cell(l1,c).value=entry[c-1]
                c+=1
            c=1
        l1+=1
    autofit(person)
            
