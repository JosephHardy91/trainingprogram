#change training backup based on tfi

import time

tfi_path=r"C:\Users\User\SyncedFolder\Training\QAM R9 work area\TrainingRecords\Training File Indices\Training File Index.xlsm".replace("\\","/")
tfi='Training File Index.xlsm'
print all_wkbks()
open_wkbk(tfi_path)
rname=active_wkbk()
active_wkbk(tfi)
tfi_sheets=all_sheets()

sheetpairs=[]
active_wkbk(rname)
print "Collecting eligible sheets"
for tsheet in tfi_sheets:
    for rsheet in all_sheets():
        if rsheet in tsheet:
            sheetpairs.append((rsheet,tsheet))
sheet_dict={}
active_wkbk(tfi)
print "Collecting TFI data"
for (reportsheet,tfisheet) in sheetpairs:
    active_sheet(tfisheet)
    print tfisheet
    sheetdata=[]
    i=3
    while not Cell(i,1).is_empty():
        print (' '.join([str(Cell(i,1).value),str(Cell(i,2).value)]),True if Cell(i,4).value is not None else False),Cell(i,4).value
        sheetdata.append((' '.join([str(Cell(i,1).value),str(Cell(i,2).value)]),True if Cell(i,4).value is not None else False))
        i+=1

    sheet_dict[reportsheet]=sheetdata

print "Applying TFI data to eligible sheets to find discrepancies"
active_wkbk(rname)
for rsheet in sheet_dict:
    print rsheet
    active_sheet(rsheet)
    data=sheet_dict[rsheet]
    #print data
    procs,vals=zip(*data)
    procs,vals=list(procs),list(vals)
    i=2
    while not Cell(i,1).is_empty():
        if ' '.join([str(Cell(i,1).value),str(Cell(i,2).value)]) in procs:
            ind=procs.index(' '.join([str(Cell(i,1).value),str(Cell(i,2).value)]))
            print rsheet,' '.join([str(Cell(i,1).value),str(Cell(i,2).value)]),procs[ind],vals[ind],vals

            Cell(i,9).value="Yes" if vals[ind]==True else "No"
        i+=1

time.sleep(300)
