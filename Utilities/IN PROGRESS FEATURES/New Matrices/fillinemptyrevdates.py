#update rev number
import datetime
import time
procdict={}
qa=active_wkbk()
bnpptraining=["C:/Users/Administrator/Desktop/Training Program Management/tdpstack/general/bnpptraining.xlsx","bnpptraining.xlsx","Version of Procedures"]
open_wkbk(bnpptraining[0])
quit=1
while quit!=0:
        if Cell(5,3).value=="Rev Number":
                x=3
        elif Cell(5,4).value=="Rev Number":
                x=4
        else:
                print "no good"
                time.sleep(10)
                quit=1
                break
        active_wkbk(bnpptraining[1])
        l1=2
        active_sheet(bnpptraining[2])
        while not Cell(l1,1).is_empty():
                proclist=[]
                proclist.append("R"+str(Cell(l1,2).value))
                proclist.append((Cell(l1,3).value).strftime('%m/%d/%Y'))
                procdict[str(Cell(l1,1).value)]=proclist
                l1+=1

        active_wkbk(qa)
        l1=5
        while str(Cell(l1,1).value).lower()!="review date":
                truthiness=1
                while truthiness:
                        if str(Cell(l1,1).value)[len(str(Cell(l1,1).value))-1]==" ":
                                Cell(l1,1).value=str(Cell(l1,1).value)[:len(str(Cell(l1,1).value))-1]
                                #print Cell(l1,1).value+"x"
                        else:
                                truthiness=0
                if not Cell(l1,1).is_empty():
                        if "QAM" in Cell(l1,1).value:
                                proc=str(Cell(l1,1).value)
                        elif "QAP" in Cell(l1,1).value:
                                proc=str(Cell(l1,1).value)[4:]
                        else:
                                proc=None
                if not Cell(l1,x).is_empty() and not Cell(l1,1).is_empty():
                        if len(str(Cell(l1,x).value).split(" "))!=2:
                                currev=str(Cell(l1,x).value)
                                print currev,Cell(l1,1).value
                                try:
                                        if currev in procdict[proc]:
                                                rev=currev+" "+procdict[proc][1]
                                                Cell(l1,x).value=rev
                                except:
                                        pass
                l1+=1

        quit=0
close_wkbk(bnpptraining[1])
time.sleep(100)
