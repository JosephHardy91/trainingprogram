import tkFileDialog
import tkMessageBox
from Tkinter import *
import os
import docx
import datetime
import sys
import zipfile
#--------internal imports------------
from features import *
from decisionparse import *
from cell import *
import datetime
import itertools
#utilities.py
class suppress_stdout_stderr(object):
    '''
    A context manager for doing a "deep suppression" of stdout and stderr in
    Python, i.e. will suppress all print, even if the print originates in a
    compiled C/Fortran sub-function.
       This will not suppress raised exceptions, since exceptions are printed
    to stderr just before a script exits, and after the context manager has
    exited (at least, I think that is why it lets exceptions through).

    '''
    def __init__(self):
        # Open a pair of null files
        self.null_fds =  [os.open(os.devnull,os.O_RDWR) for x in range(2)]
        # Save the actual stdout (1) and stderr (2) file descriptors.
        self.save_fds = (os.dup(1), os.dup(2))

    def __enter__(self):
        # Assign the null pointers to stdout and stderr.
        os.dup2(self.null_fds[0],1)
        os.dup2(self.null_fds[1],2)

    def __exit__(self, *_):
        # Re-assign the real stdout/stderr back to (1) and (2)
        os.dup2(self.save_fds[0],1)
        os.dup2(self.save_fds[1],2)
        # Close the null files
        os.close(self.null_fds[0])
        os.close(self.null_fds[1])
def all_cases(string):
    '''creates a list of all possible case mixes of a string'''
    return map(''.join,itertools.product(*((c.upper(),c.lower()) for c in string)))
def find_self_directory(extra=None):
    #dir1=os.path.expanduser('~')
    dir1=os.path.realpath(__file__)
    while True:
        dir1=os.path.split(dir1)
        if "Training Program Management" in dir1[0]:
            dir1='\\'.join(dir1)[:-1]
        elif "Training Program Management" in dir1[1]:
            dir1='\\'.join(dir1)
            break
    #print dir1
    #homedir=os.path.splitdrive(dir1)
    for sub,dir,files in os.walk(dir1):
        if extra is not None:
            #print sub
            if "Training Program Management" in sub and any(word in sub for word in all_cases(extra)):
                return sub
        elif extra is None:
            if "Training Program Management" in sub:
                return sub
    return None
def matrixtoname(wkbk_name):
    name=wkbk_name.split(" ")
    return wkbk_name[:2]
def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))
def proper_date():
    d=datetime.datetime.now()
    day,month,year=d.day,d.month,d.year
    if len(str(day))==1:
        day="0"+str(day)
    else:
        day=str(day)
    if len(str(month))==1:
        month="0"+str(month)
    else:
        month=str(month)
    year=str(year)[2:]
    date=month+day+year
    return date
def pop():
    pop_up_message('Success','Ok')
def parseprocname(string):
    t=0
    number=""
    rev=""
    name=""
    for x in string:
        if t==0 and x==" ":
            t=1
        elif t==1 and x==" ":
            t=2
        elif t==2 and x==" ":
            t=3
        elif t==1:
            number+=x
        elif t==2:
            rev+=x
        elif t==3:
            name+=x
    return number,int(rev[1:]),parsename(name)
def tdpstack(directory,section):
    tdptxt=open("tdpstack.txt","w")
    tdptxt.write("tdpddic=%s"%directory)
    tdptxt.close()
    names=directory+"/"+section+"/namedictionary.txt"
    positions=directory+"/"+section+"/positiondict.txt"
    dir=directory+"/"+section+"/"
    if section=="forms":
        dir=directory+"/"+section+"/forms"
        print "Select report file"
        reportfile=specify_file('reports')
        training_form=directory+"/"+section+"/Exhibit QAM 2.4 Process and Procedure Training Form.docx"
        bnpp_file=directory+"/"+"general"+"/bnpptraining.xlsx"
        #print bnpp_file
        return names,positions,dir,reportfile,training_form,bnpp_file
    elif section=="report":
        bnpp_file=directory+"/"+"general/"
        return bnpp_file
def convert_to_date(datestring,years):
    dates=datestring.split("/")
    imonth,iday,iyear=[int(x) for x in dates]
    conTD=datetime.date(iyear,imonth,iday)
    year=years+iyear
    expiry=datetime.date(year,imonth,iday)
    return datestring,conTD,expiry
def parsename(name):
    if "xlsx" in name:
        cut=5
    elif "xls" or "pdf" in name:
        cut=4
    new=""
    t=0
    j=0
    for x in range(0,len(name)):
        if name[x]=="-":
            t=1
        if t==1 and name[x]==" ":
            t=2
            print name[x+1]
        elif t==2 and name[x]==" ":
            #print name[x-1]
            j=x+1
            #print name[j]
            break
    return name[j:-cut]
def pathic(path2):
    i=0
    for x in path2:
        if x=="/":
            i=0
        i+=1
    i-=1
    return path2[-i:]
def pop_up_message(label,message):
    tops=Tk()
    top=Frame(tops)
    top.pack()
    def box(label,message):
        label=label
        message=message
        tkMessageBox.showinfo(label,message)
    B1=Button(top,text=label)
    B1.pack()
def closewindow():
    B1.destroy()
def specify_file(extra=None):
    root2=Tk()
    root2.withdraw()
    homedir=find_self_directory(extra)
    print "Default folder:\n"+homedir+"\n"
    return tkFileDialog.askopenfilename(initialdir=homedir)
def file_parse(extra=None):
    file = specify_file(extra)
    return file, pathic(file)
def specify_directory(extra=None):
    root2 = Tk()
    root2.withdraw()
    homedir = find_self_directory(extra)
    print "Default folder:\n" + homedir + "\n"
    return tkFileDialog.askdirectory()
def commenu():
    # def menus():
    # no_selection_is_made=1
    # choices="""
    # 1 Create report from directory(specify directory)
    # 2 Create forms from report(specify report file)
    # 3 Recolor matrices(specify directory)
    # 4 Update BNPP Training(specify Training file and directory to update from)
    # """
    # while no_selection_is_made:
    # print "Select a choice."
    # try:
    # choice=int(raw_input(choices))
    # except:
    # print "invalid choice"
    # if choice==1:
    # return create_report()
    # elif choice==2:
    # return create_forms()
    # elif choice==3:
    # return recolor_matrices()
    # elif choice==4:
    # return bnpp_update()
    # else:
    # print "invalid choice"
    return "nothing"
def addformulas(all_sheets, save_path):
    try:
        tdptxt = open("tdpstack.txt", "r")
        for line in tdptxt:
            if "tdpddic" in line:
                tdpddic = line.split("=")[1]
                break
        tdptxt.close()
        bnpp_file = tdpstack(tdpddic, "report")
    except:
        print "Specify tdpstack directory"
        bnpp_file = tdpstack(specify_directory(), "report")
    for sheet in all_sheets:
        active_sheet(sheet)
        Cell(sheet, 1, 6).value = "Current Rev"
        Cell(sheet, 1, 7).value = "Current Date"
        Cell(sheet, 1, 8).value = "Match"
        Cell(sheet, 1, 9).value = "Training backup?"
        Cell(sheet, 1, 10).value = "Notes"
        l2 = 2
        while not Cell(l2, 1).is_empty():
            Cell(sheet, l2,
                 6).value = "=VLOOKUP(A%d,'%s[bnpptraining.xlsx]Version of Procedures'!$A2:$C$150,2,FALSE)" % (
            l2, bnpp_file)
            Cell(sheet, l2,
                 7).value = "=VLOOKUP(A%d,'%s[bnpptraining.xlsx]Version of Procedures'!$A2:$C$150,3,FALSE)" % (
            l2, bnpp_file)
            Cell(sheet, l2, 8).value = '=IF(AND(VALUE(MID(B%d,2,2))=F%d,C%d=G%d,NOT(ISBLANK(B%d))),"Yes","No")' % (
            l2, l2, l2, l2, l2)
            Cell(sheet, l2, 10).value = '=IF(ISERROR(H%d),"Need to locate doc","")' % l2
            l2 += 1
        l0 = 2
        while not Cell(l0, 1).is_empty():
            nam = Cell(l0, 1).value
            if isinstance(nam, unicode) or isinstance(nam, basestring):
                while nam[0] == " " or nam[0] == "'" or nam[len(nam) - 1] == " " or (("QAM 22" in nam or "QAM 21" in nam) and (nam != "QAM 22" and nam != "QAM 21")):
                    if nam[0] == " " or nam[0] == "'":
                        nam = nam[1:len(nam)]
                    if nam[len(nam) - 1] == " ":
                        nam = nam[:len(nam) - 1]
                    if "QAM 22" in nam and nam != u'%s' % "QAM 22":
                        nam = u'%s' % "QAM 22"
                    elif "QAM 21" in nam and nam != u'%s' % "QAM 21":
                        nam = u'%s' % "QAM 21"
                if "SI" in nam:
                    for x in range(0, len(nam)):
                        if x + 1 != len(nam) and nam[x] + nam[x + 1] == "SI" and nam[x - 1] != "-":
                            nam = nam[:x - 1] + "-" + nam[x:]
                Cell(l0, 1).value = u'%s' % nam
            l0 += 1
        autofit(sheet)
    with suppress_stdout_stderr():
        save(save_path)
    print "Formulas added. \nPlease open Report to enable content before making it backwards compatible. Thanks!\n"
