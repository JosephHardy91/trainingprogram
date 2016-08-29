import tkFileDialog
import tkMessageBox
from Tkinter import *
import os
import docx
import datetime
import sys
# --------internal imports------------
from features import *
from utilities import *

clearedlist = []
backuplist = []
warninglist = []
traintonewlist = []
trainagainlist = []
updatematrixlist = []


# code for decision parse
def sortinghat():
    print "Sorted in the sorting hat!"


def cleared(varlist):
    v = "%s\t\t%s\t\t\t%s" % (varlist[1], varlist[0], "CLEARED")
    sortinghat()
    clearedlist.append(v)


def backup(varlist):
    v = "%s\t\t%s\t%s\t%s\t\t%s" % (varlist[1], varlist[0], varlist[2], varlist[3], "BACKUP NEEDED")
    sortinghat()
    backuplist.append(v)


def warning(varlist):
    v = "%s\t\t%s\t%s\t%s\t\t%s" % (varlist[1], varlist[0], varlist[2], varlist[3], "EXPIRATION WARNING")
    sortinghat()
    warninglist.append(v)


def traintonew(varlist):
    v = "%s\t\t%s\t%s\t%s\t\t%s" % (varlist[1], varlist[0], varlist[2], varlist[3], "TRAIN TO NEW")
    sortinghat()
    traintonewlist.append(v)


def trainagain(varlist):
    v = "%s\t\t%s\t%s\t%s\t\t%s" % (varlist[1], varlist[0], varlist[2], varlist[3], "RETRAIN")
    sortinghat()
    trainagainlist.append(v)


def updatematrix(varlist):
    v = "%s\t\t%s\t%s\t%s\t\t%s" % (varlist[1], varlist[0], varlist[2], varlist[3], "UPDATE THE MATRIX")
    sortinghat()
    updatematrixlist.append(v)


def writefile(listoflists, name):
    print "Writing the results of the sorting hat..."
    txt = open("C:/Users/User/SyncedFolder/Nuclear Training/Training Program Management(Tree Backup)/Reports/summaries/%s summary.csv" % name,
               "w")
    txt.write("Person\tProcedure\tTraining Date\tExpiration Date\n")
    for list in listoflists:
        for line in list:
            txt.write("%s\n" % line)
        txt.write("\n")
    txt.close()


def decisionparse(person, procedure, TD, match, training, frequency, notes):
    print "Deciding for %s:%s" % (person,procedure)
    if "annual" in frequency:
        years = 1
    elif "three" in frequency:
        years = 3
    elif "5" in frequency:
        years = 5
    else:
        years = 0
    # print TD=='None'
    if TD != 'None' and TD != "TBD":
        datestring, conTD, expiry = convert_to_date(TD, years)
        X = datetime.date.today() > expiry
    else:
        datestring = ""
        conTD = 0
        expiry = 0
        X = False
    varlist = ["'" + procedure, person, TD, expiry, years]
    if "revision" in frequency:
        if years > 0:
            if match == "yes":
                if training == "yes":
                    if X:
                        warning(varlist)
                    else:
                        trainagain(varlist)
                elif training == "no":
                    if X:
                        backup(varlist)
                    else:
                        traintonew(varlist)
                        backup(varlist)
            elif match == "no":
                if training == "yes":
                    if X:
                        traintonew(varlist)
                    else:
                        traintonew(varlist)
                elif training == "no":
                    if X:
                        traintonew(varlist)
                        backup(varlist)
                    else:
                        traintonew(varlist)
                        backup(varlist)
        else:
            if match == "yes":
                if training == "yes":
                    cleared(varlist)
                elif training == "no":
                    backup(varlist)
            elif match == "no":
                if training == "yes":
                    if "update" in notes:
                        updatematrix(varlist)
                    else:
                        traintonew(varlist)
                elif training == "no":
                    traintonew(varlist)
                    backup(varlist)
    elif "indoctrination" in frequency:
        if match == "yes":
            if training == "yes":
                cleared(varlist)
            elif training == "no":
                backup(varlist)
        elif match == "no":
            if training == "yes":
                cleared(varlist)
            elif training == "no":
                backup(varlist)
    elif "annual" in frequency or "5" in frequency:
        if match == "yes":
            if training == "yes":
                if X:
                    warning(varlist)
                else:
                    traintonew(varlist)
            elif training == "no":
                if X:
                    backup(varlist)
                    warning(varlist)
                else:
                    traintonew(varlist)
                    backup(varlist)
        else:
            if training == "yes":
                if X:
                    warning(varlist)
                else:
                    traintonew(varlist)
            elif training == "no":
                if X:
                    backup(varlist)
                    warning(varlist)
                else:
                    traintonew(varlist)
                    backup(varlist)
    else:
        if training == "no":
            backup(varlist)
        if match == "no":
            traintonew(varlist)
