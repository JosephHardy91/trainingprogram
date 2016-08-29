# -*- coding: utf-8 -*-
# Create program with options:
# Create report from directory(specify directory)
# Create forms from report(specify report file)
# Recolor matrices(specify directory)
# Update BNPP Training(specify Training file and directory to update from)
#-------external imports-------
from Tkinter import *
import sys,os
sys.path.insert(0, os.path.join(os.path.abspath('.'),"modules"))
#--------internal imports------------
from features import *
from utilities import *
from decisionparse import*
#-------handy notes----------
#tkFileDialog.askopenfile()
#tkFileDialog.askdirectory()
#root=Tkinter.Tk()
#root.withdraw()


root=Tk()
root.wm_title("Training Management")
root.geometry('300x95')
menu=Menu(root)
root.config(menu=menu)
filemenu=Menu(menu)
menu.add_command(label="Create reports from directory",command=create_report)
#menu.add_command(label="Create one report from file",command=create_report_from_f)
menu.add_command(label="Establish report backwards compability",command=backwards_compatible)
menu.add_command(label="Create forms from report",command=create_forms)
menu.add_command(label="Recolor matrices by directory",command=recolor_matrices)
menu.add_command(label="Recolor one matrix from file",command=recolor_matrices_from_f)
menu.add_command(label="Update Procedure Index",command=bnpp_update)
menu.add_command(label="Create File Index",command=create_file_index)

mainloop()

