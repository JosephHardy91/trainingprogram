#pseudo-code for decision parse
def cleared(varlist):
	v="%s\t%s\t"%(varlist[1],varlist[0])
	clearedlist.append(v)

def backup(varlist):
	v="%s\t%s\t%s\t%s\t%s"%(varlist[1],varlist[0],varlist[2],varlist[3])
	backuplist.append(v)
	
def warning(varlist):
	v="%s\t%s\t%s\t%s\t%s"%(varlist[1],varlist[0],varlist[2],varlist[3])
	warninglist.append(v)

def traintonew(varlist):
	v="%s\t%s\t%s\t%s\t%s"%(varlist[1],varlist[0],varlist[2],varlist[3])
	traintonew.append(v)

def trainagain(varlist):
	v="%s\t%s\t%s\t%s\t%s"%(varlist[1],varlist[0],varlist[2],varlist[3])
	trainagainlist.append(v)
	
def updatematrix(varlist):
	v="%s\t%s\t%s\t%s\t%s"%(varlist[1],varlist[0],varlist[2],varlist[3])
	updatematrixlist.append(v)
def writefile(listoflists):
	txt=open("formsummary.txt","w")
	txt.write("Person\tProcedure\tTraining Date\tExpiration Date\n")
	for list in listoflists:
		for line in list:
			txt.write("%s\n"%line)
	txt.close()
def decisionparse(person,procedure,TD,expiry,match,training,frequency):
	if "annual" in frequency:
		years=1
	elif "three" in frequency:
		years=3
	elif "5" in frequency:
		years=5
	else:
		years=0
	expiry=TD+years
	X=(TD+years)>expiry
	varlist=[procedure,person,TD,expiry,years]
	if "revision" in frequency:
		if years>0:
			if match=="yes":
				if training=="yes":
					if X:
						warning(varlist)
					else:
						trainagain(varlist)
				elif training=="no":
					if X:
						backup(varlist)
					else:
						traintonew(varlist)
						backup(varlist)
			elif match=="no":
				if training=="yes":
					if X:
						traintonew(varlist)
					else:
						traintonew(varlist)
				elif training=="no":
					if X:
						traintonew(varlist)
						needbackup(varlist)
					else:
						traintonew(varlist)
						needbackup(varlist)
		else:
			if match=="yes":
				if training=="yes":
					cleared(varlist)
				elif training=="no":
					backup(varlist)
			elif match=="no":
				if training=="yes":
					if "update" in notes:
						updatematrix(varlist)
					else:
						traintonew(varlist)
				elif training=="no":
					traintonew(varlist)
					backup(varlist)
	elif "indoctrination" in frequency:
		if match=="yes":
			if training=="yes":
				cleared(varlist)
			elif training=="no":
				backup(varlist)
		elif match=="no":
			if training=="yes":
				cleared(varlist)
			elif training=="no":
				backup(varlist)
	elif "annual" in frequency or "5" in frequency:
		if match=="yes":
			if training=="yes":
				if X:
					warning(varlist)
				else:
					traintonew(varlist)
			elif training=="no":
				if X:
					backup(varlist)
					warning(varlist)
				else:
					traintonew(varlist)
					backup(varlist)
		else:
			if training=="yes":
				if X:
					warning(varlist)
				else:
					traintonew(varlist)
			elif training=="no":
				if X:
					backup(varlist)
					warning(varlist)
				else:
					traintonew(varlist)
					backup(varlist)
	else:
		if training=="no":
			backup(varlist)
		if match=="no"
			traintonew(varlist)
