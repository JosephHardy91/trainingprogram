#pull possibilities for frequency required from reports
path="C:\Users\Administrator\Desktop\Training Program Management\Reports"
def pullfreqtypes(path):
	freqlist=[]
	done=[]
	for sub,dir,files in os.walk(path):
		for file in files:
			if ".xls" in file and "backup" not in sub and file not in done:
				print sub+"\\"+file
				open_wkbk(sub+"\\"+file)
				active_wkbk(file)
				for sheet in all_sheets():
					active_sheet(sheet)
					l1=1
					while not Cell(l1,1).is_empty():	
						if str(Cell(l1,5).value).lower() not in freqlist:
							val=str(Cell(l1,5).value).lower()
							while True:
								if val[0]==" ":
									val=val[1:]
								elif val[len(val)-1]==" ":
									val=val[:len(val)-1]
								else:
									break
						if val not in freqlist:
							freqlist.append(val)
						l1+=1
				save()
				close_wkbk(file)
				done.append(file)
				print freqlist
	return freqlist
freqlist=pullfreqtypes(path)
txt=open("freqlist.txt","w")
for type in freqlist:
	txt.write("%s\n"%type)

txt.close()