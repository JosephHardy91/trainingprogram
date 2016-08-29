#import and define help
def color_workbook():
	import os
	#create list of relevant workbooks and their paths
	def dir():
		list=[]
		for sub,dir,files in os.walk(os.getcwd()):
			for file in files:
				if "Training Matrix" in file and "Historic" not in sub:
					list.append([sub+"\\"+file,file])
		return list
	#color matrices
	def clsrrs():
		rows=0
		l0=1
		while not "Review" in str(Cell(l0,1).value):
			#Cell(l0,15).value=rows
			if "Reference" in str(Cell(l0,1).value):
				startrow=l0
				collength=1
				rows=0
				if Cell(startrow,3).is_empty():
					c=0
					while not Cell(startrow,collength).is_empty() or c==1:
						if c==1 and Cell(startrow,collength+1).is_empty():
							break
						elif Cell(startrow,collength+1).is_empty():
							c=1
							collength+=1
							continue
						collength+=1
				else:
					while not Cell(startrow,collength).is_empty():
						collength+=1
					collength=collength-1
			l0+=1
			rows+=1
		rows=rows-2
		return collength,rows,startrow
	#color top row
	def colortop(startrow,collength):
		for column in range(1,collength+1):
			Cell(startrow,column).color="b7dee8"
	#bold top row
	def boldtop(collength,startrow,wkbk):
		if collength==7 and startrow==5:
			VBA("PERSONAL.XLSB!boldtitletoG")
			pass
		elif collength==8 and startrow==5:
			VBA("PERSONAL.XLSB!boldtitletoH")
			pass
		else:
			txt=open("missedbolds.txt","a")
			txt.write("for workbook %s the top was not bolded"%wkbk)
			txt.close()
	#set up rows that need to be colored
	def createrowlist(startrow,rows):
		start=0
		addgo=0
		rowlist=[]
		for row in range(startrow,startrow+rows+1):
			if row>=(startrow+2) and Cell(row,1).is_empty() and start==0:
				if not Cell(row+1,1).is_empty():
					start=1
					addgo=1
			elif row>=(startrow+2) and not Cell(row,1).is_empty() and start==0:
				start=1
				addgo=1
				rowlist.append(row)
				if not Cell(row+1,1).is_empty():
					addgo=0
			elif start==1 and addgo==1:
				rowlist.append(row)
				if not Cell(row+1,1).is_empty():
					addgo=0
			elif addgo==0:
				if not Cell(row+1,1).is_empty():
					addgo=1
		return rowlist
	#Cell(1,13).value=rowlist
	def colorrows(rowlist,startrow,rows,collength):
		for row in range(startrow,startrow+rows+1):
			for column in range(1,collength+1):
				if not Cell(row,column).color=="yellow" and row in rowlist:
					Cell(row,column).color="fde9d9"
				elif not Cell(row,column).color=="yellow" and startrow!=row:
					Cell(row,column).color="white"
	#Cell(1,14).value=collength
	def main(sub,file,suppress):
		if suppress==0:
			open_wkbk(sub)
		active_wkbk(file)
		collength,rows,startrow=clsrrs()
		rowlist=createrowlist(startrow,rows)
		colortop(startrow,collength)
		#boldtop(collength,startrow,file)
		colorrows(rowlist,startrow,rows,collength)
		if suppress==0:
			save()
			close_wkbk(file)
	#MAIN
	suppress=raw_input("Supress directory?")
	if suppress.lower()=="yes":
		suppress=1
	else:
		suppress=0
	print suppress
	if suppress==0:
		directory=dir()
		for sheet in all_sheets():
			if suppress==0:
				sub=set[0]
				file=set[1]
				main(sub,file,suppress)
	elif suppress!=0:
		file=active_wkbk()
		main("",file,suppress)
			
