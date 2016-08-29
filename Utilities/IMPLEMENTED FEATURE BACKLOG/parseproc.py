e="QAP 9-1-SI R2 Control of Processes.pdf"
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