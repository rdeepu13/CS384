import os
def output_by_subject():
	header=["rollno","register_sem","subno","subtype"]
	dir="./output_by_subject//"
	with open("regtable_old.csv","r") as f:
		for line in f:
			result=[]
			row=f.readlines()
			result.append(row[0])
			result.append(row[1])
			result.append(row[3])
			result.append(row[8])
			print(result)
			if result[0]=="rollno":
				continue
			filename="%s.csv"% result[2]
			path=os.path.join(dir,filename)
			os.mkdir(path)
			if os.path.isfile(filename):
				g=open(filename,"a")
				g.write(result)
				g.close()
			else:
				g=open(filename,"a")
				g.write(header)
				g.write(result)
				g.close()
				
	f.close()
	return

def output_individual_roll():
	header=["rollno","register_sem","subno","subtype"]
	dir="./output_individual_roll/"
	with open("regtable_old.csv","r") as f:
		for line in f:
			result=[]
			row=f.readlines()
			result.append(row[0])
			result.append(row[1])
			result.append(row[3])
			result.append(row[8])
			if result[0]=="rollno":
				continue
			filename="%s.csv"% result[0]
			path=os.path.join(dir,filename)
			os.mkdir(path)
			if os.path.isfile(filename):
				g=open(filename,"a")
				g.write(result)
				g.close()
			else:
				g=open(filename,"a")
				g.write(header)
				g.write(result)
				g.close()
				
	f.close()
	return 
