'''
Deepika Rajwar
1901EE19
'''
import os
import csv
import openpyxl

def generate_marksheet():

	#mapped roll no with names
	Dict_from_f1={}
	with open("C:\\Users\names-roll.csv",'r') as f1:
		reader=csv.reader(f1)
		Dict_from_f1={int(row[0]):row[1] for row in reader}
	f1.close()

	roll_list=[] 	#stores all roll no
	with open("C:\\Users\grade.csv") as f:
		csv_file=csv.DictReader(f)
		for row in csv_file:
			roll=row["Roll"]
			roll_list.append(roll)
			curr_sem="Sem"+row["Sem"]
			sub_code=row["SubCode"]
			filename=".\output"+roll+".xlsx"
			try:
				wb=openpyxl.load_workbook(filename)
			except:
				wb=openpyxl.Workbook()
			try:
				sheet=wb.get_sheet_by_name(curr_sem)			
			except:
				sheet=wb.create_sheet(curr_sem)
				sheet.append(["S.No","Subject No","Subject Name","L-T-P","Credit","Subject Type","Grade"])
			list=[]
			with open("subjects master.csv") as f2:
				for i,r in enumerate(f2):
					if i["subno"]==sub_code:
						list=r
			f2.close()
			list=[1]+list+row["Sub_Type"]+row["Grade"]
			sheet.append(list)	
			wb.save(filename)
	f.close()
	Grading={'AA':10,'AB':9,'BB':8,'BC':7,'CC':6,'CD':5,'DD':4,'F':0,'I':0}
	roll_list=list(set(roll_list))		#using set to remove duplicated rollno from list
	for filename in roll_list:
		file_name=".\output"+filename
		wb=openpyxl.load_workbook(file_name)
		sheet=wb.active()
		sheet['A1']="Roll No"
		sheet['B1']=roll
		sheet['A2']="Name of Student"
		sheet['B2']=Dict_from_f1[roll]
		sheet['A3']="Discipline"
		sheet['B3']=roll[3:2]
		sheet['A4']="Semester No"
		c=1
		for i in sheet['B4','I4']:
			i.value=c;
			c+=1
		sheet['A5']="Semester wise Credit Taken"
		sheet['A6']="Spi"
		sheet['A7']="Cpi"
		sheet.title="Overall"
		#Fixing the S.No in semester sheets
		total_credit=0
		total_grade=0
		credit=[]
		cpi=[]
		spi=[]
		for sheet in wb.sheetnames():
			c=1
			credit_sem=0
			grade_sem=0
			mr=sheet.maxrow
			for i in range(2,mr+1):
				sheet.cell(row=i,column=1).value=c		#Fixing Serial No here
				credit_sem+=sheet.cell(row=i,column=5).value
				grade_sem+=Grading[sheet.cell(row=i,column=7).value]
				c+=1
			credit.append(credit_sem)
			spi.append(float("{:.2f}".format(grade_sem/sum)))
			total_credit+=credit_sem
			total_grade+=grade_sem
			cpi.append(float("{:.2f}".format(total_grade/total_credit)))
		sheet=wb['Overall']
		j=0;
		for i in range(2,10):
			sheet.cell(row=5,column=i).value=credit[j];		
			sheet.cell(row=6,column=i).value=spi[j];
			sheet.cell(row=7,column=i).value=cpi[j];
			j+=1
		wb.save(filename)
	return

generate_marksheet()


