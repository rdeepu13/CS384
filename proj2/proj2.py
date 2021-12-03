import os
import csv
from datetime import datetime
import numpy as np
import pandas as pd
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog
from tkinter import *
from html2image import Html2Image
import cv

hti=Html2Image()	#constructor

window=tk.Tk()
window.geometry("1000x200")
window.config(background = "white")
window.title("GUI Based Transcript Generator")	#to define the title
entry1=tk.StringVar()
entry2=tk.StringVar()
subject_data=pd.read_csv(r"E:/1901EE19_2021/proj2/sample_input/subjects_master.csv",header=0,delimiter=',')	#edit
grade_data=pd.read_csv(r"E:/1901EE19_2021/proj2/sample_input/grades.csv")	#edit
path='E:/1901EE19_2021/proj2/transcriptsIITP'	#edit
with open("E:/1901EE19_2021/proj2/sample_input/names-roll.csv") as f:	#edit
	reader=csv.reader(f)
	get_name={row[0]:row[1] for row in reader}	#get_name maps rollno with name 
f.close()


class PDF(FPDF):
	pass

def generate_transcript(roll):
	roll.strip()
	try:
		Name = get_name[roll]
	except:
		print("This roll no. doesnt exist",roll)
		return
	Year_of_Adm = "20"+roll[0:2]
	if roll[2:4]=="01":
		total_sem=8
		Programme="Bachelor of Technology"
		page_format='A3'
		pdf_w=420
		pdf_h=297
	else:
		total_sem=4
		page_format='A4'
		pdf_w=297
		pdf_h=210
		if roll[2:4]=="11":
			Programme="Master of Technology"
		elif roll[2:4]=="12":
			Programme="M.Sc."
		else:
			Programme="P.hd"
	
	file=roll+".pdf"
	filename=os.path.join(path,file)
	dict={'CS':'Computer Science and Engineering','EE':'Electrical Engineering','CE':'Civil Engineering','ME:':'Mechanical Engineering','CB':'Chemical Engineering'}
	Course=roll[4:6]
	try:
		course=dict[Course]
	except:
		return
	pdf=PDF()
	pdf=PDF(orientation='L',unit='mm',format=page_format)
	P=open(filename,'a')
	pdf.add_page()
	pdf.set_margins(0,0,0)
	pdf.set_font('Arial','U',12)
	pdf.rect(10.0,10.0,pdf_w-20.0,pdf_h-20.0)
	pdf.image('E:/1901EE19_2021/proj2/iitp_logo.png',11.0,11.0,370,30)
	pdf.line(x1=10.0,y1=40.0,x2=pdf_w-10.0,y2=40.0)
	pdf.text(185.0,39.0,"TRANSCRIPT")

	date_generated=datetime.now().strftime("%d %b %Y %H:%M")
	line1="RollNo: "+roll+" Name: "+Name+" Year of Admission:"+Year_of_Adm
	line2="Programme:"+Programme+" Branch: "+course
	last_line="Date Generated: "+str(date_generated)

	pdf.text(130.0,45.0,line1)
	pdf.text(130.0,50.0,line2)
	pdf.text(20.0,pdf_h-12.0,last_line)	#Shows Date Generated
	pdf.rect(126.0,41.0,160.0,11.0)
	pdf.line(10.0,122.0,pdf_w-10.0,122.0)
	pdf.line(10.0,195.0,pdf_w-10.0,195.0)

	Grading={'AA':10,'AB':9,'BB':8,'BC':7,'CC':6,'CD':5,'DD':4,'DD*':4,'F*':0,'F':0,'I':0}
	total_credit=0
	grd_sum=0
	cpi=0.00
	for curr_sem in range(1,total_sem+1):
		sem_credit=0
		spi=0.00
		sum=0
		clr=0
		get_roll_data=grade_data.loc[((grade_data['Roll']==roll) & (grade_data['Sem']==curr_sem)),['Roll','SubCode','Grade']]	#iterate curr_sem to1-8 if btech 1-4 if mtech
		sublist=get_roll_data['SubCode'].tolist()
		Grd_list=get_roll_data['Grade'].tolist()
		df=pd.DataFrame()
		i=0
		for sub in sublist:
			df1=subject_data.loc[(subject_data['subno']==sub),['subno','subname','ltp','crd']]
			df=pd.concat([df,df1])
			sub_credit=int(df1['crd'].values)
			sem_credit+=sub_credit
			sum+=sub_credit*int(Grading[Grd_list[i].strip()])
			if Grading[Grd_list[i].strip()]<4: 
				clr+=sub_credit
			i+=1
		total_credit+=sem_credit
		grd_sum+=sum
		try:	#Zero Division Error
			spi=float("{:.2f}".format(sum/sem_credit))
			cpi=float("{:.2f}".format(grd_sum/total_credit))
		except:
			pass
		sheet=pd.DataFrame()
		try:
			sheet=get_roll_data.set_index('SubCode').join( df.set_index('subno')) 
		except:
			pass
		column=['subname','ltp','crd','Grade']
		sheet=sheet.reindex(columns=column)
		sheet=sheet.rename(columns={'subname':'Subject Name','ltp':'L-T-P','crd':'CRD','Grade':'GRD'})
		#print(sheet.head) 

		Sem="Semester"+str(curr_sem)
		img_f="Semester"+str(curr_sem)+".png"
		img_file=os.path.join(path,img_f)

		hti=Html2Image(size=(500,265))
		html=sheet.to_html()
		x=open("E:/index.html",'w')
		x.write(html)
		x.close()

		hti.screenshot(html_file="index.html",save_as=img_f)	#conversion to img file
		line_sem="Credits Taken: "+str(sem_credit)+" Credits Cleared: "+str((sem_credit-clr))+" SPI: "+str(spi)+" CPI: "+str(cpi)

		pdf.text(20.0+((curr_sem-1)%3)*130,58.0+(int((curr_sem-1)/3))*75,Sem)	#For printing the current semester
		pdf.image(img_f,18.0+((curr_sem-1)%3)*130,59.0+(int((curr_sem-1)/3))*75,96,53)	#prints dataframe of particular semester to image in pdf
		pdf.rect(18.0+((curr_sem-1)%3)*130,37.0+(int((curr_sem-1)/3)+1)*75,115.0,6.0)	#makes grid for semester detail
		pdf.text(18.0+((curr_sem-1)%3)*130,42.0+(int((curr_sem-1)/3)+1)*75,line_sem)	#prints details

	pdf.output(filename,'F')
	P.close()
	return

def generate1():	#Function for option 1
	R1=entry1.get()
	R2=entry2.get()
	first_s=R1[6:8]
	rest=R1[0:6]
	last_s=R2[6:8]
	first=int(first_s)
	last=int(last_s)
	for i in range(first,last+1):
		if len(str(i))==1:
			i="0"+str(i)
		for_roll=rest+str(i)
		for_roll=for_roll.upper()
		generate_transcript(for_roll)

def	generate2():	#Function for Option 2
	all_roll=get_name.keys()
	for roll in all_roll:
		generate_transcript(roll)
	
def browseFiles():
	filename1=filedialog.askopenfile(initialdir='/',title="Select a file",filetypes=[("*all files","*.*")])
	label_file_explorer.configure(text="File Opened: "+str(filename1))

def browseFiles2():
	filename2=filedialog.askopenfile(initialdir='/',title="Select a file",filetypes=[("*all files","*.*")])
	label_file_explorer2.configure(text="File Opened: "+str(filename2))

entry11=tk.Entry(window,textvariable=entry1,font=('calibre',10,'normal')) #input lower limit
entry22=tk.Entry(window,textvariable=entry2,font=('calibre',10,'normal'))	#input upper limit
entry11.grid(row=1,column=2)	#setting location for getting lower limit
entry22.grid(row=1,column=3)	#setting location for getting upper limit
btn1=tk.Button(window,text="Generate Transcripts for a given range",command=lambda:generate1()).grid(row=1,column=1) #button for option1
btn2=tk.Button(window,text="Generate all Transcripts",command=lambda:generate2()).grid(row=2,column=1)	#button for option2
label_file_explorer=Label(window,text="Upload Seal",width=100,height=2,fg='blue')
label_file_explorer2=Label(window,text="Upload Signature",width=100,height=2,fg='blue')
button_explore=tk.Button(window,text="Browse File",command=browseFiles)
button_explore2=tk.Button(window,text="Browse File",command=browseFiles2)
label_file_explorer.grid(row=3,column=1)
label_file_explorer2.grid(row=4,column=1)
button_explore.grid(row=3,column=2)	#location for taking seal
button_explore2.grid(row=4,column=2)	#location for taking signature
Quit_btn=tk.Button(window,text='Quit',command=window.quit).grid(row=5,column=1)	#Quit button
window.mainloop()