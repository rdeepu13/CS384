#Deepika Rajwar
#1901EE19

import os
import csv
import numpy as np
import pandas as pd
import openpyxl

def feedback_not_submitted():
	
	ltp_mapping_feedback_type = {1: 'lecture', 2: 'tutorial', 3:'practical'}
	output_file_name = ".\course_feedback_remaining.xlsx" 
	
	feedback_data=pd.read_csv(r".\course_feedback_submitted_by_students.csv")
	subject_data=pd.read_csv(r".\course_master_dont_open_in_excel.csv")
	student_data=pd.read_csv(r".\studentinfo.csv")

	Remaining_student_data=pd.DataFrame()	#Empty DataFrame

	with open(r".\course_registered_by_all_students.csv",'r') as F:
		registration_data=csv.DictReader(F)
		for entry in registration_data:
			Roll=entry["rollno"]
			subcode=entry["subno"]

			df=pd.DataFrame(entry,index=[0])
			df1=df[["rollno","register_sem","schedule_sem","subno"]]
			ltp_df=subject_data.loc[subject_data["subno"]==subcode]
			ltp=ltp_df[["ltp"]].values
			val=str(ltp)
			feedback=3-val.count('0')

			if val.count('0')==3: #for 0-0-0
				continue
			get_feedback=feedback_data.loc[( (feedback_data["stud_roll"]==Roll) & (feedback_data["course_code"]==subcode) )]
			#print(get_feedback)
			if (get_feedback.empty | feedback>len(get_feedback.index)):
				get_particular_data=student_data.loc[(student_data["Roll No"]==Roll),["Roll No","Name","email","aemail","contact"]]
				df2=df1.set_index('rollno').join(get_particular_data.set_index('Roll No'))
				Remaining_student_data=pd.concat([df2,Remaining_student_data]) 

	F.close()
	try:
		Remaining_student_data.to_excel(output_file_name)
	except:
		pass

feedback_not_submitted()
