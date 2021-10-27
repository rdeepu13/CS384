#Tutorial 6

#Deepika Rajwar

#1901EE19


import os

import re

import shutil


def regex_renamer():



	# Taking input from the user



	print("1. Breaking Bad")

	print("2. Game of Thrones")

	print("3. Lucifer")



	webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))

	season_padding = int(input("Enter the Season Number Padding: "))

	episode_padding = int(input("Enter the Episode Number Padding: "))




	if webseries_num==1:
		src='./wrong_srt/Breaking Bad/'

		dest='./corrected_srt/Breaking Bad/'
	elif webseries_num==2:
		src='./wrong_srt/Game of Thrones/'
		dest='./corrected_srt/Game of Thrones/'
	else:
		src='./wrong_srt/Lucifer/'

		dest='./corrected_srt/Lucifer/'


	os.makedirs(dest)	#Create directory
	for filename in os.listdir(src):
		r=re.split(r'-',filename)

		if webseries_num==1:
	#no title in case of breaking bad
			x=re.split(r'\D',r[1])

		else:
			x=re.split(r'x',r[1])

			title=re.split('\.',r[2])
		season =  "0"*(season_padding-len(x[0])) + x[0]
		episode=  "0"*(episode_padding-len(x[1])) + x[1]
		newFilename=" "
		if webseries_num==1 and filename.endswith('mp4'):
			newFilename=r[0]+" - Season "+season+" Episode "+ episode +".mp4"

		elif webseries_num==1 and filename.endswith('srt'):
			newFilename=r[0]+" - Season "+season+" Episode "+ episode +".srt"
		elif filename.endswith('mp4'):
			newFilename=r[0]+" - Season "+season+" Episode "+ episode + " - " +title[0]+".mp4"

		else:
			newFilename=r[0]+" - Season "+season+" Episode "+ episode + " - " +title[0]+".srt"
		os.rename(src+filename,src+newFilename)
		shutil.move(src+newFilename,dest+newFilename)



regex_renamer()