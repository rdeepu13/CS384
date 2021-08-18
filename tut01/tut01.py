"""
Tutorial 1
Deepika Rajwar
1901EE19
"""

def meraki_helper(n):
	"""This will detect meraki number"""
	num=str(n)	#stores number in string format
	for i in range(0, len(num)-1):
		c=int(num[i])
		d=int(num[i+1])
		if	abs(c-d)!=1:
			print("No -",n," not a Meraki number")
			return	0
			break;
	print("Yes -",n," is a Meraki number")
	return	1


input = [12,24,56,78,98,54,678,134,789,0,7,5,123,45,76345,987654321]

count=0   #stores the frequency of meraki numbers
for i in range(0, len(input)):
	count+=meraki_helper(input[i])

print("the input list contains ",count," meraki and ",(len(input)-count)," non meraki numbers")
