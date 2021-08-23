"""
Tutorial 2
Deepika Rajwar
1901EE19
"""
def get_memory_score(input_nums):
	n=len(input_nums)	#size of input list
	score=0
	tmp=[]	#empty list
	for x in input_nums:
		if type(x)!=int:
			print("Please enter a valid input list. Invalid inputs detected")
			return
	for i in range(n):
		f=int(input_nums[i])
		if len(tmp)<5:
			if f in tmp:
				score+=1
			else:
				tmp.insert(len(tmp),f)	#insert element at last position
			if len(tmp)>5:
				tmp.pop(0)	#remove the number that has been in the memory the longest time
		#print("Called ",f," Score ",score," ",*tmp,sep=",")
	return score

input_nums = [3, 4, 5, 3, 2, 1]

print("Score: "get_memory_score(input_nums))