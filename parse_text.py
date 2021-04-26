# copy Alert info data from the text file to another file 

with open('Test Email.txt') as file:
	file_content = file.readline()
	print (file_content, end='')

	file_content = file.readline()
	print (file_content, end='')




file.close()