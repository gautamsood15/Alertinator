# copy Alert info data from the text file to another file 
import os


with open('service_alert.txt') as input_file:
	
	file_content = input_file.readlines()

	for i, line in enumerate(file_content):

		if i == 13:
			if line == 'Microsoft 365 suite service alert\n':
				is_office_alert = True
			else:
				is_office_alert = False

		if i == 20:
			if line == 'Service Degradation\n':
				is_service_degradation = True 
			else:
				is_service_degradation = False


		if 44 > i > 21:
			with open('alert_info.txt', 'a') as alert_info:
				alert_info.write(line)

	print(is_service_degradation)
	print(is_office_alert)

os.remove("service_alert.txt")






