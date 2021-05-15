
import os


# adding header to the alert info 

alert_info = open('alert_info.txt', 'a')
alert_info.write('                                     Microsoft Service Degradation - Alert\n\n\n')
alert_info.close()


# retrieving data from service alert


with open('service_alert.txt') as input_file:
	
	file_content = input_file.readlines()

# extracting alert info from service alert

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

# vairables to check if the service alert is valid for messaging team to send to client

	print(is_service_degradation)
	print(is_office_alert)


# delete service alert 

os.remove("service_alert.txt")


# Adding signature to alert info


alert_info = open('alert_info.txt', 'a')

alert_info.write('\n\n')
alert_info.write('Thanks and Regards,\n')
alert_info.write('Gautam Sood,\n')
alert_info.write('Messaging Team\n')
alert_info.write('Accenture Services Private Limited\n')
alert_info.write('Email: gautam.a.sood@accenture.com\n')

alert_info.close()


