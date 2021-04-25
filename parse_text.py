
import os
import win32com.client as client

# adding header to the alert info 

alert_info = open('alert_info.txt', 'a')
alert_info.write('                                     Microsoft Service Degradation - Alert\n\n\n')
alert_info.close()


# retrieving message from service alert


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

		if i == 18:
			alert_id = line


		if 44 > i > 21:
			with open('alert_info.txt', 'a') as alert_info:
				alert_info.write('\t'+line)


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




# Check if the service alert is valid for messaging team to send to client

def is_validated(office_alert, service_degradation):
	
	if office_alert == False:
		os.remove("alert_info.txt")
		print("False Alert\n")
		print("It is NOT an O365 Alert")
	
	elif service_degradation == False:
		os.remove("alert_info.txt")
		print("False Alert\n")
		
		print("Status is NOT Service Degradation")

	else:
		print("Valid Alert")

	return

is_validated(is_office_alert, is_service_degradation)




# Send email script 

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Display()

message.To = "shailsood15@gmail.com; gauty22@gmail.com"
message.CC = "gauty22@hotmail.com; gauty@gautamsood.in"

message.Subject = alert_id + " - M365 Service Health Notification"
message.body = "This is a test case"

message.Save()
message.Send()


