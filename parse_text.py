
import os
import win32com.client as client

# adding header to the alert info 

def add_header():

	alert_info = open('alert_info.txt', 'a')
	alert_info.write('\n                                     Microsoft Service Degradation - Alert\n\n\n')
	alert_info.close()

	return



# retrieving message from service alert

def alert_parser():
	
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

	return is_service_degradation, is_office_alert, alert_id





# Adding signature to alert info

def add_signature():

	with open('alert_info.txt', 'a') as alert_info:

		alert_info.write('\n\n')
		alert_info.write('Thanks and Regards,\n')
		alert_info.write('Gautam Sood,\n')
		alert_info.write('Messaging Team\n')
		alert_info.write('Accenture Services Private Limited\n')
		alert_info.write('Email: gautam.a.sood@accenture.com\n')




#  validate for messaging team to send to client


def is_validated():
	
	office_alert, service_degradation, dummy_alert_id = alert_parser()

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






# Choose between the projects to send 

def project_selection():
	pass
	







# Send email to the clients 


def send_email():
	
	outlook = client.Dispatch("Outlook.Application")
	message = outlook.CreateItem(0)
	message.Display()

	message.To = "shailsood15@gmail.com; gauty22@gmail.com"
	message.CC = "gauty22@hotmail.com; gauty@gautamsood.in"


	alert_id, is_office_alert, is_service_degradation = alert_parser()

	message.Subject = alert_id + " - M365 Service Health Notification"
	message.body = "This is a test case"

	message.Save()
	message.Send()






#------------------------  Code Execution -------------------------------------------------

add_header()
is_validated()
add_signature()



send_email()