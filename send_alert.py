
import os
import re
import win32com.client as client


# Choose between the clients to send email to

def project_selection():

    with open('service_alert.txt') as input_file:

        file_content = input_file.readlines()

        for i, line in enumerate(file_content):

            if i == 2:                                    
                if line == 'To:	django.Messaging\n':                            # if the DL is of django projects
                    to_addresses = 'gauty22@gmail.com; gauty22@hotmail.com'
                    cc_addresses = 'gauty22@gmail.com'

                if line == 'To:	Unchained_O365\n':                                 # if the DL is of unchained projects
                    to_addresses = 'gasdas@ac.in; gasf424@hotmail.com'
                    cc_addresses = 'gauty22@gmail.com'

                else:
                    break

    return to_addresses, cc_addresses


# adding header to the alert info

def add_header():

    alert_info = open('alert_info.txt', 'a')
    alert_info.write(
        '<br/><h2 style="color:blue;margin-left:400px;">Microsoft Service Degradation - Alert</h2><br/><br/><br/>')
    alert_info.close()

    return


# retrieving message from service alert

def alert_parser():

    x = 0
    y = 0

    with open('service_alert.txt') as input_file:

        file_content = input_file.readlines()

    # extracting alert info from service alert

        for i, line in enumerate(file_content):

            if i == 13:                                                 # to get the alert type
                if line == 'Exchange Online service alert\n':
                    is_office_alert = True

                elif line == 'Microsoft Teams service alert\n':
                    is_office_alert = True

                else:
                    is_office_alert = False

            if re.search("^ID:", line):                                 # to get the ID of the alert
                alert_id = line

            if x == 1:                                                 # To get the service status
                if line == 'Service Degradation\n':
                    is_service_degradation = True
                else:
                    is_service_degradation = False
                x = 0

            if line == 'Status\n':                                  
                x = 1

            if y == 1:                                                  # to get the alert info main body from service alert
                if line == 'Are you experiencing this issue?\n':
                    y = 0

                else:
                    with open('alert_info.txt', 'a') as alert_info:
                        alert_info.write('&emsp;&emsp;'+line+'<br/><br/>')

            if line == 'Details\n':
                y = 1

    # delete service alert

    os.remove("service_alert.txt")

    return is_office_alert, is_service_degradation, alert_id


# Adding signature to alert info

def add_signature():

    with open('alert_info.txt', 'a') as alert_info:

        alert_info.write('<br/><br/>')
        alert_info.write(
            '<div><img src="https://www.crwflags.com/fotw/images/u/us$accnt.gif" align="left" width="180" height="80">')
        alert_info.write('&ensp;Thanks and Regards,<br/>')
        alert_info.write('&ensp;Gautam Sood,<br/>')
        alert_info.write('&ensp;Messaging Team,<br/>')
        alert_info.write('&ensp;Accenture Services Private Limited</div><br/>')


#  validate for messaging team to send to client


def is_validated(office_alert, service_degradation):

    valid = True

    if office_alert == False:                          # check if the alert type is correct 
        print("False Alert\n")
        print("It is NOT an O365 Alert")
        valid = False

    elif service_degradation == False:                  # check if the alert is for service degradation
        print("False Alert\n")
        print("Status is NOT Service Degradation")
        valid = False

    else:                                               # If the alert is a valid alert 
        print("Valid Alert")

    return valid


# Send email to the clients


def send_email(to_addresses, cc_addresses, alert_id, office_alert, service_degradation):

    outlook = client.Dispatch("Outlook.Application")                   # connecting with outlook application
    message = outlook.CreateItem(0)
    message.Display()

    message.To = to_addresses                                          # adding sender info in outlook
    message.CC = cc_addresses

    message.Subject = alert_id + " - M365 Service Health Notification"    # adding subject to outlook

    with open('alert_info.txt') as alert_info:                          # adding body to outlook

        file_content = alert_info.read()

    message.HTMLBody = file_content

    if is_validated(office_alert, service_degradation) == True:        # check if the alert is valid, if yes only then send
        message.Save()
        message.Send()

    else:
        print("No need to send alert to client")