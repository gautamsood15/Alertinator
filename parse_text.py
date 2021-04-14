
import os
import win32com.client as client


# Choose between the clients to send email to

def project_selection():

    with open('service_alert.txt') as input_file:

        file_content = input_file.readlines()

        for i, line in enumerate(file_content):

            if i == 2:
                if line == 'To:	IO.Hess.Messaging\n':
                	to_addresses = 'gauty22@gmail.com; gauty22@hotmail.com'
                	cc_addresses = 'shailsood15@gmail.com'

                if line == 'To:	Upfield_O365\n':
                    to_addresses = 'gautamsood15@stu.upes.ac.in; gauty22@hotmail.com'
                    cc_addresses = 'shailsood15@gmail.com'

                else:
                    break

    return to_addresses, cc_addresses


# adding header to the alert info

def add_header():

    alert_info = open('alert_info.txt', 'a')
    alert_info.write('<br/><h2 style="color:blue;margin-left:400px;">Microsoft Service Degradation - Alert</h2><br/><br/><br/>')
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

            if 34 > i > 21:
                with open('alert_info.txt', 'a') as alert_info:
                    alert_info.write('&emsp;&emsp;'+line+'<br/><br/>')

    # delete service alert

    #os.remove("service_alert.txt")

    return is_office_alert, is_service_degradation, alert_id


# Adding signature to alert info

def add_signature():

    with open('alert_info.txt', 'a') as alert_info:

        alert_info.write('<br/><br/>')
        alert_info.write('<div><img src="https://www.crwflags.com/fotw/images/u/us$accnt.gif" align="left" width="180" height="80">')
        alert_info.write('&ensp;Thanks and Regards,<br/>')
        alert_info.write('&ensp;Gautam Sood,<br/>')
        alert_info.write('&ensp;Messaging Team,<br/>')
        alert_info.write('&ensp;Accenture Services Private Limited</div><br/>')


#  validate for messaging team to send to client


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








# Send email to the clients


def send_email(to_addresses, cc_addresses, alert_id):

    outlook = client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Display()

    message.To = to_addresses
    message.CC = cc_addresses

    message.Subject = alert_id + " - M365 Service Health Notification"



    with open('alert_info.txt') as alert_info:

    	file_content = alert_info.read()

    message.HTMLBody = file_content



    message.Save()
    message.Send()


# ------------------------  Code Execution -------------------------------------------------


to_addresses, cc_addresses = project_selection()

add_header()

office_alert, service_degradation, alert_id = alert_parser()

is_validated(office_alert, service_degradation)

add_signature()

send_email(to_addresses, cc_addresses, alert_id)

# os.remove("alert_info.txt")