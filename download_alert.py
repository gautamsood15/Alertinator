import imaplib
import email

host = 'imap.outlook.com'                  # mailbox information
username = 'gauty22@hotmail.com'
password = 'Gs220797'

def get_inbox():

	mail = imaplib.IMAP4_SSL(host)           # using mailbox information to login
	mail.login(username, password)
	mail.select("inbox")

	_, search_data = mail.search(None, 'UNSEEN')          # searching for all unseen emails in inbox
	my_message = []

	for num in search_data[0].split():              # extracting the useful information from rest of the code 
		email_data = {}
		_, data = mail.fetch(num, '(RFC822)')
		_, b = data[0]
		
		email_message = email.message_from_bytes(b)           # remove the bytes from the messages
		

		for header in ['subject','to','from','date']:              # extract subject, to, from, data from the unread mails
			# print("{}: {}".format(header,email_message[header]))
			email_data[header] = email_message[header]

		for part in email_message.walk():                         # get the body of the mails
			if part.get_content_type() == "text/plain":             # get body of mails if mail is text type
				body = part.get_payload(decode=True)
				email_data['body'] = body.decode()


			elif part.get_content_type() == "text/html":              # get body of mails if mail is html type
				html_body = part.get_payload(decode=True)
				email_data['html_body'] = html_body.decode()
		
		my_message.append(email_data)
	return my_message



#                         code exection

my_inbox = get_inbox()
print(my_inbox)