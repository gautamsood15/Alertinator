import imaplib
import email

host = 'imap.outlook.com'
username = 'gauty22@hotmail.com'
password = 'Gs220797'

mail = imaplib.IMAP4_SSL(host)
mail.login(username, password)
mail.select("inbox")

_, search_data = mail.search(None, 'UNSEEN')

for num in search_data[0].split():

	_, data = mail.fetch(num, 'Test Test Test 1')
	# print(data[0])
	_, b = data[0]
	mes_str = str(b)
	print('msg_str')
	email_message = email.mnessage_from_bytea(b)
	