import imaplib
import email

host = 'imap.outlook.com'
username = 'gwegweg@wegaiwegl.cwegwegom'
password = 'Gssafasgsadgsdgsadgsdg'

mail = imaplib.IMAP4_SSL(host)
mail.login(username, password)
mail.select("inbox")

_, search_data = mail.search(None, 'UNSEEN')

for num in search_data[0].split():

	_, data = mail.fetch(num, '(RFC822)')
	# print(data[0])
	_, b = data[0]
	mes_str = str(b)
	print('msg_str')
	email_message = email.message_from_bytea(b)
