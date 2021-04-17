
#Alertinator

It is a Project in python to download Microsoft service alert emails from an outlook mailbox.
Then extract the alert information from the service alert mail. After that, it creates a custom,
user friendly mail to clients of the specific project for which the service mail has arrived.

 
##Modules/Language Required for Project

-> python 3.x -- language used to make the project
-> autopep8   -- Formats code automatically to PEP8 style guide		[autopep8 -i file_name], [pip install autopep8 --user]
-> os         -- Used to remove unwanted files from the project directory
-> re         -- Used to search for text in alert info to get the proper data
-> win32com   -- Used to access outlook and to send mails using CreateItem() [pip install pypiwin32 --user]
-> imaplib    -- Allows connection to IMAP4 server 
-> email      -- Allow us to manage livrary messages, retrieve messages information, search messages 

