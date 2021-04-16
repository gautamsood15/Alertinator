from download_alert import *
from parse_text import *

infinite_loop = False

while(infinite_loop):               #    execution loop for the code
	
	save_alert()

	#-------------------------------------------------

	to_addresses, cc_addresses = project_selection()

	add_header()

	office_alert, service_degradation, alert_id = alert_parser()

	add_signature()

	send_email(to_addresses, cc_addresses, alert_id, office_alert, service_degradation)

	os.remove("alert_info.txt")
	
