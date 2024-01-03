from outlook_api_reader import allEmails, getEmailData
from gui_func import openCitizenSheet, createCitizen, createRequest


# select all emails with category set to Rayane, get data from email
			
success = False
for email in allEmails():
	if email.Categories == 'Rayane':
		data = getEmailData(email)
		print('\n', data)
		success = True
		
print('\nDATA RETRIEVAL SUCCESSFUL' if success is True else '\nDATA RETRIEVAL UNSUCCESSFUL')



if data['Type'] == 'Citoyen':

	openCitizenSheet(data)
	
	while True:
		new = input('\nSelect citizen. Is there already an account? Type y/n.,\n').lower()
		
		if new != 'y' and new != 'n':
			print('\nWrong input')
			continue
			
		elif new == 'n':
			createCitizen(data)
		
		break
	
	createRequest(data)
	


'''
CITOYENS:
click 200, 80
wait 0.25 s
paste formatted phone number
click 300, 1000
let user choose citoyen (input if already or not)

# create requete
click 50, 250
let user input gabarit

click 400, 400
paste formatted adress (leave 2 seconds to fix it)
click 900, 400 (leave 2 seconds to choose adress)
if app number, click 900, 515, paste app
click 800, 700
click 800, 710
scroll down completely
click 400, 300
paste description
scroll up completely

if att:
	click 700, 350
	TODO: finish attachement

click 400, 1000
wait 3 seconds
click 250,350

click 400, 250 2 times
copy (request number)



'''