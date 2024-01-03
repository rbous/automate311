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
	
