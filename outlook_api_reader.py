# https://stackoverflow.com/questions/22813814/clearly-documented-reading-of-emails-functionality-with-python-win32com-outlook
# https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?view=outlook-pia

from win32com.client import Dispatch


# function to open all emails
def allEmails():

	outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
	
	# get all accounts
	folders = outlook.Folders

	# choose canu account, get all folders
	account = folders[2]
	account_folders = account.Folders

	# define draft folder to send replies
	global drafts
	drafts = account_folders[3]
	
	''' # choose folder, get subfolder
	subfolder = account_folders[8]
	subfolder_folders = subfolder.Folders 
	# if using subfolders change line below
	'''
		
	# access folder, get all emails
	inbox_folder = account_folders[1]
	print("number of emails", inbox_folder.Items.Count)
	for i, email in enumerate(inbox_folder.Items):
	    print(i, email.Subject)
	
	return inbox_folder.Items



# function to get data from email
def getEmailData(email):

	lines = (email.Body).splitlines()
	
	# if email is not in requête format
	if not lines[-3].startswith('Date et heure'):
		print('\nSorry, email is not correctly formatted.')
		return
	
	dic = {}
	
	for i, line in enumerate(lines):

		# remove 'Coordonnées - '
		line = line.replace('Coordonnées - ', '')

		if line.startswith('Demande - Détails'):
			k, v = line.split(' : ', 1)
			dic[k] = v + '\n' + '\n'.join(lines[i+1:-1])
			
			# in case there is a : in the description
			k, v = lines[-1].split(' : ', 1)
			dic[k] = v
			
			break
			
		elif ' : ' in line:
			k, v = line.split(' : ', 1)
			dic[k] = (dic.get(k, '') + ' ' + v).lstrip()
			
	# create formatted phone number (only digits)
	number = ''
	for i in dic['Numéro de téléphone (jour)']:
		if i.isdigit():
			number +=i
	dic['Téléphone'] = number
			
	# record if requester is citizen or company
	dic['Type'] = 'Entreprise' if "Nom de l'entreprise ou de l'organisation" in dic.keys() else 'Citoyen'
	
	# save attachments and add directory
	atts = []
	for att in email.Attachments:
		if not str(att).startswith('image00'):
			fullpath = 'M:\\automate311\\attachements\\' + str(att)
			att.SaveAsFile(fullpath)
			atts.append(fullpath)
	
	dic['Attachments'] = atts
	
	
	return dic
	

# reply to emails
def createReply(email):
        reply = email.Reply()
        newBody = "this is the reply"
        reply.HTMLBody = newBody + reply.HTMLBody
        reply.Move(drafts)
		
