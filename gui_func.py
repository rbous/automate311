import pyautogui as g
from time import sleep
from pyperclip import copy, paste



def openCitizenSheet(data):
	g.click(200, 80)
	sleep(0.1)
	g.click(1000, 600)
	sleep(0.7)
	g.click(600, 295)
	g.click(600, 295)
	sleep(0.1)
	g.typewrite(data['Téléphone'][-7:-4] + '-' + data['Téléphone'][-4:])
	g.click(330, 1000)
	return
	
	
def createCitizen(data):
	g.click(50, 250)
	sleep(0.6)
	g.click(350, 270)
	g.typewrite(data['Nom'])
	g.press('Tab')
	g.typewrite(data['Prénom'])
	g.press('Tab')
	g.typewrite(data['Adresse - Numéro de l\'immeuble'] + ' ' +  data['Adresse - Nom de la rue'])
	g.press('Tab')
	g.press('Tab')
	g.typewrite(data['Ville'])
	g.click(900, 275)

	# pyautogui doesn't copy '@' symbol
	copy(data['Courriel'])
	g.hotkey('ctrl', 'v')

	g.click(900, 295)
	g.typewrite(data['Téléphone'][-10:-7] + ' ' +  data['Téléphone'][-7:-4] + '-' + data['Téléphone'][-4:])
	g.click(370, 1000)
	sleep(1)
	return
	
	
def createRequest(data):
	sleep(1)
	g.click(330, 1000)
	sleep(1.3)
	g.click(50, 250)
	g.press('Tab')
	g.press('Tab')
	print('\n' + data["Demande - Type de requête"] + '\n' + data["Demande - Détails"])
	input('\nEnter Gabarit. Type anything when done.')
	g.click(400, 400)
	copy(data["Demande - Localisation de l'incident"])
	g.hotkey("ctrl", "v")
	g.click(800, 700)
	g.click(800, 710)
	g.scroll(-1000)
	g.click(400, 300)
	copy(data['Demande - Détails'])
	g.hotkey("ctrl", "v")
	g.click(1200, 500)
	g.scroll(1000)
	
	atts = data['Attachments']
	if atts:
		g.click(700, 350)
		for _ in atts:
			g.click(260, 405)
		g.moveTo(300, 400)
		for att in atts:
			g.moveRel(0, 20)
			g.click()
			g.click()
			sleep(0.2)
			g.hotkey('ctrl', 'l')
			g.typewrite('M:/automate311/attachements')
			g.press('enter')
			for _ in range(6):
				g.press('Tab')
			addy = att.split('\\')[-1]
			copy(addy)
			g.hotkey("ctrl", "v")
			g.press('enter')
			sleep(0.5)
			
	g.click(250, 350)
	g.click(1000, 420)
	sleep(5)
	
	g.click(400, 1000)
	sleep(2)
	g.click(250, 350)
	g.click(400, 250)
	g.click(400, 250)
	g.hotkey("ctrl", "v")
	print('Request number:', paste())
	