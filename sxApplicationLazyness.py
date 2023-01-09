from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from os.path import exists
import os
from time import sleep
import re

if(exists("applications.md")):
	os.remove("applications.md")

pages = int(input("Hvor mange siders ansøgninger er der? "))

driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()))
applications = []

def formatter(applications):
	print("Formatting document")
	applications = sorted(applications, key=lambda d: d['name']) 
	file = open("applications.md", "a")
	file.write("### Hold 1: \n ### Hold 2: \n ### Hold 3: \n")
	tableHeader = f'| Navn | Er blevet snakket med? | Godkendt/afvist | \n |---|---|---| \n'
	file.write(tableHeader)
	for application in applications:
		row = f'|{application["name"]}| | | | | \n'
		file.write(row)
	file.write("\n \n")
	for application in applications:
		tableHeader = f'| {application["name"]} | \n |---| \n'
		name = f'| **Navn:** [{application["name"]}]({application["link"]})  <br>'
		steamid = f' **SteamID:** [{application["steamid"]}]({application["dashboard"]})   <br>'
		good = f'**Godt:**  <br><br>'
		bad = f'**Daarligt:**  <br><br>'
		other = f'**Andet:** <br><br>'
		decision = f'**Beslutning:** <br><br> **Grund (Hvis afvist)**: <br> | \n \n'
		formatted = tableHeader + name + steamid + good + bad + other + decision
		file.write(formatted)
	file.close()

def fetchApplications():
	print("Fetching applications")
	for i in range(1, pages + 1):
		driver.get("https://stavox.dk/forums/forum/5-staffans%C3%B8gninger-%C3%A5bne/?page=" + str(i))
		list1 = driver.find_element(By.XPATH,'/html/body/main/div/div/div/div[3]/div/ol')
		elements = list1.find_elements(By.CSS_SELECTOR, "li > div > h4 > span > a")
		sleep(2)
		for element in elements:
			if re.search("\[Accepteret\]", element.text) == None and re.search("\[Afslået\]", element.text) == None:
				if (element.text != "Mundtlig Staffansøgning - How to [BEMÆRK NY FORMAT]"):
					applicant = {
						"name": element.text.title(),
						"link": element.get_attribute('href'),
						"steamid": "",
						"dashboard": "",
					}
					applications.append(applicant)

fetchApplications()
numOfApplications = 0
for application in applications:
	driver.get(application["link"])
	sleep(0.5)
	steamid = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.XPATH, '/html/body/main/div/div/div/div[3]/div[2]/form/article[1]'))
	)
	lowerApplicationText = steamid.text.lower()
	if(("steamid" in lowerApplicationText) or ("steam id" in lowerApplicationText)):
		if(re.search("STEAM_.*", steamid.text)):
			applicantSteamID = re.findall("STEAM_.*", steamid.text)
		else:
			applicantSteamID = re.findall("steam_.*", steamid.text)
		application['steamid'] = applicantSteamID[0]
		application['dashboard'] = "https://stavox.dk/dash/lookup?lookupid=" + applicantSteamID[0]
	numOfApplications = numOfApplications + 1
formatter(applications)
driver.close()

print(f'Found a total of {numOfApplications} applications')
print("Applications has been saved in applications.md")