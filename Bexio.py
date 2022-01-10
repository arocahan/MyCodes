from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
import sys


zeitDau = "2021-12"
apoName = 'Apotheke am Wipkingerplatz'

def switch(value):
    return{
        'Schwanen Apotheke Baden AG': 23,
        'Benu Apotheke Bahnhof Visp': 19,
        'BENU Apotheke Barfuesser'	: 38,
        'Benu Apotheke Dr. Villa, Chur' : 29,
        'BENU Apotheke Faellanden'	: 34,
        'Benu Apotheke Huttwil' 	: 39,
        'Benu Landquart (unter Toppharm Dr. Villa)': 35,
        'Benu Apotheke Langnau am Albis'	: 18,
        'Galexis Apotheke Niederbipp'	: 8,
        'Galenicare Management AG'	: 27,
        'BENU Apotheke Goldach'	: 22, 
        'Apotheke Zur Rose Schoenbuehl Shoppyland'	: 31,
        'Apotheke Zur Rose Buchs Wynecenter'	: 36,
        'MBZR Apotheken AG (Buchs)'	: 21,
        'MBZR Apotheken AG (Limmatplatz)'	: 24,
        'MBZR Apotheken AG (Marktgasse Bern)' : 47,
        'Apotheke Zur Rose Spreitenbach Tivoli'	: 16,
        'Medbase Apotheke Biel Brueggmoos'	: 15,
        'Medbase Apotheke Glattbrugg 1516'	: 9,
        'Medbase Apotheke Kreis 12 1519'	: 14,
        'Medbase Apotheke Meiringen 1562'	: 17,
        'Medbase Apotheke St. Gallen 1558'	: 13,
        'Medbase Apotheke Uster 1516'	: 12,
        'Pfauen Apotheke AG'	: 11,
        'Schwanen Apotheke Baden AG'	: 23,
        'Toppharm Apotheke Siebnen'	: 10,
        'TopPharm Europaallee Apotheke'	: 37,
        'TopPharm Fortuna Apotheke, Chur'	: 33,
        'Apotheke im Freihof Bruettisellen' : 44,
        'Kantonsapotheke Zürich' : 45,
        'Medbase Apotheke Zuchwil' : 40,
        'Medbase Apotheke Madretsch Biel' : 0,
        'SunStore Apotheke Basel S.Jakob'   : 48,
        'Weissenbuehl Apotheke Dr Gurtner Bern'   : 0,
        'Benu Apotheke Eulen, ZH' : 46,
        'Apotheke Zur Rose Marktgasse Bern' : 47,
        'BENU Apotheke Menziken' : 49,
        'Schermen Apotheke Mauron Bern' : 50,
        'BENU Apotheke Gesundheits-Forum Witikon' : 51,
        'SunStore Binningen' : 52,
        'SunStore Apotheke Storchengaesschen Bern' : 57,
        'Apotheke Zur Rose Crissier' : 58,
        'SunStore Apotheke Biel Centre' : 59,
        'Apotheke am Wipkingerplatz' : 54,
        'Benu Apotheke Kloten'  : 42,
        'SunStore Apotheke Biel SBB'  : 63,
        'SunStore Apotheke Nussbaumen'  : 64,
        'SunStore Apotheke Bethlehem'  : 65,
        'Benu Apotheke Freiburg Bahnhof'  : 61,
        'Benu Apotheke Rex Wohlen'  : 62,
        'Apotheke Zur Rose Basel Claramarkt'    :   60
    }.get(value, 0)
    
#print(switch(apoName))

#Seite aufrufen 



driver = webdriver.Chrome(r"C:\Users\arael\Downloads\chromedriver")

driver.get("https://idp.bexio.com/login")
driver.maximize_window()

#Login
url = 'https://idp.bexio.com/login'
values = {'username': '----',
          'password': '------'}
inputEmail = driver.find_element_by_id('j_username')
inputEmail.send_keys('------')
inputPassword = driver.find_element_by_id('j_password')
inputPassword.send_keys('----------')
button = driver.find_element_by_class_name('button')
button.click()


#findet es den Kunde und offnet die neue Rechnung Seite
kontakID = switch(apoName)
driver.get("https://office.bexio.com/index.php/kontakt/show/id/" + str(kontakID) + "#invoices")
time.sleep(2)
button1 = driver.find_element_by_link_text('Neue Rechnung')
button1.click()
time.sleep(2)
button2 = driver.find_element_by_xpath('//*[@id="editKbItemForm"]/div/div/div[6]/button')
button2.click()
time.sleep(2)
button3 = driver.find_element_by_xpath('//*[@id="mainTab"]/div[1]/div/div[1]/div[1]/div/div[1]/div[1]/a')
button3.click()

#Excel
df = openpyxl.load_workbook(r'C:\Users\arael\Desktop\2021_Pharmy_plan.xlsx', data_only=True)
worksheet = df['MJJA']

#-------------------------------------------Stundenhonorar-----------------------------------------------

#Formular ausfüllen
#time.sleep(2)
text = driver.find_element_by_xpath('//*[@id="kb_position_custom_text_ifr"]')
time.sleep(2)
text.click()
text.send_keys(Keys.ENTER)
text.send_keys('Stundenhonorar')
text.send_keys(Keys.ENTER)

totalStunde = 0.0
for i in range(1, 1000):
    s = worksheet['A' + str(i)].value
    ss = str(s)
    if zeitDau in ss:
        apo = worksheet['C' + str(i)].value
        datum = worksheet['A' + str(i)].value
        von = worksheet['D' + str(i)].value
        bis = worksheet['E' + str(i)].value
        stunden = worksheet['F' + str(i)].value
        if apo == apoName:
            text.send_keys(str(datum)[0:10] + ': ' + str(von)[0:5] + ' - ' + str(bis)[0:5] + ' Uhr (' + str(stunden) + 'h)')
            text.send_keys(Keys.ENTER)
            totalStunde += stunden

#Preise und Menge ausfüllen
time.sleep(2)
button4 = driver.find_element_by_id('kb_position_custom_amount')
button4.send_keys(str(round(totalStunde, 2)))
time.sleep(2)
button5 = driver.find_element_by_id('kb_position_custom_unit_price')
button5.send_keys(120)

#Entwurf speichern
time.sleep(2)
buttonx = driver.find_element_by_xpath('//*[@id="positions"]/div/div/div/div/div/div/div[2]/div/form/div/div[4]/button[1]')
buttonx.click()


#-------------------------------------------Fahrtzeit-----------------------------------------------


time.sleep(2)
button3 = driver.find_element_by_xpath('//*[@id="mainTab"]/div[1]/div/div[1]/div[1]/div/div[1]/div[1]/a')
button3.click()

#Formular ausfüllen
time.sleep(2)
text = driver.find_element_by_xpath('//*[@id="kb_position_custom_text_ifr"]')
text.click()
text.send_keys(Keys.ENTER)
text.send_keys('Fahrtzeit')
text.send_keys(Keys.ENTER)

"""
totalFahrtzeit = 0
for i in range(2, 1000):
    s = worksheet['A' + str(i)].value
    ss = str(s)
   
    if zeitDau in ss:
        apo = worksheet['C' + str(i)].value
        datum = worksheet['A' + str(i)].value
        stunden = worksheet['F' + str(i)].value
        Fahrtzeit = worksheet['G' + str(i)].value
        if Fahrtzeit != 0 and Fahrtzeit != None:
            FahrtzeitBez = float(Fahrtzeit) - 1.0
           
        if apo == apoName and FahrtzeitBez > 0:
            text.send_keys(str(datum)[0:10]+ ': ' + str(round(Fahrtzeit, 2)) + 'h, davon ' + str(round(FahrtzeitBez, 2)) + 'h bezahlt')
            text.send_keys(Keys.ENTER)
            totalFahrtzeit += FahrtzeitBez 

#Preise und Menge ausfüllen
time.sleep(2)
button6 = driver.find_element_by_id('kb_position_custom_amount')
button6.send_keys(str(round(totalFahrtzeit, 2)))
time.sleep(2)
button5 = driver.find_element_by_id('kb_position_custom_unit_price')
button5.send_keys(120)

#Entwurf speichern
time.sleep(2)
buttonx = driver.find_element_by_xpath('//*[@id="positions"]/div/div/div/div/div/div/div[2]/div/form/div/div[4]/button[1]')
buttonx.click()

"""
#-------------------------------------------km-----------------------------------------------

time.sleep(2)
driver.execute_script("window.scrollTo(0, 0);")



buttonp = driver.find_element_by_xpath('//*[@id="mainTab"]/div[1]/div/div[1]/div[1]/div/div[1]/div[1]/a')
buttonp.click()

#Formular ausfüllen
time.sleep(2)
text = driver.find_element_by_xpath('//*[@id="kb_position_custom_text_ifr"]')
text.click()
text.send_keys(Keys.ENTER)
text.send_keys('Km')
text.send_keys(Keys.ENTER)

#--------------
totalKm = 0
for i in range(1, 1000):
    s = worksheet['A' + str(i)].value
    ss = str(s)
    if zeitDau in ss:
        apo = worksheet['C' + str(i)].value
        datum = worksheet['A' + str(i)].value
        von = worksheet['D' + str(i)].value
        bis = worksheet['E' + str(i)].value
        stunden = worksheet['F' + str(i)].value
        km = worksheet['J' + str(i)].value
    
        if apo == apoName and (km != None and km != 0):
            text.send_keys(str(datum)[0:10]+ ': ' + str(km) + ' km')
            text.send_keys(Keys.ENTER)
            totalKm += km 

 
#Preise und Menge ausfüllen
time.sleep(2)
button7 = driver.find_element_by_id('kb_position_custom_amount')
button7.send_keys(str(round(totalKm, 2)))
time.sleep(2)
button8 = driver.find_element_by_id('kb_position_custom_unit_price')
button8.send_keys(str(0.7))

#Entwurf speichern
time.sleep(2)
buttony = driver.find_element_by_xpath('//*[@id="positions"]/div/div/div/div/div/div/div[2]/div/form/div/div[4]/button[1]')
buttony.click()

#-------------------------------------------Fahrtticket-----------------------------------------------

time.sleep(2)
driver.execute_script("window.scrollTo(0, 0);")

buttonp = driver.find_element_by_xpath('//*[@id="mainTab"]/div[1]/div/div[1]/div[1]/div/div[1]/div[1]/a')
buttonp.click()

#Formular ausfüllen
time.sleep(2)
text = driver.find_element_by_xpath('//*[@id="kb_position_custom_text_ifr"]')
text.click()
text.send_keys(Keys.ENTER)
text.send_keys('Fahrtticket')
text.send_keys(Keys.ENTER)

totalTicket = 0
for i in range(1, 1000):
    s = worksheet['A' + str(i)].value
    ss = str(s)
    #print(ss)
    if zeitDau in ss:
        apo = worksheet['C' + str(i)].value
        datum = worksheet['A' + str(i)].value
        ticket = worksheet['I' + str(i)].value
    
        if apo == apoName and (ticket != 0 and ticket != None):
            text.send_keys(str(datum)[0:10]+ ': ' + str(ticket) + ' CHF') 
            text.send_keys(Keys.ENTER)
            totalTicket += ticket 


#Preise und Menge ausfüllen
time.sleep(2)
button7 = driver.find_element_by_id('kb_position_custom_amount')
button7.send_keys(str(round(totalTicket, 2)))
time.sleep(2)
button8 = driver.find_element_by_id('kb_position_custom_unit_price')
button8.send_keys(str(1))

#Entwurf speichern
time.sleep(2)
buttony = driver.find_element_by_xpath('//*[@id="positions"]/div/div/div/div/div/div/div[2]/div/form/div/div[4]/button[1]')
buttony.click()




#-------------------------------------------Parkgebuehr-----------------------------------------------

time.sleep(2)
driver.execute_script("window.scrollTo(0, 0);")

buttonp = driver.find_element_by_xpath('//*[@id="mainTab"]/div[1]/div/div[1]/div[1]/div/div[1]/div[1]/a')
buttonp.click()

#Formular ausfüllen
time.sleep(2)
text = driver.find_element_by_xpath('//*[@id="kb_position_custom_text_ifr"]')
text.click()
text.send_keys(Keys.ENTER)
text.send_keys('Parkgebuehr')
text.send_keys(Keys.ENTER)



totalParkgebuehr = 0
for i in range(1, 1000):
    s = worksheet['A' + str(i)].value
    ss = str(s)
    #print(ss)
    if zeitDau in ss:
        apo = worksheet['C' + str(i)].value
        datum = worksheet['A' + str(i)].value
        parkticket = worksheet['L' + str(i)].value
    
        if apo == apoName and (parkticket != 0 and parkticket != None):
            text.send_keys(str(datum)[0:10]+ ': ' + str(parkticket) + ' CHF') 
            text.send_keys(Keys.ENTER)
            totalParkgebuehr += parkticket 



#Preise und Menge ausfüllen
time.sleep(2)
button7 = driver.find_element_by_id('kb_position_custom_amount')
button7.send_keys(str(round(totalParkgebuehr, 2)))
time.sleep(2)
button8 = driver.find_element_by_id('kb_position_custom_unit_price')
button8.send_keys(str(1))

#Entwurf speichern
time.sleep(2)
buttony = driver.find_element_by_xpath('//*[@id="positions"]/div/div/div/div/div/div/div[2]/div/form/div/div[4]/button[1]')
buttony.click()



#-------------------------------------------Verpflegungskosten-----------------------------------------------

time.sleep(2)
driver.execute_script("window.scrollTo(0, 0);")

buttonp = driver.find_element_by_xpath('//*[@id="mainTab"]/div[1]/div/div[1]/div[1]/div/div[1]/div[1]/a')
buttonp.click()

#Formular ausfüllen
time.sleep(2)
text = driver.find_element_by_xpath('//*[@id="kb_position_custom_text_ifr"]')
text.click()
text.send_keys(Keys.ENTER)
text.send_keys('Verpflegungskosten')
text.send_keys(Keys.ENTER)


totalUebernachtung = 0
for i in range(1, 1000):
    s = worksheet['A' + str(i)].value
    ss = str(s)
    #print(ss)
    if zeitDau in ss:
        apo = worksheet['C' + str(i)].value
        datum = worksheet['A' + str(i)].value
        uberNach = worksheet['M' + str(i)].value
    
        if apo == apoName and (uberNach != 0 and uberNach != None):
            text.send_keys(str(datum)[0:10]+ ': ' + str(uberNach) + ' CHF') 
            text.send_keys(Keys.ENTER)
            totalUebernachtung += uberNach 



#Preise und Menge ausfüllen
time.sleep(2)
button7 = driver.find_element_by_id('kb_position_custom_amount')
button7.send_keys(str(round(totalUebernachtung, 2)))
time.sleep(2)
button8 = driver.find_element_by_id('kb_position_custom_unit_price')
button8.send_keys(str(1))

#Entwurf speichern
time.sleep(2)
buttony = driver.find_element_by_xpath('//*[@id="positions"]/div/div/div/div/div/div/div[2]/div/form/div/div[4]/button[1]')
buttony.click()

#----------------------------------------------------------------

if (driver.close()):
    sys.exit()

