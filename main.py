from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import bs4
import xlrd
import xlwt
import time
import csv


file = "Liste_Villes.xls"
url = "https://www.royallepage.ca/fr/search/agents-offices/"

def OuvrirBrowser():
    brwser = webdriver.Chrome()
    return brwser

def AllerSurRemax(url,browser):
    browser.get(url)


def Recuperer_ville(NomFichierXls):
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    Ville = []
    for i in range(sheet.nrows):
        Ville.append(sheet.cell_value(i, 0))
    return Ville

def RechercherCourtierDuneVille(brwser, ville):
    brwser.implicitly_wait(1)
    search_bar = brwser.find_element_by_xpath("/html/body/div[2]/div/div[1]/section[1]/form/div[1]/div[1]/div/div[2]/div/div[1]/input")
    search_bar.clear()
    brwser.implicitly_wait(1)
    search_bar.click()
    search_bar.send_keys("*")
    emplacement_bar = brwser.find_element_by_xpath("/html/body/div[2]/div/div[1]/section[1]/form/div[1]/div[1]/div/div[2]/div/div[2]/input")
    emplacement_bar.clear()
    emplacement_bar.click()
    emplacement_bar.send_keys(ville)
    brwser.implicitly_wait(1)
    time.sleep(2)
    PremiereVilledansLaliste = brwser.find_element_by_id("suggestion-0").click()
    brwser.implicitly_wait(5)
    emplacement_bar = brwser.find_element_by_xpath("/html/body/div[2]/div/div[1]/section[1]/form/div[1]/div[1]/div/div[2]/div/div[2]/input")
    emplacement_bar.click()
    time.sleep(2)
    emplacement_bar.send_keys(Keys.ENTER)
    time.sleep(2)
        #Telephone = i.find("p").text.strip()
#def ChercherUneVille(Ville):



Villes = Recuperer_ville(file)
Browser = OuvrirBrowser()
AllerSurRemax(url, Browser)


ListeCourtier = []
while(len(Villes) > 0):
    end = 0
    PageNumbering = 1
    RechercherCourtierDuneVille(Browser, ville=Villes[0])
    print("test")
    while(end == 0):
        html = Browser.page_source
        soup = bs4.BeautifulSoup(html) # On donne le contenu à bs4 une fois tous les résultats affichés.
        RawCourtier = soup.find_all("div", {"class": "agent-info"})
        for i in RawCourtier:  # On "scrap" la page à la recherche des infos désirées.
            print("******")  # Print seulement pour débuggage.
            t = 0
            for f in i.find_all("p"):
                t+=1
            ListeCourtier.append({"NomduCourtier": i.find("h2").text,
                                  "Agence": i.find_all("a")[1].text.strip(),
                                  "Adresse": i.find("span",{"class":"agent-info__brokerage-address"}).text,
                                  "NumBureau": i.find_all("p")[0].a.text if (t != 0) else '',
                                  "NumMobile": i.find_all("p")[1].a.text if (t > 1) else ''}) #Nom = i.find("h2").text
        PageNumbering+=1
        try :
            Browser.find_element_by_link_text("{}".format(PageNumbering)).click()
            time.sleep(4)
        except :
            print("Semble ne pas être la! Fini.")
            end = 1
            Villes.pop(0)




    with open('list_courtier_output.csv', 'w', newline='') as output_file:
        dict_writer = csv.DictWriter(output_file, ["NomduCourtier", "Agence", "Adresse", "NumBureau", "NumMobile"])
        dict_writer.writeheader()
        dict_writer.writerows(ListeCourtier)
    #wait

    # brwser.implicitly_wait(1)
    # search_bar.send_keys(Keys.RETURN)
# print("test")