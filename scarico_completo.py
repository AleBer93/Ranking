import glob
import os
import re
import time
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (ElementNotInteractableException,
                                        NoSuchElementException,
                                        TimeoutException)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


class ScaricoCompleto():

    def __str__(self):
        return "Importa le liste complete e scarica i dati da Quantalys.it"

    def __init__(self, username='AVicario', password='AVicario123'):
        """
        Default download folder : self.directory_output_liste_complete
        Default browser : chromium
        
        Parameters:
            username {str} = username dell'account
            password {str} = password dell'account
            directory_output_liste_complete {WindowsPath} = percorso in cui scaricare i dati delle liste complete
            directory_input_liste_complete {WindowsPath} = percorso in cui trovare i dati delle liste complete
        """
        # Alt account username='Pomante', password='Pomante22'
        self.username = username
        self.password = password
        directory = Path().cwd()
        self.directory = directory
        self.directory_output_liste_complete = self.directory.joinpath('docs', 'export_liste_complete_from_Q')
        # self.directory_output_liste_complete = directory_output_liste_complete
        self.directory_input_liste_complete = self.directory.joinpath('docs', 'import_liste_complete_into_Q')
        # self.directory_input_liste_complete = directory_input_liste_complete
        if not os.path.exists(self.directory_output_liste_complete):
            os.makedirs(self.directory_output_liste_complete)
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory" : self.directory_output_liste_complete.__str__(),
            "download.directory_upgrade" : True}
            )
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
    
    def accesso_a_quantalys(self):
        """
        Accede a quantalys.it con chromium. Imposta come cartella di download il percorso in self.directory_output_liste_complete
        e massimizza la finestra.
        """
        print('\n...connessione a Quantalys...')
        self.driver.get("https://www.quantalys.it")
        self.driver.maximize_window()

    def login(self):
        """
        Chiude l'alert dei cookies.
        Accede all'account con username=self.username e password=self.password.
        """
         # Chiudi i cookies
        try:
            time.sleep(1)
            self.driver.find_element(by=By.XPATH, value='//*[@id="tarteaucitronPersonalize2"]').click() # Cookies
        except NoSuchElementException:
            pass
        # Connessione
        try:
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnConnexion"]'))) # Connessione
        except TimeoutException:
            pass
        else:
            self.driver.find_element(by=By.XPATH, value='//*[@id="btnConnexion"]').click()
        # Username e password
        try:
            WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="inputLogin"]'))) 
        except TimeoutException:
            pass
        else:
            time.sleep(0.5)
            self.driver.find_element(by=By.XPATH, value='//*[@id="inputLogin"]').send_keys(self.username)
            self.driver.find_element(by=By.XPATH, value='//*[@id="inputPassword"]').send_keys(self.password,Keys.ENTER)
            self.driver.find_element(by=By.XPATH, value='//*[@id="btnConnecter"]').click()

    def get_data_from_table(self, driver, table, num_pages):
        """
        Ricava i dati da una tabella html contenuta in pagine multiple

        Arguments:
            driver {str} = driver che inietta il codice nel browser
            table {url} = indirizzo url che porta alla tabella (full x-path)
            num_pages {int} =  numero di pagine in cui è divisa la tabella

        Return
            df {dataframe} = Dataframe contenente i dati estratti
        """
        element = driver.find_element(By.XPATH, table).get_attribute('outerHTML')
        df = pd.read_html(element)[0]
        for page in range(2, num_pages+1):
            # Individua il nome del primo elemento
            nome_primo_fondo = self.driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr[1]/td[2]').text
            # Quantalys non mette tutti gli li. Li aggiunge alla mano
            # a = self.driver.find_element(by=By.CSS_SELECTOR, value=list_class+' li:nth-child('+str(page)+') a')
            # L'unico modo di individuare l'anchor link che mi serve è selezionare l'anchor link in base al numero progressivo corrispondente alla pagina
            num_pagina = self.driver.find_element(by=By.LINK_TEXT, value=str(page))
            num_pagina.click()
            # Attendi che il tag li abbia l'attributo active (non funziona per il motivo nel commento sopra)
            # WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, list_class+' li:nth-child('+str(page)+')'))).get_attribute('active')
            # Attendi che il nome del primo elemento della nuova tabella sia diverso dal nome del primo elemento della tabella aggiornata
            WebDriverWait(self.driver, 20).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr[1]/td[2]'), nome_primo_fondo))
            # Scarica il nuovo dataframe
            element2 = driver.find_element(By.XPATH, table).get_attribute('outerHTML')
            df2 = pd.read_html(element2)[0]
            # Allegalo in coda al primo (df)
            df = pd.concat([df, df2], ignore_index=True)
        return df

    def export(self):
        """
        Carica le liste in quantalys.it ed esporta un file csv completo.
        Rinomina il file con nomi in successione.
        """
        # Il processo parte se la cartella di download è vuota
        while len(os.listdir(self.directory_output_liste_complete)) != 0:
            print(f"\nCi sono dei file presenti nella cartella di download: {glob.glob(self.directory_output_liste_complete.__str__()+'/*')}\n")
            _ = input('cancella i file prima di proseguire, poi premi enter\n')
        
        for filename in os.listdir(self.directory_input_liste_complete):
            file_totali = len(os.listdir(self.directory_output_liste_complete))
            if filename.startswith('lista_completa'):
                print(f'caricamento {filename}...')
                # Logo quantalys
                # try:
                #     WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="position-menu-quantalys"]/div/div[1]/a/img')))
                # except TimeoutException:
                #     pass
                # Liste
                try:
                    liste = self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Liste')
                    liste.click()
                except:
                    try:
                        # time.sleep(0.5)
                        WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Tools')))
                    except TimeoutException:
                        pass
                    finally:
                        self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Tools').click()

                    try:
                        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Liste')))
                    except TimeoutException:
                        pass
                    finally:
                        self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Liste').click()
                # Nuova lista
                try:
                    WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[2]/div/div[2]/div[1]/button')))
                except TimeoutException:
                    pass
                finally:
                    time.sleep(1)
                    self.driver.find_element(by=By.NAME, value='new').click()
                # Nome lista
                try:
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, 'nom'))) # Nome
                except TimeoutException:
                    pass
                finally:
                    time.sleep(1)
                    self.driver.find_element(by=By.NAME, value="nom").send_keys(filename[:-4], Keys.TAB, Keys.TAB, Keys.ENTER) # Conferma
                # Importa prodotti
                try:
                    WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]'))) # Importa dei prodotti
                except TimeoutException:
                    pass
                finally:
                    time.sleep(0.5)
                    self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]').click()
                # Scegli un file da importare
                try:
                    WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.NAME, 'file'))) # Seleziona lista da importare
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.NAME, value="file").send_keys(self.directory_input_liste_complete.joinpath(filename).__str__()) # Directory
                # Importa lista
                try:
                    WebDriverWait(self.driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="importForm"]/button'))) # Importa
                except TimeoutException:
                    pass
                finally:
                    time.sleep(1) # Necessario, va troppo veloce ed esporta liste vuote
                    self.driver.find_element(by=By.XPATH, value='//*[@id="importForm"]/button').click()
                    # La riga sotto è prbobalimente inutile
                    WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[3]/div[2]')))
                # Esporta
                try:
                    WebDriverWait(self.driver, 120).until(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr/td'), 'Nessun dato disponibile'))
                    WebDriverWait(self.driver, 120).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr/td'), 'Nessun dato disponibile'))
                    # WebDriverWait(self.driver,60).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[3]/div[2]'), '0 elementi'))
                except TimeoutException:
                    pass
                else:
                    WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label'))) # Seleziona tutto
                    self.driver.find_element(by=By.XPATH, value='//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label').click()
                    time.sleep(1)
                    WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[1]/div/div[2]/div/button'))) # Esporta
                    time.sleep(1.5)
                    self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[1]/div/div[2]/div/button').click()
                # Raccogli i fondi non importati perché non presenti in piattaforma
                # try:
                #     prodotti_non_presenti = self.driver.find_element(by=By.XPATH, value='//*[@id="NotImportedData"]/p').get_attribute("textContent")
                #     file_prodotti_non_presenti = open(self.directory.joinpath('docs', "prodotti_non_presenti.txt"), 'a')
                #     prodotti_non_presenti = prodotti_non_presenti.split(sep=',')
                #     prodotti_non_presenti = [item.split('(') for item in prodotti_non_presenti]
                #     prodotti_non_presenti = [[element[0], element[1][:-1]] for element in prodotti_non_presenti]
                #     for element in prodotti_non_presenti:
                #         file_prodotti_non_presenti.write(element[0]+' '+element[1]+'\n')
                #     file_prodotti_non_presenti.close()
                # except:
                #     pass
                # Esporta CSV completo
                try:
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[1]/div/div[2]/div/ul/li[4]/a'))) # CSV completo
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[1]/div/div[2]/div/ul/li[4]/a').click()
                # Rinomina file
                while len(os.listdir(self.directory_output_liste_complete)) == file_totali:
                    time.sleep(1)
                time.sleep(1.5)
                list_of_files = glob.glob(self.directory_output_liste_complete.__str__() + '/*')
                latest_file = max(list_of_files, key=os.path.getctime)
                os.rename(latest_file, self.directory_output_liste_complete.joinpath(filename[:-4]+'.csv'))

                # Scarico articolo SFDR
                self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[2]/ul/li[7]/a').click()
                try:
                    WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0"]/thead/tr/th[4]'))).text == 'SFDR'
                except TimeoutException:
                    pass
                finally:
                    # Numero elementi arrotondati sempre in eccesso
                    totale_fondi_lista = self.driver.find_element(by=By.XPATH, value='//*[@id="DataTables_Table_0_info"]').text.replace(',','')
                    print(f'{totale_fondi_lista}\n')
                    num_fondi_regex = re.compile(r'\d(\d)?(\d)?(\d)?')
                    mo = num_fondi_regex.search(totale_fondi_lista)
                    numero_fondi = mo.group()
                    # Calcolo numero pagine
                    NUMERO_FONDI_PER_PAGINA = 100
                    # numero_pagine = int(round(int(numero_fondi) / NUMERO_FONDI_PER_PAGINA, 0) + 1) - 1 # se ci sono 115 fondi non funziona
                    numero_pagine = int(int(numero_fondi) / NUMERO_FONDI_PER_PAGINA) + 1
                    print('numero pagine:', numero_pagine)
                    # Scarica l'articolo SFDR
                    df = self.get_data_from_table(self.driver, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table', numero_pagine)
                    df.to_csv(self.directory.joinpath('docs','sfdr'+filename[-6:-4]+'.csv'), sep=";", decimal=',', index=False)
                        
        self.driver.close()


if __name__ == '__main__':
    start = time.perf_counter()
    _ = ScaricoCompleto()
    _.accesso_a_quantalys()
    _.login()
    _.export()
    end = time.perf_counter()
    print("Elapsed time: ", round(end - start, 2), 'seconds')
