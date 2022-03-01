import glob
import os
import time
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import (ElementNotInteractableException,
                                        NoSuchElementException,
                                        TimeoutException)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


class ScaricoCompleto():

    def __str__(self):
        return "Importa le liste complete e scarica i dati da Quantalys.it"

    def __init__(self, username='AVicario', password='AVicario123'):
        """
        Default download folder : self.directory_output_liste_complete
        Default browser : chromium
        
        Parameters:
            username(str) = username dell'account
            password(str) = password dell'account
            directory_output_liste_complete = percorso in cui scaricare i dati delle liste complete
            directory_input_liste_complete = percorso in cui trovare i dati delle liste complete
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
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.implicitly_wait(5)
    
    def accesso_a_quantalys(self):
        """
        Accede a quantalys.it con chromium. Imposta come cartella di download il percorso in self.directory_output_liste_complete
        e massimizza la finestra.
        """
        print('\n...connessione a Quantalys...')
        self.driver.get("http://www.quantalys.it")
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
            # time.sleep(0.5)
            self.driver.find_element(by=By.XPATH, value='//*[@id="inputLogin"]').send_keys(self.username)
            self.driver.find_element(by=By.XPATH, value='//*[@id="inputPassword"]').send_keys(self.password,Keys.ENTER)

    def export(self):
        """
        Carica le liste in quantalys.it ed esporta un file csv completo.
        Rinomina il file con nomi in successione.
        """
        # Il processo parte se la cartella di download è vuota
        while len(os.listdir('./export_liste_complete_from_Q')) != 0:
            print(f"\nCi sono dei file presenti nella cartella di download: {glob.glob(self.directory_output_liste_complete.__str__()+'/*')}\n")
            _ = input('cancella i file prima di proseguire, poi premi enter\n')
        
        for filename in os.listdir(self.directory_input_liste_complete):
            file_totali = len(os.listdir(self.directory_output_liste_complete))
            if filename.startswith('lista_completa'):
                print(f'caricamento {filename}...')
                # Logo quantalys
                try:
                    WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="position-menu-quantalys"]/div/div[1]/a/img')))
                except TimeoutException:
                    pass
                # Liste
                try:
                    liste = self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Liste')
                    liste.click()
                except:
                    try:
                        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Tools')))
                    except TimeoutException:
                        pass
                    finally:
                        self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Tools').click()

                    try:
                        WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Liste')))
                    except TimeoutException:
                        pass
                    finally:
                        self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Liste').click()
                # Nuova lista
                try:
                    WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[2]/div/div[2]/div[1]/button')))
                except TimeoutException:
                    pass
                finally:
                    time.sleep(1)
                    self.driver.find_element(by=By.NAME, value='new').click()
                # Nome lista
                try:
                    WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.NAME, 'nom'))) # Nome
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.NAME, value="nom").send_keys(filename[:-4], Keys.TAB, Keys.TAB, Keys.ENTER) # Conferma
                # Importa prodotti
                try:
                    WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]'))) # Importa dei prodotti
                except TimeoutException:
                    pass
                finally:
                    time.sleep(0.5)
                    self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]').click()
                # Scegli un file da importare
                try:
                    WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.NAME, 'file'))) # Seleziona lista da importare
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.NAME, value="file").send_keys(self.directory_input_liste_complete.joinpath(filename).__str__()) # Directory
                # Importa lista
                try:
                    WebDriverWait(self.driver, 7).until(EC.presence_of_element_located((By.XPATH, '//*[@id="importForm"]/button'))) # Importa
                except TimeoutException:
                    pass
                finally:
                    time.sleep(0.5) # Necessario, va troppo veloce ed esporta liste vuote
                    self.driver.find_element(by=By.XPATH, value='//*[@id="importForm"]/button').click()
                    WebDriverWait(self.driver,60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[3]/div[2]')))
                # Esporta
                try:
                    WebDriverWait(self.driver,120).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr/td'), 'Nessun dato disponibile'))
                    # WebDriverWait(self.driver,60).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[3]/div[2]'), '0 elementi'))
                except TimeoutException:
                    pass
                else:
                    WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label'))) # Seleziona tutto
                    self.driver.find_element(by=By.XPATH, value='//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label').click()
                    WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[1]/div/div[2]/div/button'))) # Esporta
                    time.sleep(1.5)
                    self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[1]/div/div[2]/div/button').click()
                # Esporta CSV completo
                try:
                    WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[1]/div/div[2]/div/ul/li[4]/a'))) # CSV completo
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

        self.driver.close()


if __name__ == '__main__':
    start = time.time()
    _ = ScaricoCompleto()
    _.accesso_a_quantalys()
    _.login()
    _.export()
    end = time.time()
    print("Elapsed time: ", end - start, 'seconds')