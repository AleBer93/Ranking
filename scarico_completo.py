import glob
import os
import time
from pathlib import Path

from classes.quantalys import Quantalys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


class ScaricoCompleto():

    def __repr__(self):
        return "Importa le liste complete e scarica i dati da Quantalys.it"

    def __init__(self):
        """
        Default download folder : self.directory_output_liste_complete
        Default browser : chromium
        
        Parameters:
            directory_output_liste_complete {WindowsPath} = percorso in cui scaricare i dati delle liste complete
            directory_input_liste_complete {WindowsPath} = percorso in cui trovare i dati delle liste complete
        """
        directory = Path().cwd()
        self.directory = directory
        self.directory_input_liste_complete = self.directory.joinpath('docs', 'import_liste_complete_into_Q')
        self.directory_output_liste_complete = self.directory.joinpath('docs', 'export_liste_complete_from_Q')
        if not os.path.exists(self.directory_output_liste_complete):
            os.makedirs(self.directory_output_liste_complete)
        self.directory_sfdr = self.directory.joinpath('docs', 'sfdr')
        if not os.path.exists(self.directory_sfdr):
            os.makedirs(self.directory_sfdr)
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory" : self.directory_output_liste_complete.__str__(),
            "download.directory_upgrade" : True}
        )
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
    
    def export(self):
        """
        Carica le liste in quantalys.it ed esporta un file csv completo.
        Rinomina il file con nomi in successione.
        """
        q = Quantalys()

        # Accesso a Quantalys
        q.connessione(self.driver)
        
        # Log in
        q.login(self.driver, 'Avicario', 'AVicario123')

        # Il processo parte se la cartella di download è vuota
        while len(os.listdir(self.directory_output_liste_complete)) != 0:
            print(f"\nCi sono dei file presenti nella cartella di download: {glob.glob(self.directory_output_liste_complete.__str__()+'/*')}\n")
            _ = input('cancella i file prima di proseguire, poi premi enter\n')
        
        for filename in os.listdir(self.directory_input_liste_complete):
            file_totali = len(os.listdir(self.directory_output_liste_complete))
            if filename.startswith('lista_completa'):
                print(f'caricamento {filename}...')

                # Carica lista
                id_lista, numero_fondi = q.carica_lista(self.driver, filename[:-4], self.directory_input_liste_complete, filename)

                # Esporta
                # seleziona tutto
                WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label'))
                ).click()
                time.sleep(1.5)
                # esporta
                WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[1]/div/div[2]/div/button'))
                ).click()

                # Esporta CSV completo
                WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[1]/div/div[2]/div/ul/li[4]/a'))
                ).click()

                # Rinomina file
                while len(os.listdir(self.directory_output_liste_complete)) == file_totali:
                    time.sleep(1)
                time.sleep(1.5)
                list_of_files = glob.glob(self.directory_output_liste_complete.__str__() + '/*')
                latest_file = max(list_of_files, key=os.path.getctime)
                os.rename(latest_file, self.directory_output_liste_complete.joinpath(filename[:-4]+'.csv'))

                # Scarico articolo SFDR
                WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="quantasearch"]/div[2]/ul/li[7]/a'))).click()
                WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="DataTables_Table_0"]/thead/tr/th[4]'))).text == 'SFDR'
                # calcolo numero pagine
                NUMERO_FONDI_PER_PAGINA = 100
                # numero_pagine = int(round(int(numero_fondi) / NUMERO_FONDI_PER_PAGINA, 0) + 1) - 1 # se ci sono 115 fondi non funziona
                numero_pagine = int(int(numero_fondi) / NUMERO_FONDI_PER_PAGINA) + 1
                print('numero pagine:', numero_pagine)
                # scarica l'articolo
                df = q.get_data_from_table(self.driver, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table', numero_pagine)
                df.to_csv(self.directory_sfdr.joinpath('sfdr'+filename[-6:-4]+'.csv'), sep=";", decimal=',', index=False)
                        
        self.driver.close()


if __name__ == '__main__':
    start = time.perf_counter()
    _ = ScaricoCompleto()
    _.export()
    end = time.perf_counter()
    print("Elapsed time: ", round(end - start, 2), 'seconds')