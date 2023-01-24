import datetime
import glob
import os
import time
from pathlib import Path

import dateutil.relativedelta
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from classes.quantalys import Quantalys


class Scarico():
    # TODO: aggiungi doppio click anche all'aggiornamento date
    # TODO: tutti gli aggiornamenti portali anche su scarico_completo.py

    def __repr__(self):
        return "Importa le liste complete e scarica i dati da Quantalys.it"

    def __init__(self, intermediario, t1):
        """
        Default download folder : self.directory_output_liste
        Default browser : chromium

        Arguments:
            username {str} = username dell'account
            password {str} = password dell'account
            t1 {datetime} = data di calcolo indici alla fine del mese
            directory_output_liste {WindowsPath} = percorso in cui scaricare i dati delle liste
            directory_input_liste {WindowsPath} = percorso in cui trovare i dati delle liste
        """
        self.intermediario = intermediario
        self.t1 = t1
        self.t0_3Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+3)).strftime("%d/%m/%Y") # data iniziale tre anni fa
        print(f"Tre anni fa : {self.t0_3Y}.")
        self.t0_1Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+1)).strftime("%d/%m/%Y") # data iniziale un anno fa
        print(f"Un anno fa : {self.t0_1Y}.")
        directory = Path().cwd()
        self.directory = directory
        self.directory_input_liste = self.directory.joinpath('docs', 'import_liste_into_Q')
        self.directory_output_liste = self.directory.joinpath('docs', 'export_liste_from_Q')
        if not os.path.exists(self.directory_output_liste):
            os.makedirs(self.directory_output_liste)
        chrome_options = webdriver.ChromeOptions()
        # chrome_options.add_experimental_option("detach", True) -> lascia il browser aperto dopo aver eseugito tutto il codice
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": self.directory_output_liste.__str__(),
            "download.directory_upgrade": True}
            )
        # API dove trovare il chromedriver aggiornato -> https://chromedriver.storage.googleapis.com/index.html
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        classi_a_benchmark_BPPB = {
            'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 
            'OBB_EUR_BT': '2265', 'OBB_EUR_MLT': '2264', 'OBB_EUR_CORP': '2272', 'OBB_GLOB': '2309', 'OBB_EM': '2476', 'OBB_HY': '2293'}
        classi_a_benchmark_BPL = {
            'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 'AZ_GLOB': '2318',
            'OBB_EUR_BT': '2265', 'OBB_EUR_MLT': '2264', 'OBB_EUR': '2255', 'OBB_EUR_CORP': '2272', 'OBB_GLOB': '2309', 'OBB_USA': '2490',
            'OBB_EM': '2476', 'OBB_HY': '2293'}
        classi_a_benchmark_CRV = {
            'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 'AZ_GLOB': '2318',
            'OBB_EUR_BT': '2265', 'OBB_EUR_MLT': '2264', 'OBB_EUR_CORP': '2272', 'OBB_GLOB': '2309', 'OBB_EM': '2476', 'OBB_HY': '2293'}
        classi_a_benchmark_RIPA = {
            'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 'AZ_GLOB': '2318', 
            'AZ_BIO' : '2240', 'AZ_BDC' : '2318', 'AZ_FIN' : '2716', 'AZ_AMB' : '2318', 'AZ_IMM' : '2187', 'AZ_IND' : '2175', 
            'AZ_ECO' : '2174', 'AZ_SAL' : '2178', 'AZ_SPU' : '2181', 'AZ_TEC' : '2179', 'AZ_TEL' : '2180', 'AZ_ORO' : '2318', 
            'AZ_BEAR' : '2318', 
            'OBB_EUR_BT': '2265', 'OBB_EUR_MLT': '2264', 'OBB_EUR_CORP': '2272', 'OBB_EUR': '2255', 'OBB_USA': '2490', 'OBB_JAP' : '2309', 
            'OBB_GLOB': '2309', 'OBB_EM': '2476', 'OBB_HY': '2293'}
        classi_a_benchmark_RAI = {
            'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 'AZ_GLOB': '2318', 
            'OBB_EUR_BT': '2265', 'OBB_EUR_MLT': '2264', 'OBB_EUR_CORP': '2272', 'OBB_EUR': '2255', 'OBB_USA': '2490', 
            'OBB_GLOB': '2309', 'OBB_EM': '2476', 'OBB_HY': '2293'}
        match self.intermediario:
            case 'BPPB':
                self.classi_a_benchmark = classi_a_benchmark_BPPB
            case 'BPL':
                self.classi_a_benchmark = classi_a_benchmark_BPL
            case 'CRV':
                self.classi_a_benchmark = classi_a_benchmark_CRV
            case 'RIPA':
                self.classi_a_benchmark = classi_a_benchmark_RIPA
            case 'RAI':
                self.classi_a_benchmark = classi_a_benchmark_RAI

    def export(self):
        """
        Carica le liste in quantalys.it, scarica gli indicatori pertinenti ed esporta un file csv.
        Rinomina il file con nomi in successione relativi alla macrocategoria.
        """
        q = Quantalys()

        # Accesso a Quantalys
        q.connessione(self.driver)
        
        # Log in
        q.login(self.driver, 'Avicario', 'AVicario123')

        # Il processo parte se la cartella di download è vuota
        while len(os.listdir(self.directory_output_liste)) != 0:
            print(f"\nCi sono dei file presenti nella cartella di download: {glob.glob(self.directory_output_liste.__str__()+'/*')}\n")
            _ = input('cancella i file prima di proseguire, poi premi enter\n')
        
        directory = self.directory_input_liste
        elapsed_time = []
        liste_completate = 0
        file_totali = len(os.listdir(directory))
        for filename in os.listdir(directory):
            file_scaricati = len(os.listdir(self.directory_output_liste))
            start = time.perf_counter()
            print(f"\nCaricamento lista {filename}...\n")
            
            # Lista
            id_lista, numero_fondi = q.to_liste(self.driver, filename[:-4], self.directory_input_liste, filename)

            # Confronto
            NUM_MAX_FONDI_CONFRONTO_DIRETTO = 1 # 2000
            # Se il numero di fondi caricati è inferiore a NUM_MAX_FONDI_CONFRONTO_DIRETTO usa l'API nella scheda liste
            if int(numero_fondi) < NUM_MAX_FONDI_CONFRONTO_DIRETTO:
                self.driver.find_element(by=By.XPATH, value='//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label').click() # Seleziona tutto
                time.sleep(2) # Necessario, va troppo veloce.
                self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[2]/div[3]/div/button[3]').click() # Confronta
            # altrimenti passa da fondi -> confronto
            else:
                q.to_confronto(self.driver, id_lista)
            # personalizzato
            WebDriverWait(self.driver, 360).until(EC.presence_of_element_located((By.LINK_TEXT, 'Personalizzato'))).click()
            
            # Seleziona indicatori
            WebDriverWait(self.driver, 180).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]')
                )
            )
            self.driver.find_element(by=By.TAG_NAME, value='body').send_keys(Keys.PAGE_DOWN)
            time.sleep(0.5) # altrimenti non c'entra il doppio click su 'aggiorna'
            if filename.startswith('AZ') or filename.startswith('OBB'):
                self.aggiungi_indicatori_v2('Codice ISIN', 'Nome', 'Valuta', 'Information ratio da data a data', 'TEV da data a data')
            elif filename.startswith('FLEX') or filename.startswith('BIL') or filename.startswith('COMM') or filename.startswith('PERF'):
                self.aggiungi_indicatori_v2('Codice ISIN', 'Nome', 'Valuta', 'Sortino ratio da data a data', 'DSR da data a data')
            elif filename.startswith('OPP'):
                self.aggiungi_indicatori_v2('Codice ISIN', 'Nome', 'Valuta', 'Sharpe ratio da data a data', 'Volatilità da data a data')
            elif filename.startswith('LIQ'):
                self.aggiungi_indicatori_v2('Codice ISIN', 'Nome', 'Valuta', 'Perf Ann. da data a data', 'Volatilità da data a data')

            # Aggiungi benchmark
            if filename[:-6] in self.classi_a_benchmark.keys():
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_rdIndiceRefTousFonds"]').click()
                select = Select(self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_cmbIndiceRef_Comp"]'))
                select.select_by_value(self.classi_a_benchmark[filename[:-6]])
                # time.sleep(1.5) # troppo veloce
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_bntProPlusRafraichir"]'))) # Aggiorna benchmark
                element = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_bntProPlusRafraichir"]')
                ActionChains(self.driver).double_click(on_element=element).perform()
                # self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_bntProPlusRafraichir"]').click()
            else:
                # time.sleep(1.5) # troppo veloce
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_bntProPlusRafraichir"]'))) # Aggiorna benchmark
                element = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_bntProPlusRafraichir"]')
                ActionChains(self.driver).double_click(on_element=element).perform()
                # self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_bntProPlusRafraichir"]').click()
            try:
                loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
                WebDriverWait(self.driver, 5).until(EC.visibility_of(loading_img)) # da 10 a 5 perché se la lista è piccola lo devo mandare avanti a mano
            except:
                input('premi enter se ha caricato')

            # Aggiorna date a 3 anni
            # La rotellina di caricamento non viene intercettata da selenium quando i fondi sono pochi, perché compare solo per qualche istante.
            # È necessario aspettare alcuni secondi quando la lista è breve piuttosto che affidarsi alla rotellina.
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            finally:
                data_di_avvio_3_anni = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_dtDebut_txtDatePicker"]')
                data_di_avvio_3_anni.clear()
                data_di_avvio_3_anni.send_keys(self.t0_3Y) 
                data_di_fine = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_dtFin_txtDatePicker"]')
                data_di_fine.clear()
                data_di_fine.send_keys(self.t1)
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_lnkRefresh"]').click() # Aggiorna date
            try:
                loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
                WebDriverWait(self.driver, 5).until(EC.visibility_of(loading_img)) # da 10 a 5 perché se la lista è piccola lo devo mandare avanti a mano
            except:
                input('premi enter se ha caricato')
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            finally:
                WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_btnExportCSV"]')))
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_btnExportCSV"]').click()
                try:
                    loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
                    WebDriverWait(self.driver, 5).until(EC.visibility_of(loading_img)) # da 10 a 5 perché se la lista è piccola lo devo mandare avanti a mano
                except:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
                    # time.sleep(1)
                
            # try:
            #     WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            # except TimeoutException:
            #     pass
            # finally:
            #     data_di_avvio_3_anni = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_dtDebut_txtDatePicker"]')
            #     data_di_avvio_3_anni.clear()
            #     data_di_avvio_3_anni.send_keys(self.t0_3Y) 
            #     data_di_fine = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_dtFin_txtDatePicker"]')
            #     data_di_fine.clear()
            #     data_di_fine.send_keys(self.t1)
            #     self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_lnkRefresh"]').click() # Aggiorna date
            #     loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
            #     WebDriverWait(self.driver, 10).until(EC.visibility_of(loading_img))


            # Salva il file con nome a tre anni
            # try:
            #     WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            # except TimeoutException:
            #     pass
            # finally:
            #     WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_btnExportCSV"]')))
            #     self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_btnExportCSV"]').click()
            #     try:
            #         loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
            #         WebDriverWait(self.driver, 600).until(EC.visibility_of(loading_img))
            #     except TimeoutException:
            #         pass
            #     finally:
            #         WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            #         time.sleep(1)

            # Rinomina file
            while len(os.listdir(self.directory_output_liste)) == file_scaricati:
                time.sleep(1)
            time.sleep(1.5)
            list_of_files = glob.glob(self.directory_output_liste.__str__() + '/*')
            latest_file = max(list_of_files, key=os.path.getctime)
            os.rename(latest_file, self.directory_output_liste.joinpath(filename[:-4]+'_3Y.csv'))

            # Aggiorna date 1 anno
            try:
                WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_dtDebut_txtDatePicker"]')))
            except TimeoutException:
                pass
            finally:
                data_di_avvio_1_anno = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_dtDebut_txtDatePicker"]')
                data_di_avvio_1_anno.clear()
                data_di_avvio_1_anno.send_keys(self.t0_1Y)
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_lnkRefresh"]').click() # Aggiorna date
            try:
                loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
                WebDriverWait(self.driver, 5).until(EC.visibility_of(loading_img)) # da 10 a 5 perché se la lista è piccola lo devo mandare avanti a mano
            except:
                input('premi enter se ha caricato')
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            finally:
                WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_btnExportCSV"]')))
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_btnExportCSV"]').click()
                try:
                    loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
                    WebDriverWait(self.driver, 5).until(EC.visibility_of(loading_img)) # da 10 a 5 perché se la lista è piccola lo devo mandare avanti a mano
                except:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))

            # Salva il file con nome ad un anno
            # try:
            #     WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            # except TimeoutException:
            #     pass
            # finally:
            #     WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_btnExportCSV"]')))
            #     self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_btnExportCSV"]').send_keys(Keys.ENTER)
            #     try:
            #         loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
            #         WebDriverWait(self.driver, 600).until(EC.visibility_of(loading_img))
            #     except TimeoutException:
            #         pass
            #     finally:
            #         WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            #         time.sleep(1)
            
            # Rinomina file
            while len(os.listdir(self.directory_output_liste)) == file_scaricati:
                time.sleep(1)
            time.sleep(1.5)
            list_of_files = glob.glob(self.directory_output_liste.__str__() + '/*')
            latest_file = max(list_of_files, key=os.path.getctime)
            os.rename(latest_file, self.directory_output_liste.joinpath(filename[:-4]+'_1Y.csv'))

            end = time.perf_counter()
            elapsed_time.append(end - start)
            print(f"Elapsed time for {filename}: ", end - start, 'seconds')
            print(f"\nAverage elapsed time: {sum(elapsed_time)/len(elapsed_time)}.")
            liste_completate += 1
            print(f"\nTempo previsto alla fine: {datetime.timedelta(seconds=(sum(elapsed_time)/len(elapsed_time))*(file_totali-liste_completate))}")
        
        self.driver.close()


if __name__ == '__main__':
    start = time.perf_counter()
    _ = Scarico(intermediario='RAI', t1='31/12/2022')
    # _.accesso_a_quantalys()
    # _.login()
    _.export()
    end = time.perf_counter()
    print("Elapsed time: ", end - start, 'seconds')