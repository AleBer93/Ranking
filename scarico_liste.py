import datetime
import glob
import os
import time
from pathlib import Path

import dateutil.relativedelta
from classes.quantalys import Quantalys
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


class Scarico():

    def __init__(self, intermediario):
        """
        Arguments:
            intermediario {str} - intermediario a cui è destinata l'analisi
        """
        # Input
        self.intermediario = intermediario
        # Dates
        with open('docs/t1.txt') as f:
            t1 = f.read()
        t1 = datetime.datetime.strptime(t1, '%Y-%m-%d').strftime("%d/%m/%Y")
        self.t1 = t1
        self.t0_3Y = (
            datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+3)
        ).strftime("%d/%m/%Y") # data iniziale tre anni fa
        self.t0_1Y = (
            datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+1)
        ).strftime("%d/%m/%Y") # data iniziale un anno fa
        print(f'Data odierna: {self.t1}')
        print(f'Un anno fa : {self.t0_1Y}')
        print(f'Tre anni fa : {self.t0_3Y}')
        # Directories
        directory = Path().cwd()
        self.directory = directory
        self.directory_input_liste = self.directory.joinpath('docs', 'import_liste_into_Q')
        self.directory_output_liste = self.directory.joinpath('docs', 'export_liste_from_Q')
        if not os.path.exists(self.directory_output_liste):
            os.makedirs(self.directory_output_liste)
        # Browser options
        chrome_options = webdriver.ChromeOptions()
        # chrome_options.add_experimental_option("detach", True) -> lascia il browser aperto dopo aver eseugito tutto il codice
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": self.directory_output_liste.__str__(),
            "download.directory_upgrade": True}
            )
        # API dove trovare il chromedriver aggiornato -> https://chromedriver.storage.googleapis.com/index.html
        service = Service()
        self.driver = webdriver.Chrome(service=service, options=chrome_options)

        # Intermediario
        match intermediario:
            case 'BPPB':
                self.IR_TEV = [
                    'AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'OBB_EUR_BT', 'OBB_EUR_MLT', 'OBB_EUR_CORP', 'OBB_GLOB',
                    'OBB_EM', 'OBB_HY'
                ]
                self.SOR_DSR = ['FLEX_BVOL', 'FLEX_MAVOL']
                self.SHA_VOL = ['OPP']
                self.PER_VOL = ['LIQ']
            case 'BPL':
                self.IR_TEV = [
                    'AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_EUR_BT', 'OBB_EUR_MLT', 'OBB_EUR', 'OBB_EUR_CORP',
                    'OBB_GLOB', 'OBB_USA', 'OBB_EM', 'OBB_HY'
                ]
                self.SOR_DSR = ['BIL_MBVOL', 'BIL_AVOL', 'FLEX_PR', 'FLEX_DIN']
                self.SHA_VOL = ['OPP']
                self.PER_VOL = ['LIQ', 'LIQ_FOR']
            case 'CRV':
                self.IR_TEV = [
                    'AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_EUR_BT', 'OBB_EUR_MLT', 'OBB_EUR_CORP', 'OBB_GLOB',
                    'OBB_EM', 'OBB_HY',
                ]
                self.SOR_DSR = ['FLEX_PR', 'FLEX_DIN']
                self.SHA_VOL = ['OPP']
                self.PER_VOL = ['LIQ']
            case 'RIPA':                
                self.IR_TEV = [
                    'OBB_EUR_BT', 'OBB_EUR_MLT', 'OBB_EUR', 'OBB_EUR_CORP', 'OBB_GLOB', 'OBB_USA', 'OBB_JAP', 'OBB_EM', 'OBB_HY', 
                    'AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'AZ_BIO', 'AZ_BDC', 'AZ_FIN', 'AZ_AMB', 'AZ_IMM', 'AZ_IND', 
                    'AZ_ECO', 'AZ_SAL', 'AZ_SPU', 'AZ_TEC', 'AZ_TEL', 'AZ_ORO', 'AZ_BEAR', 'FLEX_PR', 'FLEX_DIN', 
                ]
                self.SOR_DSR = ['COMM', 'PERF_ASS']
                self.SHA_VOL = []
                self.PER_VOL = ['LIQ']
            case 'RAI':
                self.IR_TEV = [
                    'OBB_EUR_BT', 'OBB_EUR_MLT', 'OBB_EUR', 'OBB_USA', 'OBB_EUR_CORP', 'OBB_GLOB', 'OBB_EM', 'OBB_HY', 
                    'AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 
                ]
                self.SOR_DSR = ['BIL_PR', 'BIL_EQ', 'BIL_AGG', 'FLEX_PR', 'FLEX_DIN']
                self.SHA_VOL = ['OPP']
                self.PER_VOL = ['LIQ', 'LIQ_FOR']
            case _:
                print('specifica un intermediario')
                quit()

        self.benchmarks = {
            'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 'AZ_GLOB': '2318',
            'AZ_BIO' : '2240', 'AZ_BDC' : '2318', 'AZ_FIN' : '2716', 'AZ_AMB' : '2318', 'AZ_IMM' : '2187',
            'AZ_IND' : '2175', 'AZ_ECO' : '2174', 'AZ_SAL' : '2178', 'AZ_SPU' : '2181', 'AZ_TEC' : '2179',
            'AZ_TEL' : '2180', 'AZ_ORO' : '2318', 'AZ_BEAR' : '2318',
            'OBB_EUR_BT': '2265', 'OBB_EUR_MLT': '2264', 'OBB_EUR': '2255', 'OBB_EUR_CORP': '2272', 'OBB_GLOB': '2309',
            'OBB_USA': '2490', 'OBB_JAP' : '2309', 'OBB_EM': '2476', 'OBB_HY': '2293',
            'FLEX_PR': '2706', 'FLEX_DIN': '2707', 
        }

    def export(self):
        """
        Carica le liste in quantalys.it, scarica gli indicatori pertinenti ed esporta un file csv.
        Rinomina il file con nomi in successione relativi alla macrocategoria.
        """

        # Il processo parte se la cartella di download è vuota
        while len(os.listdir(self.directory_output_liste)) != 0:
            print(f"\nCi sono dei file presenti nella cartella di download: {glob.glob(self.directory_output_liste.__str__()+'/*')}\n")
            _ = input('cancella i file prima di proseguire, poi premi enter\n')
        
        q = Quantalys()

        # Accesso a Quantalys
        q.connessione(self.driver)

        # Cookies
        q.cookies(self.driver)
        
        # Log in
        q.login(self.driver, 'Avicario', 'AVicario123')

        # Inizializzazione variabili da usare nel ciclo
        elapsed_time = []
        liste_completate = 0
        file_totali = len(os.listdir(self.directory_input_liste))

        for filename in os.listdir(self.directory_input_liste):
            file_scaricati = len(os.listdir(self.directory_output_liste))
            start = time.perf_counter()
            print(f"\nCaricamento lista {filename}...\n")
            
            # Carica lista
            id_lista, numero_fondi = q.carica_lista(self.driver, filename[:-4], self.directory_input_liste, filename)

            # Confronto
            NUM_MAX_FONDI_CONFRONTO_DIRETTO = 1 # in realtà 2000 ma passando dalla scheda lista è tutto confusionario
            # Se il numero di fondi caricati è inferiore a NUM_MAX_FONDI_CONFRONTO_DIRETTO usa l'API nella scheda liste
            if int(numero_fondi) < NUM_MAX_FONDI_CONFRONTO_DIRETTO:
                self.driver.find_element(by=By.XPATH, value='//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label').click() # Seleziona tutto
                time.sleep(2) # Necessario, va troppo veloce.
                self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[2]/div[3]/div/button[3]').click() # Confronta
            # altrimenti passa da fondi -> confronto
            else:
                q.confronta_lista(self.driver, id_lista)
            # personalizzato
            WebDriverWait(self.driver, 360).until(EC.presence_of_element_located((By.LINK_TEXT, 'Personalizzato'))).click()
            
            # Seleziona indicatori
            WebDriverWait(self.driver, 180).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]')
                )
            )
            self.driver.find_element(by=By.TAG_NAME, value='body').send_keys(Keys.PAGE_DOWN)
            time.sleep(0.5) # altrimenti non centra il doppio click su 'aggiorna'
            if filename[:-6] in self.IR_TEV:
                q.aggiungi_indicatori_v2(
                    self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Information ratio da data a data', 'TEV da data a data'
                )
            elif filename[:-6] in self.SOR_DSR:
                q.aggiungi_indicatori_v2(
                    self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Sortino ratio da data a data', 'DSR da data a data'
                )
            elif filename[:-6] in self.SHA_VOL:
                q.aggiungi_indicatori_v2(
                    self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Sharpe ratio da data a data', 'Volatilità da data a data'
                )
            elif filename[:-6] in self.PER_VOL:
                q.aggiungi_indicatori_v2(
                    self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Perf Ann. da data a data', 'Volatilità da data a data'
                )

            # if filename.startswith('AZ') or filename.startswith('OBB'):
            #     q.aggiungi_indicatori_v2(
            #         self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Information ratio da data a data', 'TEV da data a data'
            #     )
            # elif filename.startswith('FLEX') or filename.startswith('BIL') or filename.startswith('COMM') or filename.startswith('PERF'):
            #     q.aggiungi_indicatori_v2(
            #         self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Sortino ratio da data a data', 'DSR da data a data'
            #     )
            # elif filename.startswith('OPP'):
            #     q.aggiungi_indicatori_v2(
            #         self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Sharpe ratio da data a data', 'Volatilità da data a data'
            #     )
            # elif filename.startswith('LIQ'):
            #     q.aggiungi_indicatori_v2(
            #         self.driver, 'Codice ISIN', 'Nome', 'Valuta', 'Perf Ann. da data a data', 'Volatilità da data a data'
            #     )

            # Aggiungi benchmark
            # if filename[:-6] in self.benchmarks.keys():
            if filename[:-6] in self.IR_TEV:
                self.driver.find_element(by=By.ID, value='Contenu_Contenu_rdIndiceRefTousFonds').click()
                select = Select(self.driver.find_element(by=By.ID, value='Contenu_Contenu_cmbIndiceRef_Comp'))
                select.select_by_value(self.benchmarks[filename[:-6]])
                # time.sleep(1.5) # troppo veloce
            
            # Aggiorna indicatori / benchmark lista personalizzato
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'Contenu_Contenu_bntProPlusRafraichir'))
            )
            ActionChains(self.driver).double_click(on_element=element).perform()
            try:
                # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                # Codice vecchio
                # loading_img = self.driver.find_element(by=By.ID, value='Contenu_Contenu_loader_imgLoad')
                # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                # WebDriverWait(self.driver, 5).until(EC.visibility_of(loading_img)) 
            except TimeoutException:
                input('premi enter se ha caricato')
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))

            # Cambia date a 3 anni
            data_di_avvio_3_anni = self.driver.find_element(by=By.ID, value='Contenu_Contenu_dtDebut_txtDatePicker')
            data_di_avvio_3_anni.clear()
            data_di_avvio_3_anni.send_keys(self.t0_3Y) 
            data_di_fine = self.driver.find_element(by=By.ID, value='Contenu_Contenu_dtFin_txtDatePicker')
            data_di_fine.clear()
            data_di_fine.send_keys(self.t1)

            # Aggiorna date
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'Contenu_Contenu_lnkRefresh'))
            )
            ActionChains(self.driver).double_click(on_element=element).perform()
            try:
                # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
            except TimeoutException:
                input('premi enter se ha caricato')
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))

            # Esporta CSV
            element = WebDriverWait(self.driver, 600).until(
                EC.presence_of_element_located((By.ID, 'Contenu_Contenu_btnExportCSV'))
            )
            ActionChains(self.driver).double_click(on_element=element).perform()
            try:
                # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
            except TimeoutException:
                input('premi enter se ha caricato')
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))

            # Rinomina file a 3 anni
            """La chiavetta USB ha una velocità di scrittura inferiore a quella del computer, 
            il cambio di nome del file non può essere eseguito all'istante.
            Aspetto che l'estensione dell'ultimo file aggiunto alla cartella di download cambi
            da .crdownload a .csv; a quel punto modifico il nome del file."""
            while len(os.listdir(self.directory_output_liste)) == file_scaricati:
                time.sleep(1)
            time.sleep(1.5)
            list_of_files = glob.glob(self.directory_output_liste.__str__() + '/*')
            latest_file = max(list_of_files, key=os.path.getctime)
            while Path(latest_file).suffix != '.csv':
                list_of_files = glob.glob(self.directory_output_liste.__str__() + '/*')
                latest_file = max(list_of_files, key=os.path.getctime)
                time.sleep(1)
                # print(Path(latest_file).suffix)
            os.rename(latest_file, self.directory_output_liste.joinpath(filename[:-4]+'_3Y.csv'))

            # Cambia date a 1 anno
            data_di_avvio_1_anno = self.driver.find_element(by=By.ID, value='Contenu_Contenu_dtDebut_txtDatePicker')
            data_di_avvio_1_anno.clear()
            data_di_avvio_1_anno.send_keys(self.t0_1Y)

            # Aggiorna date
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'Contenu_Contenu_lnkRefresh'))
            )
            ActionChains(self.driver).double_click(on_element=element).perform()
            try:
                # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
            except TimeoutException:
                input('premi enter se ha caricato')
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
            
            # Esporta CSV
            element = WebDriverWait(self.driver, 600).until(
                EC.presence_of_element_located((By.ID, 'Contenu_Contenu_btnExportCSV'))
            )
            ActionChains(self.driver).double_click(on_element=element).perform()
            try:
                # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
            except TimeoutException:
                input('premi enter se ha caricato')
            else:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
            
            # Rinomina file ad 1 anno
            """La chiavetta USB ha una velocità di scrittura inferiore a quella del computer, 
            il cambio di nome del file non può essere eseguito all'istante.
            Aspetto che l'estensione dell'ultimo file aggiunto alla cartella di download cambi
            da .crdownload a .csv; a quel punto modifico il nome del file."""
            while len(os.listdir(self.directory_output_liste)) == file_scaricati:
                time.sleep(1)
            time.sleep(1.5)
            list_of_files = glob.glob(self.directory_output_liste.__str__() + '/*')
            latest_file = max(list_of_files, key=os.path.getctime)
            while Path(latest_file).suffix != '.csv':
                list_of_files = glob.glob(self.directory_output_liste.__str__() + '/*')
                latest_file = max(list_of_files, key=os.path.getctime)
                time.sleep(1)
                # print(Path(latest_file).suffix)
            os.rename(latest_file, self.directory_output_liste.joinpath(filename[:-4]+'_1Y.csv'))

            # Contatori tempo trascorso
            end = time.perf_counter()
            elapsed_time.append(end - start)
            average_elapsed_time = sum(elapsed_time)/len(elapsed_time)
            liste_completate += 1
            file_rimanenti = file_totali-liste_completate
            if file_rimanenti != 0:
                print(f"Tempo trascorso per {filename}: ", round(end - start, 2), 'secondi')
                print(f"\nTempo medio trascorso per lista: {round(average_elapsed_time, 2)} secondi")
                print(f"\nTempo previsto alla fine: {datetime.timedelta(seconds=(average_elapsed_time)*(file_rimanenti))}")
        
        self.driver.close()


if __name__ == '__main__':
    start = time.perf_counter()
    _ = Scarico(intermediario='RAI')
    _.export()
    end = time.perf_counter()
    print("Elapsed time: ", end - start, 'seconds')