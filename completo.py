import datetime
import glob
import math
import os
import time
from pathlib import Path

import dateutil.relativedelta
import pandas as pd
from classes.quantalys import Quantalys
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
# with os.add_dll_directory('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\_blpapi'):
#     import blpapi
from xbbg import blp


class Completo():

    def __init__(self, intermediario):
        """
        Arguments:
            intermediario {str} - intermediario a cui è destinata l'analisi
        """
        self.intermediario = intermediario

        # Directories
        directory = Path().cwd()
        self.directory = directory
        self.directory_output_liste_complete = self.directory.joinpath('docs', 'export_liste_complete_from_Q')
        self.directory_sfdr = self.directory.joinpath('docs', 'sfdr')
        self.directory_liste_sfdr = self.directory.joinpath('docs', 'sfdr', 'liste_sfdr')
        self.directory_input_liste = self.directory.joinpath('docs', 'import_liste_into_Q')
        self.directory_alfa_nulli = self.directory.joinpath('docs', 'alfa_nulli')
        if not os.path.exists(self.directory_alfa_nulli):
            os.makedirs(self.directory_alfa_nulli)

        self.file_completo = 'completo.csv'

    def concatenazione_liste_complete(self):
        """Concatena verticalmente i file excel all'interno della cartella directory_output_liste_complete.
        Salva il risultato con il nome completo.csv
        """
        df = pd.concat((pd.read_csv(self.directory_output_liste_complete.joinpath(filename), sep = ';', decimal=',', engine='python',
            encoding = "utf_16_le", skipfooter=1) for filename in os.listdir(self.directory_output_liste_complete)), ignore_index=True)
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def individua_t1(self):
        """Individua t1, ovvero la data finale di calcolo degli indicatori, all'interno della colonna
        'Data di calcolo fine mese' presente nel file completo.csv. 
        Questo dato deve essere nullo o uguale per tutte le volte in cui non lo è.
        Salva il risultato nel file t1.txt
        """
        df = pd.read_csv(self.file_completo, sep=';', decimal=',', index_col=None)
        data_calcolo_fine_mese = df.loc[df['Data di calcolo fine mese'].notna(), 'Data di calcolo fine mese'].unique().tolist()
        try:
            assert len(data_calcolo_fine_mese) == 1, 'Ci sono più date utilizzate per il calcolo degli indicatori alla fine del mese'
        except AssertionError as e:
            print(e)
        else:
            with open('docs/t1.txt', 'w') as f:
                f.write(data_calcolo_fine_mese[0])
            t1 = datetime.datetime.strptime(data_calcolo_fine_mese[0], '%Y-%m-%d').strftime("%d/%m/%Y")
            print(f'La data di calcolo degli indicatori alla fine del mese è {t1}\n')

    def concatenazione_sfdr(self):
        """
        Concatena i file sfdr all'interno della directory_sfdr l'uno sotto l'altro.
        Salva il risultato con il nome sfdr.csv
        """
        df = pd.concat((pd.read_csv(self.directory_liste_sfdr.joinpath(filename), sep = ';', decimal=',', engine='python', 
            encoding = "unicode_escape") for filename in os.listdir(self.directory_liste_sfdr)), ignore_index=True)
        df.to_csv(self.directory_sfdr.joinpath('sfdr.csv'), sep=";", mode='w', index=False, decimal=',')
    
    def merge_completo_sfdr(self):
        """Concatena orizzontalmente i file completo.csv e sfdr.csv
        Controlla che i nomi dei due file uniti siano identici, tolti ad entrambi gli spazi bianchi.
        """
        df_completo = pd.read_csv(self.file_completo, sep=';', decimal=',', index_col=None)
        df_sfdr = pd.read_csv(self.directory_sfdr.joinpath('sfdr.csv'), sep=';', decimal=',', index_col=None)
        df_sfdr = df_sfdr[['Nome', 'SFDR']]
        assert len(df_completo) == len(df_sfdr)
        df = pd.concat([df_completo, df_sfdr], axis=1)
        df['Nome del fondo'].str.replace(" ","").equals(df['Nome'].str.replace(" ",""))
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def fondi_non_presenti(self):
        """Identifica i fondi non presenti in piattaforma e salvali nel percorso docs/prodotti_non_presenti.csv
        """
        df_1 = pd.read_csv(self.file_completo, sep=';', decimal=',', index_col=None)
        df_2 = pd.read_excel('catalogo_fondi.xlsx')
        df_3 = pd.concat([df_1['Codice ISIN'], df_2['isin']])
        df_4 = df_3.drop_duplicates(keep=False)
        prodotti_non_presenti = df_2.loc[df_2['isin'].isin(df_4), ['isin', 'valuta', 'nome']]
        print(f'Ci sono {len(prodotti_non_presenti)} prodotti non presenti nella piattaforma.\n')
        # print(f'I prodotti non presenti nella piattaforma sono i seguenti:\n{prodotti_non_presenti}')
        prodotti_non_presenti.to_csv(self.directory.joinpath('docs', 'prodotti_non_presenti.csv'), sep=';', decimal=',', index=False)

    def seleziona_colonne(self):
        """Seleziona le colonne desiderate dal file completo.csv
        """
        colonne = [
            'Codice ISIN', 'Valuta', 'Nome del fondo', 'Categoria Quantalys', 'SRI', 'Rischio 1 anno fine mese',
            'Rischio 3 anni") fine mese', 'Alpha 1 anno fine mese', 'Info 1 anno fine mese', 'Alpha 3 anni") fine mese',
            'Info 3 anni") fine mese',
        ]
        # colonne = ['Codice ISIN', 'Valuta', 'Nome del fondo', 'Categoria Quantalys', 'Rischio 1 anno fine mese',
        #     'Rischio 3 anni") fine mese', 'Info 1 anno fine mese', 'Alpha 1 anno fine mese', 'Info 3 anni") fine mese',
        #     'Alpha 3 anni") fine mese', 'SRI', 'SFDR',
        # ]
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        df = df[colonne]
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def correzione_micro_russe(self):
        """
        Corregge le righe delle microcategorie Az. Paesi Emerg. Europa e Russia & Az. Paesi Emerg. Europa ex Russia
        perchè vanno a capo dalla sesta colonna in poi.
        """
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        indexes_to_drop = []
        for row in range(len(df)):
            if df.loc[row, 'Categoria Quantalys'] == 'Az. Paesi Emerg. Europa e Russia' or df.loc[row, 'Categoria Quantalys'] == 'Az. Paesi Emerg. Europa ex Russia':
                if all(df.iloc[row, 5:len(df.columns)+1].isnull()): # solo se tutte le celle successive sono vuote
                    df.iloc[row, 5:len(df.columns)+1] = df.iloc[row+1, 1:len(df.columns)-4].values # sia values che tolist() le copia come stringhe
                    indexes_to_drop.append(row+1)
        df.drop(df.index[indexes_to_drop], inplace=True)
        if indexes_to_drop:
            print(f"\nLe righe da eliminare dopo aver copiato il loro contenuto in quella sopra sono: {indexes_to_drop}")
        df.reset_index(drop=True, inplace=True)
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def correzione_alfa_IR_nulli(self):
        """
        Quantalys calcola l'alfa fino alla quarta cifra dopo la virgola,
        se le prime quatto cifre sono 0, l'alfa sarà 0, e così anche l'IR.
        Un valore di alfa e IR pari a 0 inficia i due metodi successivi in cui viene
        calcolata la TEV e viene calcolato l'indicatore corretto.
        Sostiuisci i valori di alfa e IR 0 con i valori corretti.
        """
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        with open('docs/t1.txt') as f:
            t1 = f.read()
        t1 = datetime.datetime.strptime(t1, '%Y-%m-%d').strftime("%d/%m/%Y")
        t0_3Y = (
            datetime.datetime.strptime(t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+3)
        ).strftime("%d/%m/%Y") # data iniziale tre anni fa
        t0_1Y = (
            datetime.datetime.strptime(t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+1)
        ).strftime("%d/%m/%Y") # data iniziale un anno fa

        # Il processo parte se la cartella di download è vuota
        while len(os.listdir(self.directory_alfa_nulli)) != 0:
            print(f"\nCi sono dei file presenti nella cartella di download: {glob.glob(self.directory_alfa_nulli.__str__()+'/*')}\n")
            _ = input('cancella i file prima di proseguire, poi premi enter\n')

        # Individua i fondi che hanno un alfa a 3Y e un IR a 3Y pari a 0.
        alfa_nulli_3Y = df.loc[(df['Alpha 3 anni") fine mese']==0) | (df['Info 3 anni") fine mese']==0)]
        if len(alfa_nulli_3Y['Codice ISIN'].to_list()) != 0:
            print(alfa_nulli_3Y['Codice ISIN'].to_list())
        
        # Individua i fondi che hanno un alfa a 1Y e un IR a 1Y pari a 0.
        alfa_nulli_1Y = df.loc[(df['Alpha 1 anno fine mese']==0) & (df['Info 1 anno fine mese']==0)]
        if len(alfa_nulli_1Y['Codice ISIN'].to_list()) != 0:
            print(alfa_nulli_1Y['Codice ISIN'].to_list())
        
        if not alfa_nulli_3Y.empty or not alfa_nulli_1Y.empty:
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_experimental_option(
                "prefs", {"download.default_directory" : self.directory_alfa_nulli.__str__(),"download.directory_upgrade" : True}
            )
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            q = Quantalys()
            # Accesso a Quantalys
            q.connessione(driver)
            # Log in
            q.login(driver, 'Avicario', 'AVicario123')
            
            if not alfa_nulli_3Y.empty:
                # Export
                file_scaricati = len(os.listdir(self.directory_alfa_nulli))

                # passa da fondi confronto e incolla i pochi isin
                q.confronta_isin(driver, *[isin for isin in alfa_nulli_3Y['Codice ISIN'].to_list()])
                
                # Personalizzato
                WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.LINK_TEXT, 'Personalizzato'))).click()
                
                # Seleziona indicatori
                WebDriverWait(driver, 180).until(
                    EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]')
                    )
                )
                driver.find_element(by=By.TAG_NAME, value='body').send_keys(Keys.PAGE_DOWN)
                time.sleep(0.5) # altrimenti non centra il doppio click su 'aggiorna'
                q.aggiungi_indicatori_v2(
                    driver, 'Codice ISIN', 'Nome', 'Valuta', 'Information ratio da data a data',
                    'Alfa da data a data', 'TEV da data a data'
                )

                # Seleziona benchmark di default
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'Contenu_Contenu_rdIndiceRefFonds'))).click()
                
                # Aggiorna indicatori / benchmark lista personalizzato
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'Contenu_Contenu_bntProPlusRafraichir'))
                )
                ActionChains(driver).double_click(on_element=element).perform()
                try:
                    # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                except TimeoutException:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
            
                # Cambia date a 3 anni
                data_di_avvio_3_anni = driver.find_element(by=By.ID, value='Contenu_Contenu_dtDebut_txtDatePicker')
                data_di_avvio_3_anni.clear()
                data_di_avvio_3_anni.send_keys(t0_3Y) 
                data_di_fine = driver.find_element(by=By.ID, value='Contenu_Contenu_dtFin_txtDatePicker')
                data_di_fine.clear()
                data_di_fine.send_keys(t1)

                # Aggiorna date
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'Contenu_Contenu_lnkRefresh'))
                )
                ActionChains(driver).double_click(on_element=element).perform()
                try:
                    # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                except TimeoutException:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))

                # Esporta CSV
                element = WebDriverWait(driver, 600).until(
                    EC.presence_of_element_located((By.ID, 'Contenu_Contenu_btnExportCSV'))
                )
                ActionChains(driver).double_click(on_element=element).perform()
                try:
                    # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                except TimeoutException:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))

                # Rinomina file a 3 anni
                while len(os.listdir(self.directory_alfa_nulli)) == file_scaricati:
                    time.sleep(1)
                time.sleep(1.5)
                list_of_files = glob.glob(self.directory_alfa_nulli.__str__() + '/*')
                latest_file = max(list_of_files, key=os.path.getctime)
                os.rename(latest_file, self.directory_alfa_nulli.joinpath('alfa_nulli_3Y.csv'))

            if not alfa_nulli_1Y.empty:
                # Export
                file_scaricati = len(os.listdir(self.directory_alfa_nulli))

                # passa da fondi confronto e incolla i pochi isin
                q.confronta_isin(driver, *[isin for isin in alfa_nulli_1Y['Codice ISIN'].to_list()])

                # Personalizzato
                WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.LINK_TEXT, 'Personalizzato'))).click()
                
                # Seleziona indicatori
                WebDriverWait(driver, 180).until(
                    EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]')
                    )
                )
                driver.find_element(by=By.TAG_NAME, value='body').send_keys(Keys.PAGE_DOWN)
                time.sleep(0.5) # altrimenti non centra il doppio click su 'aggiorna'
                q.aggiungi_indicatori_v2(
                    driver, 'Codice ISIN', 'Nome', 'Valuta', 'Information ratio da data a data',
                    'Alfa da data a data', 'TEV da data a data'
                )

                # Seleziona benchmark di default
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'Contenu_Contenu_rdIndiceRefFonds'))).click()
                
                # Aggiorna indicatori / benchmark lista personalizzato
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'Contenu_Contenu_bntProPlusRafraichir'))
                )
                ActionChains(driver).double_click(on_element=element).perform()
                try:
                    # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                except TimeoutException:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))

                # Cambia date a 1 anno
                data_di_avvio_1_anno = driver.find_element(by=By.ID, value='Contenu_Contenu_dtDebut_txtDatePicker')
                data_di_avvio_1_anno.clear()
                data_di_avvio_1_anno.send_keys(t0_1Y)
                data_di_fine = driver.find_element(by=By.ID, value='Contenu_Contenu_dtFin_txtDatePicker')
                data_di_fine.clear()
                data_di_fine.send_keys(t1)

                # Aggiorna date
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'Contenu_Contenu_lnkRefresh'))
                )
                ActionChains(driver).double_click(on_element=element).perform()
                try:
                    # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                except TimeoutException:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                
                # Esporta CSV
                element = WebDriverWait(driver, 600).until(
                    EC.presence_of_element_located((By.ID, 'Contenu_Contenu_btnExportCSV'))
                )
                ActionChains(driver).double_click(on_element=element).perform()
                try:
                    # solo 5 perché se la lista è piccola devo mandare avanti il processo a mano
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                except TimeoutException:
                    input('premi enter se ha caricato')
                else:
                    WebDriverWait(driver, 600).until(EC.invisibility_of_element_located((By.ID, 'Contenu_Contenu_loader_imgLoad')))
                
                # Rinomina file ad 1 anno
                while len(os.listdir(self.directory_alfa_nulli)) == file_scaricati:
                    time.sleep(1)
                time.sleep(1.5)
                list_of_files = glob.glob(self.directory_alfa_nulli.__str__() + '/*')
                latest_file = max(list_of_files, key=os.path.getctime)
                os.rename(latest_file, self.directory_alfa_nulli.joinpath('alfa_nulli_1Y.csv'))

            driver.close()

        # Correzione indicatore corretto a 3 anni #
        while any(df['Info 3 anni") fine mese']==0) or any(df['Alpha 3 anni") fine mese']==0):
            print("Ci sono dei fondi con alpha 3 anni o information ratio 3 anni uguale a 0\n")
            print("è necessario aggiornarli per l'analisi successiva")
            _ = input(f'apri il file {self.file_completo}, aggiorna i dati, poi premi enter\n')
            df = pd.read_csv('completo.csv', sep=";", decimal=',', index_col=None)

        # Correzione indicatore corretto ad 1 anno #
        while any(df['Info 1 anno fine mese']==0) or any(df['Alpha 1 anno fine mese']==0):
            print("Ci sono dei fondi con alpha 1 anno o information ratio 1 anno uguale a 0\n")
            print("è necessario aggiornarli per l'analisi successiva")
            _ = input(f'apri il file {self.file_completo}, aggiorna i dati, poi premi enter\n')
            df = pd.read_csv('completo.csv', sep=";", decimal=',', index_col=None)

        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def merge_completo_catalogo(self):
        """Concatena orizzontalmente i file completo.csv e catalogo_fondi.xlsx
        """
        df_completo = pd.read_csv(self.file_completo, sep=";", decimal=',',index_col=None)
        df_catalogo = pd.read_excel('catalogo_fondi.xlsx')
        df_catalogo = df_catalogo[['isin', 'commissione']]
        df_merged = pd.merge(df_completo, df_catalogo, left_on='Codice ISIN', right_on='isin', how='left')
        print(f"Il primo file contiene {len(df_completo)} fondi, mentre il secondo ne contiene {len(df_catalogo)}.\
        \nL'unione dei due file contiene {len(df_merged)} fondi.\n")
        df_merged.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def assegna_macro(self):
        """Assegna una macrocategoria ad ogni microcategoria.
        """
        BPPB_dict = {
            'Monetari Euro' : 'LIQ', 'Monetari Euro dinamici' : 'LIQ', 'Monet. altre valute europee' : 'LIQ', 
            'Monetari altre valute europ' : 'LIQ', 
            'Obblig. euro gov. breve termine' : 'OBB_EUR_BT', 'Obblig. Euro breve term.' : 'OBB_EUR_BT', 
            'Obblig. Euro a scadenza' : 'OBB_EUR_BT', 
            'Obblig. Euro gov. medio termine' : 'OBB_EUR_MLT', 'Obblig. Euro gov. lungo termine' : 'OBB_EUR_MLT', 
            'Obblig. Euro lungo termine' : 'OBB_EUR_MLT', 'Obblig. Euro medio term.' : 'OBB_EUR_MLT', 'Obblig. Euro gov.' : 'OBB_EUR_MLT', 
            'Obblig. Euro all maturities' : 'OBB_EUR_MLT', 'Obblig. Europa' : 'OBB_EUR_MLT', 'Obblig. Sterlina inglese' : 'OBB_EUR_MLT', 
            'Obblig. Franco svizzero' : 'OBB_EUR_MLT', 'Obblig. Indicizz. Inflation Linked' : 'OBB_EUR_MLT', 
            'Obblig. Euro corporate' : 'OBB_EUR_CORP', 
            'Obblig. paesi emerg. Asia' : 'OBB_EM', 'Obblig. paesi emerg. Europa' : 'OBB_EM', 'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 
            'Obblig. Paesi Emerg.' : 'OBB_EM', 'Obblig. paesi emerg. a scadenza' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Hard Currency' : 'OBB_EM', 'Obblig. Paesi Emerg. Local Currency' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Asia Local Ccy' : 'OBB_EM',
            'Obblig. Dollaro US breve term.' : 'OBB_GLOB', 'Obblig. USD medio-lungo term.' : 'OBB_GLOB', 
            'Obblig. Dollaro US medio-lungo term.' : 'OBB_GLOB', 'Obblig. USD corporate' : 'OBB_GLOB', 
            'Obblig. Dollaro US corporate' : 'OBB_GLOB', 'Obblig. Doll. US all maturities' : 'OBB_GLOB', 
            'Obblig. Dollaro US all mat' : 'OBB_GLOB', 'Obblig. Asia' : 'OBB_GLOB', 'Obblig. globale' : 'OBB_GLOB',
            'Obblig. globale corporate' : 'OBB_GLOB', 'Obblig. Yen' : 'OBB_GLOB', 'Obblig. altre valute' : 'OBB_GLOB',
            "Obblig. Indicizz. all'inflaz. USD" : 'OBB_GLOB', 'Obblig. Global Inflation Linked' : 'OBB_GLOB', 
            'Monetari Dollaro USA' : 'OBB_GLOB', 'Monet. ex Europa altre valute' : 'OBB_GLOB', 
            'Monetari ex Europa altre valute' : 'OBB_GLOB', 
            'Obblig. Euro high yield' : 'OBB_HY', 'Obblig. Europa High Yield' : 'OBB_HY', 
            'Obblig. Dollaro US high yield' : 'OBB_HY', 'Obblig. globale high yield' : 'OBB_HY', 
            'Az. Europa' : 'AZ_EUR', 'Az. Area Euro' : 'AZ_EUR', 'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Area Euro Growth' : 'AZ_EUR', 
            'Az. Area Euro Value' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 
            'Az. Europa Value' : 'AZ_EUR', 'Az. Belgio' : 'AZ_EUR', 'Az. Francia' : 'AZ_EUR', 'Az. Francia small cap' : 'AZ_EUR', 
            'Az. Germania' : 'AZ_EUR', 'Az. Germania small cap' : 'AZ_EUR', 'Az. Spagna' : 'AZ_EUR', 'Az. Paesi Bassi' : 'AZ_EUR', 
            'Az. Italia' : 'AZ_EUR', 'Az. UK' : 'AZ_EUR', 'Az. UK small cap' : 'AZ_EUR', 'Az. Svizzera' : 'AZ_EUR', 
            'Az.Svizzera small cap' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Europa altri paesi' : 'AZ_EUR', 
            'Azionario USA' : 'AZ_NA', 'Az. USA' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 'Az. USA Growth' : 'AZ_NA', 
            'Az. USA Value' : 'AZ_NA', 'Az. Canada' : 'AZ_NA', 
            'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 'Az. Giappone small cap' : 'AZ_PAC', 
            'Az. Pacifico' : 'AZ_PAC', 
            'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 'Az. Paesi Emerg. Europa ex Russia' : 'AZ_EM', 'Az. Russia' : 'AZ_EM',
            'Az. paesi emerg. Asia' : 'AZ_EM', 'Az. BRIC' : 'AZ_EM', 'Az. Grande Cina' : 'AZ_EM', 'Az. Cina' : 'AZ_EM',
            'Az. paesi emerg. America Latina' : 'AZ_EM', 'Az. paesi emerg. altre zone' : 'AZ_EM', 'Az. paesi emerg. Mondo' : 'AZ_EM', 
            'Az. India' : 'AZ_EM', 'Az. Brasile' : 'AZ_EM', 'Az. Altri paesi emerg.' : 'AZ_EM', 
            'Commodities a leva' : 'OPP', 'Commodities Bear' : 'OPP', 'Commodities' : 'OPP', 'Obblig. Convertib. Euro' : 'OPP', 
            'Obblig. Convertib. Europa' : 'OPP', 'Obblig. Convertib. Dollaro US' : 'OPP', 'Obblig. Convertib. Glob.' : 'OPP', 
            'Az. real estate Europa' : 'OPP', 'Az. Biotech' : 'OPP', 'Az. beni di consumo' : 'OPP', 'Az. ambiente' : 'OPP', 
            'Az. energia, materie prime, oro' : 'OPP', 'Az. energia. materie prime. oro' : 'OPP', 'Az. energia materie prime oro' : 'OPP', 
            'Az. real estate Mondo' : 'OPP', 'Az. industria' : 'OPP', 'Az. salute   farmaceutico' : 'OPP',
            'Az. salute – farmaceutico' : 'OPP', 'Az. salute - farmaceutico' : 'OPP', 'Az. Servizi di pubblica utilita' : 'OPP', 
            'Az. servizi finanziari' : 'OPP', 'Az. tecnologia' : 'OPP', 'Az. telecomunicazioni' : 'OPP', 'Az. Oro' : 'OPP', 
            'Az. Bear' : 'OPP', 'Obblig. Bear' : 'OPP', 'Valuta Long/Short' : 'OPP', 'Altri' : 'OPP',
            'Bilanc. Prud. Europa' : 'FLEX', 'Bilanc. Prud. Global Euro' : 'FLEX', 'Bilanc. Prud. Dollaro US' : 'FLEX', 
            'Bilanc. Prud. Global' : 'FLEX', 'Bilanc. Prud. altre valute' : 'FLEX', 'Bilanc. Equilib. Europa' : 'FLEX', 
            'Bilanc. Equil. Global Euro' : 'FLEX', 'Bilanc. Equil. Dollaro US' : 'FLEX', 'Bilanc. Equil. Global' : 'FLEX', 
            'Bilanc. Equil. altre valute' : 'FLEX', 'Bilanc. Aggress. Europa' : 'FLEX', 'Bilanc. Aggress. Global Euro' : 'FLEX', 
            'Bilanc. aggress. Dollaro US' : 'FLEX', 'Bilanc. Aggress. Global' : 'FLEX', 'Bilanc. Aggress. altre valute' : 'FLEX', 
            'Flessibili Europa' : 'FLEX', 'Fless. Global Euro' : 'FLEX', 'Flessibili prudenti Europa' : 'FLEX', 
            'Flessibili Dollaro US' : 'FLEX', 'Flessibili prudenti globale' : 'FLEX', 'Fless. Global' : 'FLEX', 
            'Fondi a scadenza pred. Euro' : 'FLEX', 'Fondi a scadenza pred. altre valute' : 'FLEX', 'Perf. ass. Dividendi' : 'FLEX', 
            'Perf. Ass. Arbitr.Fus.-acquis. Euro' : 'FLEX', 'Perf. assoluta strategia valute' : 'FLEX', 
            'Perf. assoluta Market Neutral Euro' : 'FLEX', 'Perf. ass. Long/Short eq.' : 'FLEX', 'Perf. assoluta tassi' : 'FLEX', 
            'Perf. assoluta volatilita' : 'FLEX', 'Perf. assoluta multi-strategia' : 'FLEX', 'Perf. assoluta (GBP)' : 'FLEX', 
            'Perf. ass. USD' : 'FLEX', 'Fondi  a garanzia o a formula Euro' : 'FLEX', 'Az. globale' : 'FLEX', 
            'Az. globale small cap' : 'FLEX', 'Az. globale Growth' : 'FLEX', 'Az. globale Value' : 'FLEX', 
            'F.a garanz. o a formul. altr valu.' : 'FLEX', 
        }
        BPL_dict = {
            'Monetari Euro' : 'LIQ', 'Monetari Euro dinamici' : 'LIQ', 
            'Monet. ex Europa altre valute' : 'LIQ_FOR', 'Monetari ex Europa altre valute' : 'LIQ_FOR', 
            'Monet. altre valute europee' : 'LIQ_FOR', 'Monetari altre valute europ' : 'LIQ_FOR', 'Monetari Dollaro USA' : 'LIQ_FOR', 
            'Obblig. euro gov. breve termine' : 'OBB_EUR_BT', 'Obblig. Euro breve term.' : 'OBB_EUR_BT', 
            'Obblig. Euro gov. medio termine' : 'OBB_EUR_MLT', 'Obblig. Euro gov. lungo termine' : 'OBB_EUR_MLT', 
            'Obblig. Euro lungo termine' : 'OBB_EUR_MLT', 'Obblig. Euro medio term.' : 'OBB_EUR_MLT', 'Obblig. Euro gov.' : 'OBB_EUR_MLT', 
            'Obblig. Euro all maturities' : 'OBB_EUR_MLT',  'Obblig. Euro a scadenza' : 'OBB_EUR_MLT', 
            'Obblig. Indicizz. Inflation Linked' : 'OBB_EUR_MLT', 'Obblig. Convertib. Euro' : 'OBB_EUR_MLT', 
            'Fondi a scadenza pred. Euro' : 'OBB_EUR_MLT', 
            'Obblig. Europa' : 'OBB_EUR', 'Obblig. Sterlina inglese' : 'OBB_EUR', 'Obblig. Franco svizzero' : 'OBB_EUR', 
            'Obblig. Convertib. Europa' : 'OBB_EUR', 
            'Obblig. Euro corporate' : 'OBB_EUR_CORP', 
            'Obblig. paesi emerg. Asia' : 'OBB_EM', 'Obblig. paesi emerg. Europa' : 'OBB_EM',  'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 
            'Obblig. Paesi Emerg.' : 'OBB_EM', 'Obblig. paesi emerg. a scadenza' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Hard Currency' : 'OBB_EM', 'Obblig. Paesi Emerg. Local Currency' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Asia Local Ccy' : 'OBB_EM', 
            'Obblig. Dollaro US breve term.' : 'OBB_USA', 'Obblig. USD medio-lungo term.' : 'OBB_USA', 
            'Obblig. Dollaro US medio-lungo term.' : 'OBB_USA', 'Obblig. USD corporate' : 'OBB_USA', 
            'Obblig. Dollaro US corporate' : 'OBB_USA', 'Obblig. Doll. US all maturities' : 'OBB_USA', 
            'Obblig. Dollaro US all mat' : 'OBB_USA', 'Obblig. Convertib. Dollaro US' : 'OBB_USA', 
            "Obblig. Indicizz. all'inflaz. USD" : 'OBB_USA', 
            'Obblig. Asia' : 'OBB_GLOB', 'Obblig. globale' : 'OBB_GLOB', 'Obblig. globale corporate' : 'OBB_GLOB', 
            'Obblig. Yen' : 'OBB_GLOB', 'Obblig. altre valute' : 'OBB_GLOB', 'Obblig. Global Inflation Linked' : 'OBB_GLOB', 
            'Obblig. Convertib. Glob.' : 'OBB_GLOB', 'Fondi a scadenza pred. altre valute' : 'OBB_GLOB', 
            'Obblig. Euro high yield' : 'OBB_HY', 'Obblig. Europa High Yield' : 'OBB_HY', 
            'Obblig. Dollaro US high yield' : 'OBB_HY', 'Obblig. globale high yield' : 'OBB_HY', 
            'Az. Europa' : 'AZ_EUR', 'Az. Area Euro' : 'AZ_EUR', 'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Area Euro Growth' : 'AZ_EUR', 
            'Az. Area Euro Value' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 
            'Az. Europa Value' : 'AZ_EUR', 'Az. Belgio' : 'AZ_EUR', 'Az. Francia' : 'AZ_EUR', 'Az. Francia small cap' : 'AZ_EUR', 
            'Az. Germania' : 'AZ_EUR', 'Az. Germania small cap' : 'AZ_EUR', 'Az. Spagna' : 'AZ_EUR', 'Az. Paesi Bassi' : 'AZ_EUR', 
            'Az. Italia' : 'AZ_EUR', 'Az. UK' : 'AZ_EUR', 'Az. UK small cap' : 'AZ_EUR', 'Az. Svizzera' : 'AZ_EUR', 
            'Az.Svizzera small cap' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Europa altri paesi' : 'AZ_EUR', 
            'Azionario USA' : 'AZ_NA', 'Az. USA' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 'Az. USA Growth' : 'AZ_NA', 
            'Az. USA Value' : 'AZ_NA', 'Az. Canada' : 'AZ_NA',
            'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 'Az. Giappone small cap' : 'AZ_PAC', 
            'Az. Pacifico' : 'AZ_PAC',
            'Az. Brasile' : 'AZ_EM', 'Az. Cina' : 'AZ_EM', 'Az. India' : 'AZ_EM', 'Az. Russia' : 'AZ_EM', 
            'Az. Altri paesi emerg.' : 'AZ_EM', 'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 
            'Az. Paesi Emerg. Europa ex Russia' : 'AZ_EM', 'Az. paesi emerg. Asia' : 'AZ_EM', 'Az. BRIC' : 'AZ_EM', 
            'Az. Grande Cina' : 'AZ_EM', 'Az. paesi emerg. America Latina' : 'AZ_EM', 'Az. paesi emerg. altre zone' : 'AZ_EM', 
            'Az. paesi emerg. Mondo' : 'AZ_EM', 
            'Az. globale' : 'AZ_GLOB', 'Az. globale small cap' : 'AZ_GLOB', 'Az. globale Growth' : 'AZ_GLOB', 
            'Az. globale Value' : 'AZ_GLOB', 
            'Commodities a leva' : 'OPP', 'Commodities Bear' : 'OPP', 'Commodities' : 'OPP', 'Az. real estate Europa' : 'OPP', 
            'Az. Biotech' : 'OPP', 'Az. beni di consumo' : 'OPP', 'Az. ambiente' : 'OPP', 'Az. energia, materie prime, oro' : 'OPP', 
            'Az. energia. materie prime. oro' : 'OPP', 'Az. energia materie prime oro' : 'OPP', 'Az. real estate Mondo' : 'OPP', 
            'Az. industria' : 'OPP', 'Az. salute   farmaceutico' : 'OPP', 'Az. salute – farmaceutico' : 'OPP', 
            'Az. salute - farmaceutico' : 'OPP', 'Az. Servizi di pubblica utilita' : 'OPP', 'Az. servizi finanziari' : 'OPP', 
            'Az. tecnologia' : 'OPP', 'Az. telecomunicazioni' : 'OPP', 'Az. Oro' : 'OPP', 'Az. Bear' : 'OPP', 'Obblig. Bear' : 'OPP', 
            'Altri' : 'OPP', 'Perf. ass. Dividendi' : 'OPP', 'Perf. Ass. Arbitr.Fus.-acquis. Euro' : 'OPP', 
            'Perf. assoluta strategia valute' : 'OPP', 'Perf. assoluta Market Neutral Euro' : 'OPP', 'Perf. ass. Long/Short eq.' : 'OPP', 
            'Perf. assoluta tassi' : 'OPP', 'Perf. assoluta volatilita' : 'OPP', 'Perf. assoluta multi-strategia' : 'OPP', 
            'Perf. assoluta (GBP)' : 'OPP', 'Perf. ass. USD' : 'OPP', 'Fondi  a garanzia o a formula Euro' : 'OPP', 
            'Valuta Long/Short' : 'OPP', 'F.a garanz. o a formul. altr valu.' : 'OPP', 
            'Bilanc. Prud. Europa' : 'BIL', 'Bilanc. Prud. Global Euro' : 'BIL', 'Bilanc. Prud. Dollaro US' : 'BIL', 
            'Bilanc. Prud. Global' : 'BIL', 'Bilanc. Prud. altre valute' : 'BIL', 'Bilanc. Equilib. Europa' : 'BIL', 
            'Bilanc. Equil. Global Euro' : 'BIL', 'Bilanc. Equil. Dollaro US' : 'BIL', 'Bilanc. Equil. Global' : 'BIL', 
            'Bilanc. Equil. altre valute' : 'BIL', 'Bilanc. Aggress. Europa' : 'BIL', 'Bilanc. Aggress. Global Euro' : 'BIL', 
            'Bilanc. aggress. Dollaro US' : 'BIL', 'Bilanc. Aggress. Global' : 'BIL', 'Bilanc. Aggress. altre valute' : 'BIL', 
            'Flessibili Europa' : 'FLEX', 'Fless. Global Euro' : 'FLEX', 'Flessibili prudenti Europa' : 'FLEX', 
            'Flessibili Dollaro US' : 'FLEX', 'Flessibili prudenti globale' : 'FLEX', 'Fless. Global' : 'FLEX',
        }      
        CRV_dict = {
            'Monetari Euro' : 'LIQ', 'Monetari Euro dinamici' : 'LIQ', 'Monet. altre valute europee' : 'LIQ', 
            'Monetari altre valute    europ' : 'LIQ', 
            'Obblig. euro gov. breve termine' : 'OBB_EUR_BT', 'Obblig. Euro breve term.' : 'OBB_EUR_BT', 
            'Obblig. Euro gov. medio termine' : 'OBB_EUR_MLT', 'Obblig. Euro gov. lungo termine' : 'OBB_EUR_MLT', 
            'Obblig. Euro lungo termine' : 'OBB_EUR_MLT', 'Obblig. Euro medio term.' : 'OBB_EUR_MLT', 'Obblig. Euro gov.' : 'OBB_EUR_MLT', 
            'Obblig. Euro all maturities' : 'OBB_EUR_MLT', 'Obblig. Europa' : 'OBB_EUR_MLT', 'Obblig. Sterlina inglese' : 'OBB_EUR_MLT', 
            'Obblig. Franco svizzero' : 'OBB_EUR_MLT', 'Obblig. Indicizz. Inflation Linked' : 'OBB_EUR_MLT', 
            'Obblig. Euro corporate' : 'OBB_EUR_CORP',
            'Obblig. paesi emerg. Asia' : 'OBB_EM', 'Obblig. paesi emerg. Europa' : 'OBB_EM', 'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 
            'Obblig. Paesi Emerg.' : 'OBB_EM', 'Obblig. paesi emerg. a scadenza' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Hard Currency' : 'OBB_EM', 'Obblig. Paesi Emerg. Local Currency' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Asia Local Ccy' : 'OBB_EM', 
            'Obblig. Dollaro US breve term.' : 'OBB_GLOB', 'Obblig. USD medio-lungo term.' : 'OBB_GLOB', 
            'Obblig. Dollaro US medio-lungo term.' : 'OBB_GLOB', 'Obblig. USD corporate' : 'OBB_GLOB', 
            'Obblig. Dollaro US corporate' : 'OBB_GLOB', 'Obblig. Doll. US all maturities' : 'OBB_GLOB', 
            'Obblig. Dollaro US all mat' : 'OBB_GLOB', 'Obblig. Asia' : 'OBB_GLOB', 'Obblig. globale' : 'OBB_GLOB', 
            'Obblig. globale corporate' : 'OBB_GLOB', 'Obblig. Yen' : 'OBB_GLOB', 'Obblig. altre valute' : 'OBB_GLOB', 
            "Obblig. Indicizz. all'inflaz. USD" : 'OBB_GLOB', 'Obblig. Global Inflation Linked' : 'OBB_GLOB', 
            'Monetari Dollaro USA' : 'OBB_GLOB', 'Monet. ex Europa altre valute' : 'OBB_GLOB', 
            'Monetari ex Europa altre valute' : 'OBB_GLOB', 
            'Obblig. Euro high yield' : 'OBB_HY', 'Obblig. Europa High Yield' : 'OBB_HY', 
            'Obblig. Dollaro US high yield' : 'OBB_HY', 'Obblig. globale high yield' : 'OBB_HY',
            'Az. Europa' : 'AZ_EUR', 'Az. Area Euro' : 'AZ_EUR', 'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Area Euro Growth' : 'AZ_EUR', 
            'Az. Area Euro Value' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 
            'Az. Europa Value' : 'AZ_EUR', 'Az. Belgio' : 'AZ_EUR', 'Az. Francia' : 'AZ_EUR', 'Az. Francia small cap' : 'AZ_EUR', 
            'Az. Germania' : 'AZ_EUR', 'Az. Germania small cap' : 'AZ_EUR', 'Az. Spagna' : 'AZ_EUR', 'Az. Paesi Bassi' : 'AZ_EUR', 
            'Az. Italia' : 'AZ_EUR', 'Az. UK' : 'AZ_EUR', 'Az. UK small cap' : 'AZ_EUR', 'Az. Svizzera' : 'AZ_EUR', 
            'Az.Svizzera small cap' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Europa altri paesi' : 'AZ_EUR',
            'Azionario USA' : 'AZ_NA', 'Az. USA' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 'Az. USA Growth' : 'AZ_NA', 
            'Az. USA Value' : 'AZ_NA', 'Az. Canada' : 'AZ_NA', 
            'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 'Az. Giappone small cap' : 'AZ_PAC', 
            'Az. Pacifico' : 'AZ_PAC',
            'Az. Brasile' : 'AZ_EM', 'Az. Cina' : 'AZ_EM', 'Az. India' : 'AZ_EM', 'Az. Russia' : 'AZ_EM', 
            'Az. Altri paesi emerg.' : 'AZ_EM', 'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 
            'Az. Paesi Emerg. Europa ex Russia' : 'AZ_EM', 'Az. paesi emerg. Asia' : 'AZ_EM', 'Az. BRIC' : 'AZ_EM', 
            'Az. Grande Cina' : 'AZ_EM', 'Az. paesi emerg. America Latina' : 'AZ_EM', 'Az. paesi emerg. altre zone' : 'AZ_EM', 
            'Az. paesi emerg. Mondo' : 'AZ_EM', 
            'Az. globale' : 'AZ_GLOB', 'Az. globale small cap' : 'AZ_GLOB', 'Az. globale Growth' : 'AZ_GLOB', 
            'Az. globale Value' : 'AZ_GLOB', 
            'Bilanc. Prud. Europa' : 'FLEX', 'Bilanc. Prud. Global Euro' : 'FLEX', 'Bilanc. Prud. Dollaro US' : 'FLEX', 
            'Bilanc. Prud. Global' : 'FLEX', 'Bilanc. Prud. altre valute' : 'FLEX', 'Bilanc. Equilib. Europa' : 'FLEX', 
            'Bilanc. Equil. Global Euro' : 'FLEX', 'Bilanc. Equil. Dollaro US' : 'FLEX', 'Bilanc. Equil. Global' : 'FLEX', 
            'Bilanc. Equil. altre valute' : 'FLEX', 'Bilanc. Aggress. Europa' : 'FLEX', 'Bilanc. Aggress. Global Euro' : 'FLEX', 
            'Bilanc. aggress. Dollaro US' : 'FLEX', 'Bilanc. Aggress. Global' : 'FLEX', 'Bilanc. Aggress. altre valute' : 'FLEX', 
            'Flessibili Europa' : 'FLEX', 'Fless. Global Euro' : 'FLEX', 'Flessibili prudenti Europa' : 'FLEX', 
            'Flessibili Dollaro US' : 'FLEX', 'Flessibili prudenti globale' : 'FLEX', 'Fless. Global' : 'FLEX', 
            'Commodities a leva' : 'OPP', 'Commodities Bear' : 'OPP', 'Commodities' : 'OPP', 'Obblig. Convertib. Euro' : 'OPP', 
            'Obblig. Convertib. Europa' : 'OPP', 'Obblig. Convertib. Dollaro US' : 'OPP', 'Obblig. Convertib. Glob.' : 'OPP', 
            'Az. real estate Europa' : 'OPP', 'Az. Biotech' : 'OPP', 'Az. beni di consumo' : 'OPP', 'Az. ambiente' : 'OPP', 
            'Az. energia, materie prime, oro' : 'OPP', 'Az. energia. materie prime. oro' : 'OPP', 'Az. energia materie prime oro' : 'OPP', 
            'Az. real estate Mondo' : 'OPP', 'Az. industria' : 'OPP', 'Az. salute   farmaceutico' : 'OPP', 
            'Az. salute – farmaceutico' : 'OPP', 'Az. salute - farmaceutico' : 'OPP', 'Az. Servizi di pubblica utilita' : 'OPP', 
            'Az. servizi finanziari' : 'OPP', 'Az. tecnologia' : 'OPP', 'Az. telecomunicazioni' : 'OPP', 'Az. Oro' : 'OPP', 
            'Az. Bear' : 'OPP', 'Obblig. Bear' : 'OPP', 'Valuta Long/Short' : 'OPP', 'Altri' : 'OPP', 'Perf. ass. Dividendi' : 'OPP', 
            'Perf. Ass. Arbitr.Fus.-acquis. Euro' : 'OPP', 'Perf. assoluta strategia valute' : 'OPP', 
            'Perf. assoluta Market Neutral Euro' : 'OPP', 'Perf. ass. Long/Short eq.' : 'OPP', 'Perf. assoluta tassi' : 'OPP', 
            'Perf. assoluta volatilita' : 'OPP', 'Perf. assoluta multi-strategia' : 'OPP', 'Perf. assoluta (GBP)' : 'OPP', 
            'Perf. ass. USD' : 'OPP', 'Fondi  a garanzia o a formula Euro' : 'OPP', 'Fondi a scadenza pred. Euro' : 'OPP', 
            'Fondi a scadenza pred. altre valute' : 'OPP', 'Obblig. Euro a scadenza' : 'OPP', 'F.a garanz. o a formul. altr valu.' : 'OPP', 
        }
        RIPA_dict = {
            'Monetari Euro' : 'LIQ', 
            'Obblig. euro gov. breve termine' : 'OBB_EUR_BT', 'Obblig. Euro breve term.' : 'OBB_EUR_BT', 
            'Obblig. Euro gov. medio termine' : 'OBB_EUR_MLT', 'Obblig. Euro lungo termine' : 'OBB_EUR_MLT', 
            'Obblig. Euro medio term.' : 'OBB_EUR_MLT', 'Obblig. Euro gov. lungo termine' : 'OBB_EUR_MLT', 
            'Obblig. Euro gov.' : 'OBB_EUR_MLT', 'Obblig. Euro all maturities' : 'OBB_EUR_MLT', 'Obblig. Euro a scadenza' : 'OBB_EUR_MLT', 
            'Obblig. Indicizz. Inflation Linked' : 'OBB_EUR_MLT', 'Obblig. Convertib. Euro' : 'OBB_EUR_MLT', 
            'Obblig. Euro corporate' : 'OBB_EUR_CORP', 
            'Obblig. Europa' : 'OBB_EUR', 'Obblig. Sterlina inglese' : 'OBB_EUR', 'Obblig. Franco svizzero' : 'OBB_EUR', 
            'Obblig. Convertib. Europa' : 'OBB_EUR', 
            'Obblig. Asia' : 'OBB_GLOB', 'Obblig. globale' : 'OBB_GLOB', 'Obblig. globale corporate' : 'OBB_GLOB', 
            'Obblig. altre valute' : 'OBB_GLOB', 'Obblig. Global Inflation Linked' : 'OBB_GLOB', 'Obblig. Convertib. Glob.' : 'OBB_GLOB', 
            'Obblig. Paesi Emerg.' : 'OBB_EM', 'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 'Obblig. paesi emerg. a scadenza' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Local Currency' : 'OBB_EM', 'Obblig. Paesi Emerg. Asia Local Ccy' : 'OBB_EM', 
            'Obblig. Dollaro US breve term.' : 'OBB_USA', 'Obblig. USD medio-lungo term.' : 'OBB_USA', 
            'Obblig. Dollaro US corporate' : 'OBB_USA', 'Obblig. Dollaro US all mat' : 'OBB_USA', 
            "Obblig. Indicizz. all'inflaz. USD" : 'OBB_USA', 
            'Obblig. Yen' : 'OBB_JAP', 
            'Obblig. Euro high yield' : 'OBB_HY', 'Obblig. Europa High Yield' : 'OBB_HY', 
            'Obblig. Dollaro US high yield' : 'OBB_HY', 'Obblig. globale high yield' : 'OBB_HY', 
            'Az. Area Euro' : 'AZ_EUR', 'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Area Euro Growth' : 'AZ_EUR', 
            'Az. Area Euro Value' : 'AZ_EUR', 'Az. Europa' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 
            'Az. Europa Value' : 'AZ_EUR', 'Az. Belgio' : 'AZ_EUR', 'Az. Francia' : 'AZ_EUR', 'Az. Francia small cap' : 'AZ_EUR', 
            'Az. Germania' : 'AZ_EUR', 'Az. Germania small cap' : 'AZ_EUR', 'Az. Spagna' : 'AZ_EUR', 'Az. Paesi Bassi' : 'AZ_EUR', 
            'Az. Italia' : 'AZ_EUR', 'Az. UK' : 'AZ_EUR', 'Az. UK small cap' : 'AZ_EUR', 'Az. Svizzera' : 'AZ_EUR', 
            'Az.Svizzera small cap' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Europa altri paesi' : 'AZ_EUR', 
            'Az. USA' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 'Az. USA Growth' : 'AZ_NA', 'Az. USA Value' : 'AZ_NA', 
            'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 'Az. Giappone small cap' : 'AZ_PAC', 
            'Az. Pacifico' : 'AZ_PAC', 
            'Az. Brasile' : 'AZ_EM', 'Az. Cina' : 'AZ_EM', 'Az. India' : 'AZ_EM', 'Az. Altri paesi emerg.' : 'AZ_EM', 
            'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 'Az. Paesi Emerg. Europa ex Russia' : 'AZ_EM', 'Az. paesi emerg. Asia' :'AZ_EM', 
            'Az. BRIC' : 'AZ_EM', 'Az. Grande Cina' : 'AZ_EM', 'Az. paesi emerg. America Latina' : 'AZ_EM', 
            'Az. paesi emerg. altre zone' : 'AZ_EM', 'Az. paesi emerg. Mondo' : 'AZ_EM',
            'Az. globale' : 'AZ_GLOB', 'Az. globale small cap' : 'AZ_GLOB', 'Az. globale Growth' : 'AZ_GLOB', 
            'Az. globale Value' : 'AZ_GLOB', 
            'Az. Biotech' : 'AZ_BIO', 'Az. beni di consumo' : 'AZ_BDC', 'Az. servizi finanziari' : 'AZ_FIN', 'Az. ambiente' : 'AZ_AMB', 
            'Az. real estate Europa' : 'AZ_IMM', 'Az. real estate Mondo' : 'AZ_IMM', 'Az. industria' : 'AZ_IND', 
            'Az. energia materie prime oro' : 'AZ_ECO', 'Az. salute - farmaceutico' : 'AZ_SAL', 
            'Az. Servizi di pubblica utilita' : 'AZ_SPU', 'Az. tecnologia' : 'AZ_TEC', 'Az. telecomunicazioni' : 'AZ_TEL', 
            'Az. Oro' : 'AZ_ORO', 'Az. Bear' : 'AZ_BEAR', 
            'Perf. ass. Dividendi' : 'PERF_ASS', 'Perf. Ass. Arbitr.Fus.-acquis. Euro' : 'PERF_ASS', 
            'Perf. assoluta strategia valute' : 'PERF_ASS', 'Perf. assoluta Market Neutral Euro' : 'PERF_ASS', 
            'Perf. ass. Long/Short eq.' : 'PERF_ASS', 'Perf. assoluta tassi' : 'PERF_ASS', 'Perf. assoluta volatilita' : 'PERF_ASS', 
            'Perf. assoluta multi-strategia' : 'PERF_ASS', 'Perf. assoluta (GBP)' : 'PERF_ASS', 'Perf. ass. USD' : 'PERF_ASS', 
            'Flessibili prudenti Europa' : 'FLEX', 'Flessibili prudenti globale' : 'FLEX', 'Flessibili Europa' : 'FLEX', 
            'Fless. Global' : 'FLEX', 'Flessibili Dollaro US' : 'FLEX', 
            'Commodities' : 'COMM', 'Commodities a leva' : 'COMM', 'Commodities Bear' : 'COMM', 
            'Monetari altre valute europ' : 'ND', 'Monetari Dollaro USA' : 'ND', 'Monetari ex Europa altre valute' : 'ND', 
            'Bilanc. Prud. Europa' : 'ND', 'Bilanc. Prud. Dollaro US' : 'ND', 'Bilanc. Prud. altre valute' : 'ND', 
            'Bilanc. Prud. Global' : 'ND', 'Bilanc. Equilib. Europa' : 'ND', 'Bilanc. Equil. Dollaro US' : 'ND', 
            'Bilanc. Equil. Global' : 'ND', 'Bilanc. Equil. altre valute' : 'ND', 'Bilanc. aggress. Dollaro US' : 'ND', 
            'Bilanc. Aggress. altre valute' : 'ND', 'Bilanc. Aggress. Global' : 'ND', 'Fondi a scadenza pred. Euro' : 'ND', 
            'Obblig. paesi emerg. Asia' : 'ND', 'Altri' : 'ND', 
        }
        RAI_dict = {
            'Monetari Euro' : 'LIQ',
            'Monetari Dollaro USA' : 'LIQ_FOR', 'Monetari altre valute europ' : 'LIQ_FOR', 
            'Obblig. Euro breve term.' : 'OBB_EUR_BT', 'Obblig. euro gov. breve termine' : 'OBB_EUR_BT', 
            'Obblig. Euro all maturities' : 'OBB_EUR_MLT', 
            'Obblig. Euro gov.' : 'OBB_EUR_MLT', 'Obblig. Euro a scadenza' : 'OBB_EUR_MLT', 'Obblig. Euro medio term.' : 'OBB_EUR_MLT', 
            'Obblig. Euro lungo termine' : 'OBB_EUR_MLT', 'Obblig. Euro gov. medio termine' : 'OBB_EUR_MLT', 
            'Obblig. Euro gov. lungo termine' : 'OBB_EUR_MLT', 'Obblig. Indicizz. Inflation Linked' : 'OBB_EUR_MLT', 
            'Obblig. Convertib. Europa' : 'OBB_EUR_MLT', 
            'Obblig. Euro corporate' : 'OBB_EUR_CORP', 
            'Obblig. Europa' : 'OBB_EUR', 'Obblig. Franco svizzero' : 'OBB_EUR', 'Obblig. Sterlina inglese' : 'OBB_EUR',
            'Obblig. Dollaro US all mat' : 'OBB_USA', 'Obblig. Dollaro US breve term.' : 'OBB_USA', 
            'Obblig. Dollaro US corporate' : 'OBB_USA', 
            'Obblig. globale' : 'OBB_GLOB', 'Obblig. globale corporate' : 'OBB_GLOB', 'Obblig. Global Inflation Linked' : 'OBB_GLOB', 
            'Obblig. Convertib. Glob.' : 'OBB_GLOB', 'Obblig. Asia' : 'OBB_GLOB', 'Obblig. altre valute' : 'OBB_GLOB', 
            'Obblig. Euro high yield' : 'OBB_HY', 'Obblig. Europa High Yield' : 'OBB_HY', 'Obblig. Dollaro US high yield' : 'OBB_HY', 
            'Obblig. globale high yield' : 'OBB_HY', 
            'Obblig. Paesi Emerg.' : 'OBB_EM', 'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 'Obblig. paesi emerg. Asia' : 'OBB_EM', 
            'Az. Europa' : 'AZ_EUR', 'Az. Area Euro' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 
            'Az. Europa Value' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Svizzera' : 'AZ_EUR', 'Az.Svizzera small cap' : 'AZ_EUR', 
            'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Germania' : 'AZ_EUR', 'Az. Italia' : 'AZ_EUR', 'Az. Spagna' : 'AZ_EUR', 
            'Az. UK' : 'AZ_EUR', 
            'Az. USA' : 'AZ_NA',  'Az. USA Value' : 'AZ_NA',  'Az. USA Growth' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 
            'Az. Pacifico' : 'AZ_PAC', 'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 
            'Az. Giappone small cap' : 'AZ_PAC', 
            'Az. paesi emerg. Mondo' : 'AZ_EM', 'Az. paesi emerg. America Latina' : 'AZ_EM', 'Az. paesi emerg. Asia' : 'AZ_EM', 
            'Az. Altri paesi emerg.' : 'AZ_EM', 'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 'Az. paesi emerg. altre zone' : 'AZ_EM', 
            'Az. Brasile' : 'AZ_EM', 'Az. Grande Cina' : 'AZ_EM', 'Az. India' : 'AZ_EM', 'Az. Russia' : 'AZ_EM', 'Az. Cina' : 'AZ_EM', 
            'Az. globale' : 'AZ_GLOB', 'Az. globale Value' : 'AZ_GLOB', 'Az. globale Growth' : 'AZ_GLOB', 
            'Az. globale small cap' : 'AZ_GLOB', 
            'Bilanc. Prud. Europa' : 'BIL', 'Bilanc. Prud. Dollaro US' : 'BIL', 'Bilanc. Prud. Global' : 'BIL', 
            'Bilanc. Prud. altre valute' : 'BIL', 
            'Bilanc. Equilib. Europa' : 'BIL', 'Bilanc. Equil. Dollaro US' : 'BIL', 'Bilanc. Equil. Global' : 'BIL', 
            'Bilanc. Equil. altre valute' : 'BIL', 'Bilanc. Aggress. Global' : 'BIL', 'Bilanc. Aggress. altre valute' : 'BIL', 
            'Flessibili prudenti Europa' : 'FLEX', 'Flessibili prudenti globale' : 'FLEX', 'Flessibili Europa' : 'FLEX', 'Flessibili Dollaro US' : 'FLEX', 'Fless. Global' : 'FLEX', 
            'Az. Servizi di pubblica utilita' : 'OPP', 'Az. ambiente' : 'OPP', 'Az. beni di consumo' : 'OPP', 'Az. tecnologia' : 'OPP',
            'Az. real estate Europa' : 'OPP', 'Az. salute - farmaceutico' : 'OPP', 'Az. energia materie prime oro' : 'OPP', 
            'Az. industria' : 'OPP', 'Az. Oro' : 'OPP', 'Az. servizi finanziari' : 'OPP', 'Az. telecomunicazioni' : 'OPP', 
            'Az. real estate Mondo' : 'OPP', 'Commodities' : 'OPP', 'Fondi  a garanzia o a formula Euro' : 'OPP', 
            'Perf. assoluta multi-strategia' : 'OPP', 'Perf. assoluta tassi' : 'OPP', 'Perf. assoluta Market Neutral Euro' : 'OPP', 
            'Perf. assoluta strategia valute' : 'OPP', 'Perf. assoluta volatilita' : 'OPP', 'Perf. ass. USD' : 'OPP', 
            'Perf. ass. Long/Short eq.' : 'OPP', 'Altri' : 'OPP', 'Fondi a scadenza pred. Euro' : 'OPP', 
        }
        
        def cancella_macro(dataframe, delete_keyword, path):
            """
            Cancella le micro categorie appartenenti alle macrocategorie mappate come delete_keyword.
            Metodo usato solo da RIPA.
            """
            DELETE_KEYWORD = delete_keyword
            # df = pd.read_csv(file, sep=";", decimal=',', index_col=None)
            df = dataframe
            df_to_delete = df[df['macro_categoria'] == DELETE_KEYWORD]
            if df_to_delete.empty:
                pass
            else:
                print(f'{df_to_delete.shape[0]} fondi sono stati eliminati dal catalogo')
                df_to_delete.to_csv(path.joinpath('docs', 'prodotti_cancellati.csv'), sep=";", decimal=',', index=False)
                df = df.loc[df['macro_categoria'] != DELETE_KEYWORD, :]
            return df

        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        if self.intermediario == 'BPPB':
            df['macro_categoria'] = df['Categoria Quantalys'].map(BPPB_dict)
        elif self.intermediario == 'BPL':
            df['macro_categoria'] = df['Categoria Quantalys'].map(BPL_dict)
        elif self.intermediario == 'CRV':
            df['macro_categoria'] = df['Categoria Quantalys'].map(CRV_dict)
        elif self.intermediario == 'RIPA':
            """
            Nel caso di Ripa, solo le micro categorie appartenenti ad una macro nella sezione asset allocationdi Quantalys
            vengono analizzate, mentre le altre vengono scartate. Assegno dunque tutte le micro categorie senza macro 
            ad una macro fittizzia 'ND', che mi tornerà utile per scartare interamente questa macro dall'analisi.
            """
            df['macro_categoria'] = df['Categoria Quantalys'].map(RIPA_dict)
            df = cancella_macro(dataframe=df, delete_keyword='ND', path=self.directory)
        elif self.intermediario == 'RAI':
            df['macro_categoria'] = df['Categoria Quantalys'].map(RAI_dict)
        print(f"Ci sono {df['macro_categoria'].isnull().sum()} fondi a cui non è stata assegnata una macro categoria.\n")
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def discriminazione_flessibili_e_bilanciati(self):
        """
        Discrimina flessibili e bilanciati secondo la loro volatilità oppure le loro micro categorie
        BPPB: metodo volatilità
        BPL: metodo etichette
        CRV: metodo volatilità
        RIPA: metodo etichette
        RAI: metodo etichette
        """
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)

        def discrimina_per_rischio(row, etichette):
            """Discrimina i fondi in base alla loro indicatore di volatilità
            1. Se il fondo possiede il dato del rischio a 3 anni e questo è superiore al 5%,
            restituisci etichette[0] altrimenti etichette[1].
            2. Se il fondo non possiede il dato del rischio a 3 anni,
            ma possiede il dato del rischio a 1 anno e questo è superiore al 5%,
            restituisci etichette[0] altrimenti etichette[1].
            3. Se il fondo non possiede il dato del rischio a 3 anni nè ad 1 anno,
            ma possiede il dato SRI e questo è superiore a 3,
            restituisci etichette[0] altrimenti etichette[1].
            4. Se il fondo non possiede il dato del rischio a 3 anni nè ad 1 anno nè l'SRI,
            restituisci etichette[0].

            Arguments:
                row {pandas.core.series.Series} -- riga in cui sono presenti i dati relativi
                    al rischio a 3 anni in prima posizione, al rischio ad 1 anno in seconda,
                    e al SRI in terza.
                etichette {list} -- lista contenente un elemento in prima posizione
                    che identifica i fondi flessibili a bassa volatilità,
                    ed uno in seconda posizione che identifica quelli a media ed
                    alta volatilità.

            Raises:
                TypeError: il primo argomento non è una Series di Pandas
                TypeError: il secondo argomento non è una lista di due argomenti

            Returns:
                string -- valore da applicare alla colonna 'macro_categoria'
                    al posto di FLEX
            """
            if type(row) != pd.Series:
                raise TypeError('row deve essere una Series di Pandas')
            if type(etichette) != list or len(etichette) != 2:
                raise TypeError('etichette deve essere una lista di due elementi')
            if not math.isnan(row['Rischio 3 anni") fine mese']):
                if row['Rischio 3 anni") fine mese'] < 0.05:
                    return etichette[0]
                else:
                    return etichette[1]
            else:
                if not math.isnan(row['Rischio 1 anno fine mese']):
                    if row['Rischio 1 anno fine mese'] < 0.05:
                        return etichette[0]
                    else:
                        return etichette[1]
                else:
                    if not math.isnan(row['SRI']):
                        if row['SRI'] < 3:
                            return etichette[0]
                        else:
                            return etichette[1]
                    else:
                        return etichette[0]
        
        if self.intermediario == 'BPPB':
            print("sto discriminando i flessibili in base alla loro volatilità...\n")
            df_fondi_senza_rischio = df.loc[
                (df['macro_categoria'] == 'FLEX') & (df['Rischio 3 anni") fine mese'].isnull()) &
                (df['Rischio 1 anno fine mese'].isnull()) & (df['SRI'].isnull()),
                ['Codice ISIN', 'Nome del fondo', 'Categoria Quantalys']
            ]
            if not df_fondi_senza_rischio.empty:
                print("I seguenti fondi flessibili non sono stati classificati:\n")
                print(df_fondi_senza_rischio)
                print('\nGli verrà assegnata la categoria bassa_volatilità.\n')
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df.loc[
                df['macro_categoria'] == 'FLEX', ['Rischio 3 anni") fine mese', 'Rischio 1 anno fine mese', 'SRI']
            ].apply(lambda x: discrimina_per_rischio(x, ['FLEX_BVOL', 'FLEX_MAVOL']), axis=1)
        elif self.intermediario == 'BPL':
            print("sto discriminando i flessibili e i bilanciati in base alla loro classe di appartenenza...\n")
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df['Categoria Quantalys'].map({
                'Flessibili prudenti globale' : 'FLEX_PR', 'Flessibili prudenti Europa' : 'FLEX_PR', 'Flessibili Europa' : 'FLEX_DIN',
                'Flessibili Dollaro US' : 'FLEX_DIN', 'Fless. Global Euro' : 'FLEX_DIN', 'Fless. Global' : 'FLEX_DIN'}, na_action='ignore')
            df.loc[df['macro_categoria'] == 'BIL', 'macro_categoria'] = df['Categoria Quantalys'].map({
                'Bilanc. Prud. Europa' : 'BIL_MBVOL', 'Bilanc. Prud. Dollaro US' : 'BIL_MBVOL', 'Bilanc. Prud. Global Euro' : 'BIL_MBVOL',
                'Bilanc. Prud. Global' : 'BIL_MBVOL', 'Bilanc. Prud. altre valute' : 'BIL_MBVOL',  'Bilanc. Equilib. Europa' : 'BIL_MBVOL',
                'Bilanc. Equil. Dollaro US' : 'BIL_MBVOL', 'Bilanc. Equil. Global Euro' : 'BIL_MBVOL',
                'Bilanc. Equil. Global' : 'BIL_MBVOL', 'Bilanc. Equil. altre valute' : 'BIL_MBVOL', 'Bilanc. Aggress. Europa' : 'BIL_AVOL',
                'Bilanc. aggress. Dollaro US' : 'BIL_AVOL', 'Bilanc. Aggress. Global Euro' : 'BIL_AVOL',
                'Bilanc. Aggress. Global' : 'BIL_AVOL', 'Bilanc. Aggress. altre valute' : 'BIL_AVOL'}, na_action='ignore')
        elif self.intermediario == 'CRV':
            print("sto discriminando i flessibili in base alla loro volatilità...\n")
            df_fondi_senza_rischio = df.loc[
                (df['macro_categoria'] == 'FLEX') & (df['Rischio 3 anni") fine mese'].isnull()) &
                (df['Rischio 1 anno fine mese'].isnull()) & (df['SRI'].isnull()),
                ['Codice ISIN', 'Nome del fondo', 'Categoria Quantalys']
            ]
            if not df_fondi_senza_rischio.empty:
                print("I seguenti fondi flessibili non sono stati classificati:\n")
                print(df_fondi_senza_rischio)
                print('\nGli verrà assegnata la categoria bassa_volatilità.\n')
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df.loc[
                df['macro_categoria'] == 'FLEX', ['Rischio 3 anni") fine mese', 'Rischio 1 anno fine mese', 'SRI']
            ].apply(lambda x: discrimina_per_rischio(x, ['FLEX_PR', 'FLEX_DIN']), axis=1)
        elif self.intermediario == 'RIPA':
            print("sto discriminando i flessibili e i bilanciati in base alla loro classe di appartenenza...\n")
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df['Categoria Quantalys'].map({
                'Flessibili prudenti globale' : 'FLEX_PR', 'Flessibili prudenti Europa' : 'FLEX_PR', 'Flessibili Europa' : 'FLEX_DIN', 
                'Flessibili Dollaro US' : 'FLEX_DIN', 'Fless. Global Euro' : 'FLEX_DIN', 'Fless. Global' : 'FLEX_DIN',
                }, na_action='ignore')
        elif self.intermediario == 'RAI':
            print("sto discriminando i flessibili e i bilanciati in base alla loro classe di appartenenza...\n")
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df['Categoria Quantalys'].map({
                'Flessibili prudenti globale' : 'FLEX_PR', 'Flessibili prudenti Europa' : 'FLEX_PR', 'Flessibili Europa' : 'FLEX_DIN',
                'Flessibili Dollaro US' : 'FLEX_DIN', 'Fless. Global Euro' : 'FLEX_DIN', 'Fless. Global' : 'FLEX_DIN',}, na_action='ignore')
            df.loc[df['macro_categoria'] == 'BIL', 'macro_categoria'] = df['Categoria Quantalys'].map({
                'Bilanc. Prud. Europa' : 'BIL_PR', 'Bilanc. Prud. Dollaro US' : 'BIL_PR', 'Bilanc. Prud. Global Euro' : 'BIL_PR',
                'Bilanc. Prud. Global' : 'BIL_PR', 'Bilanc. Prud. altre valute' : 'BIL_PR',  'Bilanc. Equilib. Europa' : 'BIL_EQ',
                'Bilanc. Equil. Dollaro US' : 'BIL_EQ', 'Bilanc. Equil. Global Euro' : 'BIL_EQ', 'Bilanc. Equil. Global' : 'BIL_EQ',
                'Bilanc. Equil. altre valute' : 'BIL_EQ', 'Bilanc. Aggress. Europa' : 'BIL_AGG', 'Bilanc. aggress. Dollaro US' : 'BIL_AGG',
                'Bilanc. Aggress. Global Euro' : 'BIL_AGG', 'Bilanc. Aggress. Global' : 'BIL_AGG',
                'Bilanc. Aggress. altre valute' : 'BIL_AGG'}, na_action='ignore')
            # Vecchio metodo
            # df['categoria_flessibili'] = df.loc[
            #     (df['macro_categoria'] == 'FLEX') & (df['Rischio 3 anni") fine mese'].notnull()), 'Rischio 3 anni") fine mese'
            # ].apply(lambda x: 'bassa_vola' if x < 0.05 else 'media_alta_vola')
            # df.loc[df['categoria_flessibili'].isnull(), 'categoria_flessibili'] = df.loc[
            #     (df['macro_categoria'] == 'FLEX') & (df['Rischio 1 anno fine mese'].notnull()), 'Rischio 1 anno fine mese'
            # ].apply(lambda x: 'bassa_vola' if x < 0.05 else 'media_alta_vola')
            # df.loc[df['categoria_flessibili'].isnull(), 'categoria_flessibili'] = df.loc[
            #     (df['macro_categoria'] == 'FLEX') & (df['SRI'].notnull()), 'SRI'
            # ].apply(lambda x: 'bassa_vola' if x < 3 else 'media_alta_vola')
            # df.loc[(df['macro_categoria'] == 'FLEX') & (df['categoria_flessibili'].isnull()), 'categoria_flessibili'] = 'bassa_vola'
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def sconta_commissioni(self):
        """Sconta le commissioni dei fondi in base alla loro macro categoria
        Metodo usato solo da CRV"""
        if self.intermediario == 'CRV':
            df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
            sconti = {'LIQ' : 0.85, 'OBB_EUR_BT' : 0.35, 'OBB_EUR_MLT' : 0.35, 'OBB_EUR_CORP' : 0.35, 'OBB_EM' : 0.35,
                'OBB_GLOB' : 0.35, 'OBB_HY' : 0.35, 'AZ_EUR' : 0.30, 'AZ_NA' : 0.30, 'AZ_PAC' : 0.30, 'AZ_EM' : 0.30,
                'AZ_GLOB' : 0.30, 'FLEX_PR' : 0.60, 'FLEX_DIN' : 0.60, 'OPP' : 0.50
            }
            df['commissione'] = df['commissione']*df['macro_categoria'].apply(lambda x : sconti[x])
            df.to_csv(self.file_completo, sep=";", decimal=',', index=False)
        else:
            return None

    def scarico_datadiavvio(self):
        """Scarica la data di avvio dei fondi nel file_bloomberg utilizzando la libreria di Bloomberg.
        Aggiungi la data di avvio al file completo.
        """
        print("\nSto scaricando le dati di avvio dei fondi da Bloomberg...")
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        df_bl = blp.bdp('/isin/' + df['Codice ISIN'], flds="fund_incept_dt") #/isin/IT0001029823
        df_bl.reset_index(inplace=True)
        df_bl['isin_code'] = df_bl['index'].str[6:]
        df_bl.reset_index(drop=True, inplace=True)
        # df_bl.to_csv(self.directory.joinpath('docs', 'data_di_avvio.csv'), sep=";")
        df_merged = pd.merge(df, df_bl, left_on='Codice ISIN', right_on='isin_code', how='left')
        print('scaricate!')
        fondi_senza_data_di_avvio = df_merged.loc[df_merged['fund_incept_dt'].isna(), ['Codice ISIN', 'Valuta', 'Nome del fondo']]
        print(f"\nCi sono {df_merged['fund_incept_dt'].isnull().sum()} fondi senza data di avvio:\n{fondi_senza_data_di_avvio}\n")
        df_merged.to_csv(self.file_completo, sep=";", decimal=',', index=False)
        df = pd.read_csv('completo.csv', sep=";", decimal=',', index_col=None)
        while any(df_merged['fund_incept_dt'].isna()):
            print("Ci sono delle date mancanti, è necessario aggiornarle per l'analisi successiva,")
            _ = input(f'apri il file {self.file_completo}, aggiungi le date, poi premi enter\n')
            df_merged = pd.read_csv('completo.csv', sep=";", decimal=',', index_col=None)

    def seleziona_e_rinomina_colonne(self):
        """Seleziona e rinomina solo le colonne utili del file completo.
        """
        colonne = [
            'Codice ISIN', 'Valuta', 'Nome del fondo', 'Categoria Quantalys', 'macro_categoria', 'fund_incept_dt',
            'commissione', 'SFDR', 'Alpha 1 anno fine mese', 'Info 1 anno fine mese', 'Alpha 3 anni") fine mese',
            'Info 3 anni") fine mese',
        ]
        colonne_rinominate = [
            'ISIN', 'valuta', 'nome', 'micro_categoria', 'macro_categoria', 'data_di_avvio', 'commissione',
            'SFDR', 'Alpha 1 anno fine mese', 'Info 1 anno fine mese', 'Alpha 3 anni") fine mese',
            'Info 3 anni") fine mese',
        ]
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        df = df[colonne]
        df.columns = colonne_rinominate
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def creazione_liste_input(self):
        """
        Crea file csv, con dimensioni massime pari a 1999 elementi, da importare in Quantalys.it.
        Directory in cui vengono salvati i file : '.docs/import_liste_into_Q/'
        """
        df_com = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        if not os.path.exists(self.directory_input_liste):
            os.makedirs(self.directory_input_liste)
        for categoria in df_com['macro_categoria'].unique():
            chunks = len(df_com.loc[df_com['macro_categoria'] == categoria])//2000 +1 # blocchi da 2000 elementi
            for chunk in range(chunks):
                df = df_com.loc[df_com['macro_categoria'] == categoria, ['ISIN', 'valuta']]
                df = df.iloc[0 + 1999 * chunk : 1999 + 1999 * chunk]
                df.columns = ['codice isin', 'divisa']
                df.to_csv(self.directory_input_liste.joinpath(categoria + '_' + str(chunk) + '.csv'), sep=";", index=False)


if __name__ == '__main__':
    start = time.perf_counter()
    _ = Completo(intermediario='CRV')
    # _.concatenazione_liste_complete()
    # _.individua_t1()
    # _.seleziona_colonne()
    # _.concatenazione_sfdr()
    # _.merge_completo_sfdr()
    # _.fondi_non_presenti()
    # _.correzione_micro_russe()
    # _.correzione_alfa_IR_nulli()
    # _.merge_completo_catalogo()
    _.assegna_macro()
    # _.discriminazione_flessibili_e_bilanciati()
    # _.sconta_commissioni()
    # _.scarico_datadiavvio()
    # _.seleziona_e_rinomina_colonne()
    # _.creazione_liste_input()
    end = time.perf_counter()
    print("Elapsed time: ", round(end - start, 2), 'seconds')