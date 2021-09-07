import os
import re
import glob
import time
import datetime
import dateutil.relativedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.touch_actions import TouchActions
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException


class Scarico():
    """
    Importa le liste complete e scarica i dati da Quantalys.it
    """
    # TODO : sistema il drag and drop
    # username='Pomante', password='Pomante22'

    def __init__(self, t1, username='AVicario', password='AVicario123', directory_output_liste="C:\\Users\\Administrator\\Desktop\\Sbwkrq\\Ranking\\export_liste_from_Q", directory_input_liste='C:\\Users\\Administrator\\Desktop\\Sbwkrq\\Ranking\\import_liste_into_Q\\'):
        """
        Initialize the class.
        Default download folder : self.directory_output_liste
        Default browser : chromium

        Parameters:
        username(str) = username dell'account
        password(str) = password dell'account
        t1 = data finale
        directory_output_liste = percorso in cui scaricare i dati delle liste
        directory_input_liste = percorso in cui trovare i dati delle liste
        """
        self.username = username
        self.password = password
        self.t1 = t1
        self.t0_3Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+3)).strftime("%d/%m/%Y") # data iniziale tre anni fa
        print(f"Tre anni fa : {self.t0_3Y}.")
        self.t0_1Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+1)).strftime("%d/%m/%Y") # data iniziale un anno fa
        print(f"Un anno fa : {self.t0_1Y}.")
        self.directory_input_liste = directory_input_liste
        self.directory_output_liste = directory_output_liste
        if not os.path.exists(self.directory_output_liste):
            print('ahaha')
            os.makedirs(self.directory_output_liste)
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": self.directory_output_liste,
            "download.directory_upgrade": True}
            )
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.implicitly_wait(5)
        
    def accesso_a_quantalys(self):
        """
        Accede a quantalys.it con chromium. Imposta come cartella di download il percorso in self.directory_output_liste
        e massimizza la finestra.
        """
        print('\n...connessione a Quantalys...')
        self.driver.get("http://www.quantalys.it")
        self.driver.maximize_window()

    def login(self):
        """
        Accede all'account con username=self.username e password=self.password.
        Chiude l'alert dei cookies.
        """
        # Chiudi i cookies
        try:
            time.sleep(1)
            self.driver.find_element_by_xpath('//*[@id="tarteaucitronPersonalize2"]').click() # Cookies
        except NoSuchElementException:
            pass
        # Connessione
        try:
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnConnexion"]'))) # Connessione
        except TimeoutException:
            pass
        else:
            self.driver.find_element_by_xpath('//*[@id="btnConnexion"]').click()
        # Username e password
        try:
            WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="inputLogin"]'))) 
        except TimeoutException:
            pass
        else:
            time.sleep(0.5)
            self.driver.find_element_by_xpath('//*[@id="inputLogin"]').send_keys(self.username)
            self.driver.find_element_by_xpath('//*[@id="inputPassword"]').send_keys(self.password,Keys.ENTER)
    
    def rimuovi_indicatori(self, numero_iniziale, numero_finale=10):
        """
        Rimuovi gli indicatori presenti nel menu "indicatori calcolati" a partire dal numero iniziale fino al finale.

        Parameters:
        numero_iniziale(int) = 
        numero_finale(int) = 
        """
        for i in range(numero_iniziale, numero_finale):
            oggetto = 'imgDelete_' + str(i)
            try:
                self.driver.find_element_by_id(oggetto).click()
                time.sleep(1)
            except NoSuchElementException:
                pass
        
    def aggiungi_indicatori(self, *indicatori):
        """
        Aggiungi gli indicatori nel menu "indicatori calcolati"
        
        Parameters:
        indicatori(tuple) = tuple di indicatori da aggiungere.
        """
        # Modifica la posizione di rilascio dei due indicatori finali.
        # Se gli indicatori sono già presenti pass!
        # Prova a ridurre il tempo di esecuzione, rallenta troppo il codice
        for item in indicatori:
            if item == 'Codice ISIN':
                self.driver.find_element_by_partial_link_text('Codice ISIN').click()
            elif item == 'Nome':
                self.driver.find_element_by_partial_link_text('Nome').click()
            else:
                self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_cColonnesSelector_ctrlTreeColonnes_searchInput"]').clear()
                self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_cColonnesSelector_ctrlTreeColonnes_searchInput"]').send_keys(item)
                time.sleep(2) 
                source_element = self.driver.find_element_by_partial_link_text(item)
                time.sleep(2)
                dest_element = self.driver.find_element_by_xpath('//*[@id="SelectableItemsWrapper"]/div[3]/div')
                time.sleep(2)
                ActionChains(self.driver).drag_and_drop(source_element, dest_element).perform()
                time.sleep(2)

    def export(self, intermediario):
        """
        Carica le liste in quantalys.it, scarica gli indicatori pertinenti ed esporta un file csv.
        Rinomina il file con nomi in successione relativi alla macrocategoria.
        """
        # Il processo parte se la cartella di download è vuota
        while len(os.listdir('./export_liste_from_Q')) != 0:
            print(f"\nCi sono dei file presenti nella cartella di download: {glob.glob(self.directory_output_liste+'/*')}\n")
            _ = input('cancella i file prima di proseguire, poi premi enter\n')
        
        directory = self.directory_input_liste
        elapsed_time = []
        for filename in os.listdir(directory):
            file_totali = len(os.listdir(self.directory_output_liste))
            start = time.perf_counter()
            print(f"\nCaricamento lista {filename}...\n")
            try:
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="position-menu-quantalys"]/div/div[1]/a/img')))
            except TimeoutException:
                pass

            try:
                liste = self.driver.find_element_by_partial_link_text('Liste') # Liste
                liste.click()
            except NoSuchElementException:
                self.driver.find_element_by_partial_link_text('Tools').click() # Tools
                try:
                    WebDriverWait(self.driver, 1).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Liste'))) #Liste
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_partial_link_text('Liste').click() 

            try:
                WebDriverWait(self.driver, 7).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[2]/div/div[2]/div[1]/button'))) # Nuova lista
            except TimeoutException:
                pass
            finally:
                time.sleep(1.5) # Necessario, va troppo veloce.
                self.driver.find_element_by_name('new').click()

            try:
                WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.NAME, 'nom'))) # Nome
            except TimeoutException:
                pass
            finally:
                self.driver.find_element_by_name("nom").send_keys(filename[:-4], Keys.TAB, Keys.TAB, Keys.ENTER)

            try:
                WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]'))) # Importa dei prodotti
            except TimeoutException:
                pass
            finally:
                time.sleep(1.5) # Necessario, va troppo veloce.
                id_lista = self.driver.find_element_by_xpath('/html/body/div[1]/div[3]/input[1]').get_attribute('value') # prendi la chiave unica della lista
                self.driver.find_element_by_xpath('//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]').click()

            try:
                WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.NAME, 'file'))) # Seleziona lista da importare
            except TimeoutException:
                pass
            finally:
                self.driver.find_element_by_name("file").send_keys(self.directory_input_liste + filename)

            try:
                WebDriverWait(self.driver, 7).until(EC.presence_of_element_located((By.XPATH, '//*[@id="importForm"]/button'))) # Importa
            except TimeoutException:
                pass
            finally:
                self.driver.find_element_by_xpath('//*[@id="importForm"]/button').click()

            try:
                WebDriverWait(self.driver,120).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[3]/div[2]'), '0 elementi'))
                # WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr[1]/td[1]/label'))) # Seleziona tutti i fondi
            except TimeoutException:
                pass
            finally:
                # print(self.driver.find_element_by_xpath('//*[@id="DataTables_Table_0"]/tbody/tr/td').text)
                totale_fondi_lista = self.driver.find_element_by_xpath('//*[@id="DataTables_Table_0_info"]').text.replace(',','') # Totale fondi
                print(totale_fondi_lista)
                num_fondi_regex = re.compile(r'\d(\d)?(\d)?(\d)?')
                mo = num_fondi_regex.search(totale_fondi_lista)
                numero_fondi = mo.group()
            NUM_MAX_FONDI_CONFRONTO_DIRETTO = 1300
            if int(numero_fondi) < NUM_MAX_FONDI_CONFRONTO_DIRETTO:
                self.driver.find_element_by_xpath('//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label').click()
                time.sleep(2) # Necessario, va troppo veloce.
                self.driver.find_element_by_xpath('//*[@id="quantasearch"]/div[2]/div[3]/div/button[3]').click() # Confronta
            else:
                self.driver.find_element_by_partial_link_text('Fondi').click() # Fondi
                try:
                    WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, 'Confronto'))) # Confronto
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_partial_link_text('Confronto').click()
                
                try:
                    WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_selectFonds_searchButton"]'))) # Cerca
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_selectFonds_ctrlTreeListe_ddlDynatree"]').click()
                    # Seleziona il nome usando le regular expressions
                    # num_fondi_regex = re.compile(r'\d(\d)?(\d)?(\d)?')
                    # mo = num_fondi_regex.search(totale_fondi_lista)
                    # time.sleep(2)
                    # self.driver.find_element_by_partial_link_text(filename[:-4]+' ('+mo.group()+' fondi)').click()
                    # Seleziona il nome usando l'identificatore unico della lista
                    json_file_liste_string = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_selectFonds_ctrlTreeListe_hidJson"]').get_attribute('value')
                    json_file_liste_string = json_file_liste_string.replace('false', "'False'") # necessario per il passaggio successivo
                    json_file_liste_list = eval(json_file_liste_string) # converte la stringa in una lista di dizionari
                    _ = (__ for __ in json_file_liste_list if __['key'] == str(id_lista)) # scegli il dizionario che contiene l'id della lista appena caricata
                    nome_lista = next(_)['title'] # ricava il nome della lista
                    time.sleep(2)
                    self.driver.find_element_by_partial_link_text(nome_lista).click()
                    self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_selectFonds_ctrlTreeListe_hypValider"]').click()

                try:
                    WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_selectFonds_searchButton"]'))) # Cerca
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_selectFonds_searchButton"]').click()
                
                try:
                    WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_selectFonds_listeFonds_HeaderButton"]'))) # Seleziona tutti
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_selectFonds_listeFonds_HeaderButton"]').click()

                try:
                    WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_btnComparer1"]'))) # Confronta
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_btnComparer1"]').click()

            try:
                WebDriverWait(self.driver, 360).until(EC.presence_of_element_located((By.LINK_TEXT, 'Personalizzato'))) # Personalizzato
            except TimeoutException:
                pass
            finally:
                self.driver.find_element_by_link_text('Personalizzato').click()
            
            # Seleziona indicatori
            # time.sleep(0.5) #invisibility of element located //*[@id="Contenu_Contenu_loader_imgLoad"] or /html/body/div[2]/form/table/tbody/tr[1]/td/div/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div[2]/div/img
            try:
                WebDriverWait(self.driver, 180).until(EC.presence_of_element_located((By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]'))) # Seleziona indicatori
            except TimeoutException:
                pass
            finally:
                self.driver.find_element_by_tag_name('body').send_keys(Keys.PAGE_DOWN)
                self.rimuovi_indicatori(6)
                try:
                    ind_1 = self.driver.find_element_by_xpath('//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]').text
                    ind_2 = self.driver.find_element_by_xpath('//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[2]').text
                    ind_3 = self.driver.find_element_by_xpath('//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[3]').text
                except NoSuchElementException:
                    self.rimuovi_indicatori(1) # Rimuovi tutti gli indicatori
                    self.aggiungi_indicatori('Codice ISIN', 'Nome', 'Valuta')
                    if filename.startswith('AZ') or filename.startswith('OBB'):
                        # Aggiungi TEV e IR
                        self.aggiungi_indicatori('TEV da data a data', 'Information ratio da data a data')
                    elif filename.startswith('FLEX') or filename.startswith('BIL'):
                        # Aggiungi DSR e Sortino
                        self.aggiungi_indicatori('DSR da data a data', 'Sortino ratio da data a data')
                    elif filename.startswith('OPP'):
                        # Aggiungi Volatilità e Sharpe
                        self.aggiungi_indicatori('Volatilità da data a data', 'Sharpe ratio da data a data')
                    elif filename.startswith('LIQ'):
                        # Aggiungi volatilità e performance
                        self.aggiungi_indicatori('Volatilità da data a data', 'Perf Ann. da data a data')
                else:
                    if ind_1 != 'Codice ISIN' or ind_2 != 'Nome' or ind_3 != 'Valuta':
                        self.rimuovi_indicatori(1) # Rimuovi tutti gli indicatori
                        self.aggiungi_indicatori('Codice ISIN', 'Nome', 'Valuta')
                        if filename.startswith('AZ') or filename.startswith('OBB'):
                            # Aggiungi TEV e IR
                            self.aggiungi_indicatori('TEV da data a data', 'Information ratio da data a data')
                        elif filename.startswith('FLEX') or filename.startswith('BIL'):
                            # Aggiungi DSR e Sortino
                            self.aggiungi_indicatori('DSR da data a data', 'Sortino ratio da data a data')
                        elif filename.startswith('OPP'):
                            # Aggiungi Volatilità e Sharpe
                            self.aggiungi_indicatori('Volatilità da data a data', 'Sharpe ratio da data a data')
                        elif filename.startswith('LIQ'):
                            # Aggiungi volatilità e performance
                            self.aggiungi_indicatori('Volatilità da data a data', 'Perf Ann. da data a data')
                    else:
                        try:
                            ind_4 = self.driver.find_element_by_xpath('//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[4]').text
                            ind_5 = self.driver.find_element_by_xpath('//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[5]').text
                        except NoSuchElementException:
                            self.rimuovi_indicatori(4)
                            if filename.startswith('AZ') or filename.startswith('OBB'):
                                # Aggiungi TEV e IR
                                self.aggiungi_indicatori('TEV da data a data', 'Information ratio da data a data')
                            elif filename.startswith('FLEX') or filename.startswith('BIL'):
                                # Aggiungi DSR e Sortino
                                self.aggiungi_indicatori('DSR da data a data', 'Sortino ratio da data a data')
                            elif filename.startswith('OPP'):
                                # Aggiungi Volatilità e Sharpe
                                self.aggiungi_indicatori('Volatilità da data a data', 'Sharpe ratio da data a data')
                            elif filename.startswith('LIQ'):
                                # Aggiungi volatilità e performance
                                self.aggiungi_indicatori('Volatilità da data a data', 'Perf Ann. da data a data')
                        else:
                            if filename.startswith('AZ') or filename.startswith('OBB'):
                                if ind_4 == 'Information ratio da data a data' and ind_5 == 'TEV da data a data':
                                    pass
                                else:
                                    self.rimuovi_indicatori(4)
                                    self.aggiungi_indicatori('TEV da data a data', 'Information ratio da data a data')
                            elif filename.startswith('FLEX') or filename.startswith('BIL'):
                                if ind_4 == 'Sortino ratio da data a data' and ind_5 == 'DSR da data a data':
                                    pass
                                else:
                                    self.rimuovi_indicatori(4)
                                    self.aggiungi_indicatori('DSR da data a data', 'Sortino ratio da data a data')
                            elif filename.startswith('OPP'):
                                if ind_4 == 'Sharpe ratio da data a data' and ind_5 == 'Volatilità da data a data':
                                    pass
                                else:
                                    self.rimuovi_indicatori(4)
                                    self.aggiungi_indicatori('Volatilità da data a data', 'Sharpe ratio da data a data')
                            elif filename.startswith('LIQ'):
                                if ind_4 == 'Perf Ann. da data a data' and ind_5 == 'Volatilità da data a data':
                                    pass
                                else:
                                    self.rimuovi_indicatori(4)
                                    self.aggiungi_indicatori('Volatilità da data a data', 'Perf Ann. da data a data')

                # Prova 1 : 5 elementi, il secondo è sbagliato
                # Prova 2 : 5 elementi, il quarto è sbagliato
                # Prova 3 : 4 elementi
                # Prova 4 : 3 elementi
                # Prova 5 : 3 elementi,  il secondo è sbagliato
                # Prova 6 : 2 elementi

            # Aggiungi benchmark
            classi_a_benchmark_BPPB = {'AZ_EUR': '    MSCI Europe', 'AZ_NA': '    MSCI USA', 'AZ_PAC': '    MSCI Pacific', 'AZ_EM': '    MSCI Emerging Markets', 
                'OBB_BT': '    ICE BofA 1-3 Y Euro Broad Mkt', 'OBB_MLT': '    ICE BofA Euro Broad Market', 'OBB_CORP': '    ICE BofA Euro Corporate', 'OBB_GLOB': '    ICE BofA Global Broad Market',
                'OBB_EM': '    ICE BofA Glb Cross Corp& Gov'}
            classi_a_benchmark_BPL = {'AZ_EUR': '    MSCI Europe', 'AZ_NA': '    MSCI USA', 'AZ_PAC': '    MSCI Pacific', 'AZ_EM': '    MSCI Emerging Markets', 'AZ_GLOB': '    MSCI World',
                'OBB_BT': '    ICE BofA 1-3 Y Euro Broad Mkt', 'OBB_MLT': '    ICE BofA Euro Broad Market', 'OBB_EUR': '    ICE BofA Pan-Europe Broad Mkt', 'OBB_CORP': '    ICE BofA Euro Corporate',
                'OBB_GLOB': '    ICE BofA Global Broad Market', 'OBB_USA': '    ICE BofA US Broad Market', 'OBB_EM': '    ICE BofA Glb Cross Corp& Gov', 'OBB_GLOB_HY': '    ICE BofA Global High Yield'}
            classi_a_benchmark_CRV = {'AZ_EUR': '    MSCI Europe', 'AZ_NA': '    MSCI USA', 'AZ_PAC': '    MSCI Pacific', 'AZ_EM': '    MSCI Emerging Markets', 'AZ_GLOB': '    MSCI World',
                'OBB_BT': '    ICE BofA 1-3 Y Euro Broad Mkt', 'OBB_MLT': '    ICE BofA Euro Broad Market', 'OBB_CORP': '    ICE BofA Euro Corporate', 'OBB_GLOB': '    ICE BofA Global Broad Market', 
                'OBB_EM': '    ICE BofA Glb Cross Corp& Gov', 'OBB_GLOB_HY': '    ICE BofA Global High Yield'}
            if intermediario == 'BPPB':
                classi_a_benchmark = classi_a_benchmark_BPPB
            elif intermediario == 'BPL':
                classi_a_benchmark = classi_a_benchmark_BPL
            elif intermediario == 'CRV':
                classi_a_benchmark = classi_a_benchmark_CRV
            if filename[:-6] in classi_a_benchmark.keys():
                self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_rdIndiceRefTousFonds"]').click() # Aggiungi benchmark se classe a benchmark
                time.sleep(1)
                select = Select(self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_cmbIndiceRef_Comp"]'))
                time.sleep(2)
                select.select_by_visible_text(classi_a_benchmark[filename[:-6]])

                try:
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_bntProPlusRafraichir"]'))) # Aggiorna benchmark
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_bntProPlusRafraichir"]').click()
                    loading_img = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_loader_imgLoad"]')
                    WebDriverWait(self.driver, 10).until(EC.visibility_of(loading_img))
            else: # Aggiorna anche per le classi non a benchmark
                try:
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_bntProPlusRafraichir"]')))
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_bntProPlusRafraichir"]').click()
                    loading_img = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_loader_imgLoad"]')
                    WebDriverWait(self.driver, 10).until(EC.visibility_of(loading_img))

            # Aggiorna date a 3 anni
            try:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            except TimeoutException:
                pass
            finally:
                data_di_avvio_3_anni = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_dtDebut_txtDatePicker"]')
                data_di_avvio_3_anni.clear()
                data_di_avvio_3_anni.send_keys(self.t0_3Y) 
                data_di_fine = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_dtFin_txtDatePicker"]')
                data_di_fine.clear()
                data_di_fine.send_keys(self.t1)
                self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_lnkRefresh"]').click() # Aggiorna date
                loading_img = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_loader_imgLoad"]')
                WebDriverWait(self.driver, 10).until(EC.visibility_of(loading_img))


            # Salva il file con nome a tre anni
            try:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            except TimeoutException:
                pass
            finally:
                WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_btnExportCSV"]')))
                self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_btnExportCSV"]').click()
                try:
                    loading_img = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_loader_imgLoad"]')
                    WebDriverWait(self.driver, 600).until(EC.visibility_of(loading_img))
                except TimeoutException:
                    pass
                finally:
                    WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
                    time.sleep(1)

            # Rinomina file
            while len(os.listdir(self.directory_output_liste)) == file_totali:
                time.sleep(1)
            time.sleep(1.5)
            list_of_files = glob.glob(self.directory_output_liste + '/*')
            latest_file = max(list_of_files, key=os.path.getctime)
            os.rename(latest_file, self.directory_output_liste + '/'+filename[:-4]+'_3Y.csv')

            # Aggiorna date 1 anno
            try:
                WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_dtDebut_txtDatePicker"]')))
            except TimeoutException:
                pass
            finally:
                data_di_avvio_1_anno = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_dtDebut_txtDatePicker"]')
                data_di_avvio_1_anno.clear()
                data_di_avvio_1_anno.send_keys(self.t0_1Y)
                self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_lnkRefresh"]').click() # Aggiorna date
                loading_img = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_loader_imgLoad"]')
                WebDriverWait(self.driver, 10).until(EC.visibility_of(loading_img))


            # Salva il file con nome ad un anno
            try:
                WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
            except TimeoutException:
                pass
            finally:
                WebDriverWait(self.driver, 600).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_btnExportCSV"]')))
                self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_btnExportCSV"]').send_keys(Keys.ENTER)
                try:
                    loading_img = self.driver.find_element_by_xpath('//*[@id="Contenu_Contenu_loader_imgLoad"]')
                    WebDriverWait(self.driver, 600).until(EC.visibility_of(loading_img))
                except TimeoutException:
                    pass
                finally:
                    WebDriverWait(self.driver, 600).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_loader_imgLoad"]')))
                    time.sleep(1)
            
            # Rinomina file
            while len(os.listdir(self.directory_output_liste)) == file_totali:
                time.sleep(1)
            time.sleep(1.5)
            list_of_files = glob.glob(self.directory_output_liste + '/*')
            latest_file = max(list_of_files, key=os.path.getctime)
            os.rename(latest_file, self.directory_output_liste + '/'+filename[:-4]+'_1Y.csv')

            end = time.perf_counter()
            elapsed_time.append(end - start)
            print(f"Elapsed time for {filename}: ", end - start, 'seconds')
            print(f"\nAverage elapsed time: {sum(elapsed_time)/len(elapsed_time)}.")

        
        self.driver.close()


if __name__ == '__main__':
    start = time.perf_counter()
    _ = Scarico(t1='30/06/2021')
    _.accesso_a_quantalys()
    _.login()
    _.export(intermediario='CRV')
    end = time.perf_counter()
    print("Elapsed time: ", end - start, 'seconds')