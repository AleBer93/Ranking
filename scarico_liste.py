import datetime
import glob
import os
import re
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


class Scarico():
    # TODO : sistema il drag and drop

    def __str__(self):
        return "Importa le liste complete e scarica i dati da Quantalys.it"

    def __init__(self, intermediario, t1, username='AVicario', password='AVicario123'):
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
        self.username = username
        self.password = password
        # alt account username='Pomante', password='Pomante22'
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": self.directory_output_liste.__str__(),
            "download.directory_upgrade": True}
            )
        # API dove trovare il chromedriver aggiornato -> https://chromedriver.storage.googleapis.com/index.html
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        
    def accesso_a_quantalys(self):
        """
        Accede a quantalys.it con chromium. Imposta come cartella di download il percorso in self.directory_output_liste
        e massimizza la finestra.
        """
        print('\n...connessione a Quantalys...')
        self.driver.get("https://www.quantalys.it")
        self.driver.maximize_window()

    def login(self):
        """
        Accede all'account con username=self.username e password=self.password.
        Chiude l'alert dei cookies.
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
    
    def rimuovi_indicatori(self, numero_iniziale, numero_finale=10):
        """
        Rimuovi gli indicatori presenti nel menu "indicatori calcolati" a partire dal numero iniziale fino al finale.

        Arguments:
            numero_iniziale {int} = punto di partenza
            numero_finale {int} = punto di fine
        """
        for i in range(numero_iniziale, numero_finale):
            oggetto = 'imgDelete_' + str(i)
            try:
                WebDriverWait(self.driver, 1).until(EC.presence_of_element_located((By.ID, oggetto)))
            except:
                pass
            else:
                self.driver.find_element(by=By.ID, value=oggetto).click()
        
    def aggiungi_indicatori(self, *indicatori):
        """
        Aggiungi gli indicatori nel menu "indicatori calcolati"
        
        Arguments:
            indicatori {arg} = indicatori da aggiungere.
        """
        # Modifica la posizione di rilascio dei due indicatori finali.
        # Se gli indicatori sono già presenti pass!
        # Prova a ridurre il tempo di esecuzione, rallenta troppo il codice
        for item in indicatori:
            if item == 'Codice ISIN':
                self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Codice ISIN').click()
            elif item == 'Nome':
                self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Nome').click()
            else:
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_cColonnesSelector_ctrlTreeColonnes_searchInput"]').clear()
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_cColonnesSelector_ctrlTreeColonnes_searchInput"]').send_keys(item)
                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, item)))
                source_element = self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value=item)
                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div')))
                dest_element = self.driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div')
                ActionChains(self.driver).drag_and_drop(source_element, dest_element).perform()

    def export(self):
        """
        Carica le liste in quantalys.it, scarica gli indicatori pertinenti ed esporta un file csv.
        Rinomina il file con nomi in successione relativi alla macrocategoria.
        """
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
                time.sleep(1)
                self.driver.find_element(by=By.NAME, value="nom").send_keys(filename[:-4], Keys.TAB, Keys.TAB, Keys.ENTER) # Conferma
            # Importa prodotti
            try:
                WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]'))) # Importa dei prodotti
            except TimeoutException:
                pass
            finally:
                time.sleep(0.5) # Necessario, va troppo veloce.
                id_lista = self.driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[3]/input[1]').get_attribute('value') # prendi la chiave unica della lista
                self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]').click()
            # Scegli un file da importare
            try:
                WebDriverWait(self.driver, 3).until(EC.presence_of_element_located((By.NAME, 'file'))) # Seleziona lista da importare
            except TimeoutException:
                pass
            finally:
                self.driver.find_element(by=By.NAME, value="file").send_keys(self.directory_input_liste.joinpath(filename).__str__())
            # Importa lista
            try:
                WebDriverWait(self.driver, 7).until(EC.presence_of_element_located((By.XPATH, '//*[@id="importForm"]/button'))) # Importa
            except TimeoutException:
                pass
            finally:
                self.driver.find_element(by=By.XPATH, value='//*[@id="importForm"]/button').click()

            try:
                WebDriverWait(self.driver,120).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr/td'), 'Nessun dato disponibile'))
                # WebDriverWait(self.driver,120).until_not(EC.text_to_be_present_in_element((By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[3]/div[2]'), '0 elementi'))
            except TimeoutException: 
                pass
            finally:
                # print(self.driver.find_element_by_xpath('//*[@id="DataTables_Table_0"]/tbody/tr/td').text)
                totale_fondi_lista = self.driver.find_element(by=By.XPATH, value='//*[@id="DataTables_Table_0_info"]').text.replace(',','')
                print(f'{totale_fondi_lista}\n')
                num_fondi_regex = re.compile(r'\d(\d)?(\d)?(\d)?')
                mo = num_fondi_regex.search(totale_fondi_lista)
                numero_fondi = mo.group()
            NUM_MAX_FONDI_CONFRONTO_DIRETTO = 1300
            if int(numero_fondi) < NUM_MAX_FONDI_CONFRONTO_DIRETTO:
                self.driver.find_element(by=By.XPATH, value='//*[@id="DataTables_Table_0"]/thead/tr/th[1]/label').click()
                time.sleep(2) # Necessario, va troppo veloce.
                self.driver.find_element(by=By.XPATH, value='//*[@id="quantasearch"]/div[2]/div[3]/div/button[3]').click() # Confronta
            else:
                self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Fondi').click() # Fondi
                try:
                    WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, 'Confronto'))) # Confronto
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Confronto').click()
                
                try:
                    WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_selectFonds_searchButton"]'))) # Cerca
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_selectFonds_ctrlTreeListe_ddlDynatree"]').click()
                    # Seleziona il nome usando le regular expressions
                    # num_fondi_regex = re.compile(r'\d(\d)?(\d)?(\d)?')
                    # mo = num_fondi_regex.search(totale_fondi_lista)
                    # time.sleep(2)
                    # self.driver.find_element_by_partial_link_text(filename[:-4]+' ('+mo.group()+' fondi)').click()
                    # Seleziona il nome usando l'identificatore unico della lista
                    json_file_liste_string = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_selectFonds_ctrlTreeListe_hidJson"]').get_attribute('value')
                    json_file_liste_string = json_file_liste_string.replace('false', "'False'") # necessario per il passaggio successivo
                    json_file_liste_list = eval(json_file_liste_string) # converte la stringa in una lista di dizionari
                    _ = (__ for __ in json_file_liste_list if __['key'] == str(id_lista)) # scegli il dizionario che contiene l'id della lista appena caricata
                    nome_lista = next(_)['title'] # ricava il nome della lista
                    time.sleep(2)
                    self.driver.find_element(by=By.PARTIAL_LINK_TEXT, value=nome_lista).click()
                    self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_selectFonds_ctrlTreeListe_hypValider"]').click()

                try:
                    WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_selectFonds_searchButton"]'))) # Cerca
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_selectFonds_searchButton"]').click()
                
                try:
                    WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_selectFonds_listeFonds_HeaderButton"]'))) # Seleziona tutti
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_selectFonds_listeFonds_HeaderButton"]').click()

                try:
                    WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_btnComparer1"]'))) # Confronta
                except TimeoutException:
                    pass
                finally:
                    self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_btnComparer1"]').click()

            try:
                WebDriverWait(self.driver, 360).until(EC.presence_of_element_located((By.LINK_TEXT, 'Personalizzato'))) # Personalizzato
            except TimeoutException:
                pass
            finally:
                self.driver.find_element(by=By.LINK_TEXT, value='Personalizzato').click()
            
            # Seleziona indicatori
            # time.sleep(0.5) #invisibility of element located //*[@id="Contenu_Contenu_loader_imgLoad"] or /html/body/div[2]/form/table/tbody/tr[1]/td/div/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div[2]/div/img
            try:
                WebDriverWait(self.driver, 180).until(EC.presence_of_element_located((By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]'))) # Seleziona indicatori
            except TimeoutException:
                pass
            finally:
                self.driver.find_element(by=By.TAG_NAME, value='body').send_keys(Keys.PAGE_DOWN)
                self.rimuovi_indicatori(6)
                try:
                    ind_1 = self.driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]').text
                    ind_2 = self.driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[2]').text
                    ind_3 = self.driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[3]').text
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
                            ind_4 = self.driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[4]').text
                            ind_5 = self.driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[5]').text
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

            # Aggiungi benchmark
            classi_a_benchmark_BPPB = {'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 
                'OBB_BT': '2265', 'OBB_MLT': '2264', 'OBB_CORP': '2272', 'OBB_GLOB': '2309', 'OBB_EM': '2476'}
            classi_a_benchmark_BPL = {'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 'AZ_GLOB': '2318',
                'OBB_BT': '2265', 'OBB_MLT': '2264', 'OBB_EUR': '2255', 'OBB_CORP': '2272', 'OBB_GLOB': '2309', 'OBB_USA': '2490',
                'OBB_EM': '2476', 'OBB_GLOB_HY': '2293'}
            classi_a_benchmark_CRV = {'AZ_EUR': '2320', 'AZ_NA': '2453', 'AZ_PAC': '2325', 'AZ_EM': '2598', 'AZ_GLOB': '2318',
                'OBB_BT': '2265', 'OBB_MLT': '2264', 'OBB_CORP': '2272', 'OBB_GLOB': '2309', 'OBB_EM': '2476', 'OBB_GLOB_HY': '2293'}
            if self.intermediario == 'BPPB':
                classi_a_benchmark = classi_a_benchmark_BPPB
            elif self.intermediario == 'BPL':
                classi_a_benchmark = classi_a_benchmark_BPL
            elif self.intermediario == 'CRV':
                classi_a_benchmark = classi_a_benchmark_CRV
            if filename[:-6] in classi_a_benchmark.keys():
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_rdIndiceRefTousFonds"]').click() # Aggiungi benchmark se classe a benchmark
                # time.sleep(1) # troppo lento
                select = Select(self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_cmbIndiceRef_Comp"]'))
                # time.sleep(2) # troppo lento
                select.select_by_value(classi_a_benchmark[filename[:-6]])
            
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_bntProPlusRafraichir"]'))) # Aggiorna benchmark
                self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_bntProPlusRafraichir"]').click()
            try:
            # except TimeoutException:
            #     pass
            # finally:
                loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
                WebDriverWait(self.driver, 5).until(EC.visibility_of(loading_img)) # da 10 a 5 perché se la lista è piccola lo devo mandare avanti a mano
            except:
                input('premi enter se ha caricato')
            # else: # Aggiorna anche per le classi non a benchmark
            #     try:
            #         WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Contenu_Contenu_bntProPlusRafraichir"]')))
            #     except TimeoutException:
            #         pass
            #     finally:
            #         self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_bntProPlusRafraichir"]').click()
            #         loading_img = self.driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_loader_imgLoad"]')
            #         WebDriverWait(self.driver, 10).until(EC.visibility_of(loading_img))

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
    _ = Scarico(intermediario='BPPB', t1='31/01/2022')
    _.accesso_a_quantalys()
    _.login()
    _.export()
    end = time.perf_counter()
    print("Elapsed time: ", end - start, 'seconds')
