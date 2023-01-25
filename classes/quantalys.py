import re
import time

import pandas as pd
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class Quantalys():

    def __repr__(self):
        return "Implementa della funzionalità di Quantalys come il raggiungimento di percorsi o lo scarico dati"

    def __init__(self):
        pass
    
    def connessione(self, driver):
        """
        Accede a quantalys.it con chromium. Imposta come cartella di download il percorso in self.directory_output_liste
        e massimizza la finestra.
        """
        print('\n...connessione a Quantalys...')
        driver.get("https://www.quantalys.it")
        driver.maximize_window()

    def login(self, driver, username='AVicario', password='AVicario123'):
        """
        Chiude l'alert dei cookies.
        Accede all'account con username=self.username e password=self.password.
        Main account username='AVicario', password='AVicario123'
        Alt account username='Pomante', password='Pomante22'
        """
        # Chiudi i cookies
        time.sleep(1)
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'tarteaucitronPersonalize2')))
            driver.find_element(by=By.ID, value='tarteaucitronPersonalize2').click() # Cookies
        except NoSuchElementException:
            pass
        # Login form
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'btnConnexion')))
        driver.find_element(by=By.ID, value='btnConnexion').click()
        # Username e password
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'inputLogin'))) 
        driver.find_element(by=By.ID, value='inputLogin').send_keys(username)
        driver.find_element(by=By.ID, value='inputPassword').send_keys(password)
        driver.find_element(by=By.ID, value='btnConnecter').click()

    def rimuovi_indicatori(self, driver, numero_iniziale, numero_finale=10):
        """
        Rimuovi gli indicatori presenti nel menu "indicatori calcolati" a partire dal numero iniziale fino al finale.

        Arguments:
            numero_iniziale {int} = punto di partenza
            numero_finale {int} = punto di fine
        """
        for i in range(numero_iniziale, numero_finale):
            oggetto = 'imgDelete_' + str(i)
            try:
                WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.ID, oggetto)))
            except:
                pass
            else:
                driver.find_element(by=By.ID, value=oggetto).click()

    def aggiungi_indicatori_v1(self, driver, *indicatori):
        """
        Aggiungi gli indicatori nel menu "indicatori calcolati"
        
        Arguments:
            indicatori {arg} = indicatori da aggiungere.
        
        Codice da aggiungere al metodo export
        # self.rimuovi_indicatori(6)
        # try:
        #     ind_1 = driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[1]').text
        #     ind_2 = driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[2]').text
        #     ind_3 = driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[3]').text
        # except NoSuchElementException:
        #     self.rimuovi_indicatori(1) # Rimuovi tutti gli indicatori
        #     self.aggiungi_indicatori_v1('Codice ISIN', 'Nome', 'Valuta')
        #     if filename.startswith('AZ') or filename.startswith('OBB'):
        #         # Aggiungi TEV e IR
        #         self.aggiungi_indicatori_v1('TEV da data a data', 'Information ratio da data a data')
        #     elif filename.startswith('FLEX') or filename.startswith('BIL') or filename.startswith('COMM') or filename.startswith('PERF'):
        #         # Aggiungi DSR e Sortino
        #         self.aggiungi_indicatori_v1('DSR da data a data', 'Sortino ratio da data a data')
        #     elif filename.startswith('OPP'):
        #         # Aggiungi Volatilità e Sharpe
        #         self.aggiungi_indicatori_v1('Volatilità da data a data', 'Sharpe ratio da data a data')
        #     elif filename.startswith('LIQ'):
        #         # Aggiungi volatilità e performance
        #         self.aggiungi_indicatori_v1('Volatilità da data a data', 'Perf Ann. da data a data')
        # else:
        #     if ind_1 != 'Codice ISIN' or ind_2 != 'Nome' or ind_3 != 'Valuta':
        #         self.rimuovi_indicatori(1) # Rimuovi tutti gli indicatori
        #         self.aggiungi_indicatori_v1('Codice ISIN', 'Nome', 'Valuta')
        #         if filename.startswith('AZ') or filename.startswith('OBB'):
        #             # Aggiungi TEV e IR
        #             self.aggiungi_indicatori_v1('TEV da data a data', 'Information ratio da data a data')
        #         elif filename.startswith('FLEX') or filename.startswith('BIL') or filename.startswith('COMM') or filename.startswith('PERF'):
        #             # Aggiungi DSR e Sortino
        #             self.aggiungi_indicatori_v1('DSR da data a data', 'Sortino ratio da data a data')
        #         elif filename.startswith('OPP'):
        #             # Aggiungi Volatilità e Sharpe
        #             self.aggiungi_indicatori_v1('Volatilità da data a data', 'Sharpe ratio da data a data')
        #         elif filename.startswith('LIQ'):
        #             # Aggiungi volatilità e performance
        #             self.aggiungi_indicatori_v1('Volatilità da data a data', 'Perf Ann. da data a data')
        #     else:
        #         try:
        #             ind_4 = driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[4]').text
        #             ind_5 = driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div/div/ol/li[5]').text
        #         except NoSuchElementException:
        #             self.rimuovi_indicatori(4)
        #             if filename.startswith('AZ') or filename.startswith('OBB'):
        #                 # Aggiungi TEV e IR
        #                 self.aggiungi_indicatori_v1('TEV da data a data', 'Information ratio da data a data')
        #             elif filename.startswith('FLEX') or filename.startswith('BIL') or filename.startswith('COMM') or filename.startswith('PERF'):
        #                 # Aggiungi DSR e Sortino
        #                 self.aggiungi_indicatori_v1('DSR da data a data', 'Sortino ratio da data a data')
        #             elif filename.startswith('OPP'):
        #                 # Aggiungi Volatilità e Sharpe
        #                 self.aggiungi_indicatori_v1('Volatilità da data a data', 'Sharpe ratio da data a data')
        #             elif filename.startswith('LIQ'):
        #                 # Aggiungi volatilità e performance
        #                 self.aggiungi_indicatori_v1('Volatilità da data a data', 'Perf Ann. da data a data')
        #         else:
        #             if filename.startswith('AZ') or filename.startswith('OBB'):
        #                 if ind_4 == 'Information ratio da data a data' and ind_5 == 'TEV da data a data':
        #                     pass
        #                 else:
        #                     self.rimuovi_indicatori(4)
        #                     self.aggiungi_indicatori_v1('TEV da data a data', 'Information ratio da data a data')
        #             elif filename.startswith('FLEX') or filename.startswith('BIL') or filename.startswith('COMM') or filename.startswith('PERF'):
        #                 if ind_4 == 'Sortino ratio da data a data' and ind_5 == 'DSR da data a data':
        #                     pass
        #                 else:
        #                     self.rimuovi_indicatori(4)
        #                     self.aggiungi_indicatori_v1('DSR da data a data', 'Sortino ratio da data a data')
        #             elif filename.startswith('OPP'):
        #                 if ind_4 == 'Sharpe ratio da data a data' and ind_5 == 'Volatilità da data a data':
        #                     pass
        #                 else:
        #                     self.rimuovi_indicatori(4)
        #                     self.aggiungi_indicatori_v1('Volatilità da data a data', 'Sharpe ratio da data a data')
        #             elif filename.startswith('LIQ'):
        #                 if ind_4 == 'Perf Ann. da data a data' and ind_5 == 'Volatilità da data a data':
        #                     pass
        #                 else:
        #                     self.rimuovi_indicatori(4)
        #                     self.aggiungi_indicatori_v1('Volatilità da data a data', 'Perf Ann. da data a data')
        """
        # Modifica la posizione di rilascio dei due indicatori finali.
        # Se gli indicatori sono già presenti pass!
        # Prova a ridurre il tempo di esecuzione, rallenta troppo il codice
        for item in indicatori:
            if item == 'Codice ISIN':
                driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Codice ISIN').click()
            elif item == 'Nome':
                driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Nome').click()
            else:
                driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_cColonnesSelector_ctrlTreeColonnes_searchInput"]').clear()
                driver.find_element(by=By.XPATH, value='//*[@id="Contenu_Contenu_cColonnesSelector_ctrlTreeColonnes_searchInput"]').send_keys(item)
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, item)))
                source_element = driver.find_element(by=By.PARTIAL_LINK_TEXT, value=item)
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="SelectableItemsWrapper"]/div[3]/div')))
                dest_element = driver.find_element(by=By.XPATH, value='//*[@id="SelectableItemsWrapper"]/div[3]/div')
                ActionChains(driver).drag_and_drop(source_element, dest_element).perform()

    def aggiungi_indicatori_v2(self, driver, *indicatori):
        """
        Aggiungi gli indicatori nello spazio "indicatori calcolati" modificando l'attributo value
        dell'input nascosto con ID 'Contenu_Contenu_cColonnesSelector_fieldJson'. Gli indicatori
        modificati saranno visibili solo dopo aver premuto il tasto aggiorna con ID 
        'Contenu_Contenu_bntProPlusRafraichir'.
        
        Arguments:
            indicatori {arg} = indicatori da aggiungere.
        """
        key_val_items = {
            'Codice ISIN' : 'sCodeISIN',
            'Nome' : 'sNom',
            'Valuta' : 'sCurrency',
            'Information ratio da data a data' : 'nIR',
            'TEV da data a data' : 'nTEV',
            'Sharpe ratio da data a data' : 'nSharpe',
            'Volatilità da data a data' : 'nVolat',
            'Sortino ratio da data a data' : 'nSortino',
            'DSR da data a data' : 'nDSR',
            'Perf Ann. da data a data' : 'nReta',
        }
        input_value = {"listModules":[]}
        INPUT_ID = 'Contenu_Contenu_cColonnesSelector_fieldJson'
        for index, indicatore in enumerate(indicatori):
            # controlla se gli indicatori sono nel dizionario altrimenti invia messaggio in console
            key = key_val_items.get(indicatore)
            value = indicatore
            key_value_item = {"key":key,"value":value,"uid":index+1}
            input_value["listModules"].append(key_value_item)
        driver.execute_script(f"let input = document.getElementById('{INPUT_ID}'); input.value = JSON.stringify({input_value});")

    def carica_lista(self, driver, list_name, directory_input_liste, list_file):
        """
        Raggiungi https://www.quantalys.it/Listes da ovunque e crea una lista
        con il nome 'list_name' e i prodotti all'interno di 'list_file'.

        Arguments:
            list_name {str} = nome da assegnare alla lista
            directory_input_liste {Path} = percorso in cui trovare il file con la lista
            list_file {str} = nome del file da cercare in self.directory_input_liste

        Return:
            id_lista {int} = identificativo della lista nell'account
            numero_fondi {int} = numero fondi della lista caricati sulla piattaforma
        """
        # Liste / Tools -> Liste
        """Tenta di cliccare subito su liste. Non funziona la prima volta ma tutte le rimanenti."""
        try:
            driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Liste').click()
        except NoSuchElementException:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Tools'))).click()
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Liste'))).click()

        # Crea nuova lista
        time.sleep(0.5)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.NAME, 'new'))).click()

        # Nome lista
        time.sleep(0.5)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, 'nom'))).send_keys(list_name)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="modal-new-liste"]/div/div/div[3]/button[2]'))
        ).click()

        # Importa prodotti
        _ = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="quantasearch"]/div[2]/div[3]/div/button[2]'))
        )
        time.sleep(0.5) # Necessario, preme il bottone troppo veloce e il sito non risponde
        _.click()
        # prendi la chiave unica della lista
        id_lista = driver.find_element(
            by=By.XPATH, value='/html/body/div[1]/div[3]/input[1]'
        ).get_attribute('value')

        # Scegli il file da importare
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, 'file'))
        ).send_keys(directory_input_liste.joinpath(list_file).__str__())

        # Importa lista
        _ = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="importForm"]/button')))
        time.sleep(0.5) # Necessario, va troppo veloce ed esporta liste vuote
        _.click()

        # Attendi il caricamento della lista
        WebDriverWait(driver,120).until(
            EC.text_to_be_present_in_element(
                (By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr/td'),
                'Nessun dato disponibile'
            )
        )
        WebDriverWait(driver,120).until_not(
            EC.text_to_be_present_in_element(
                (By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr/td'),
                'Nessun dato disponibile'
            )
        )
        # ottieni il numero totale di fondi caricati
        totale_fondi_lista = driver.find_element(by=By.ID, value='DataTables_Table_0_info').text.replace(',','')
        print(f'{totale_fondi_lista}\n')
        num_fondi_regex = re.compile(r'\d(\d)?(\d)?(\d)?')
        mo = num_fondi_regex.search(totale_fondi_lista)
        numero_fondi = mo.group()

        return id_lista, numero_fondi

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
            nome_primo_fondo = driver.find_element(
                by=By.XPATH, value='/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr[1]/td[2]'
            ).text
            # Quantalys non mette tutti gli li. Li aggiunge alla mano
            # a = driver.find_element(by=By.CSS_SELECTOR, value=list_class+' li:nth-child('+str(page)+') a')
            # L'unico modo di individuare l'anchor link che mi serve è selezionare l'anchor link in base al numero progressivo corrispondente alla pagina
            num_pagina = driver.find_element(by=By.LINK_TEXT, value=str(page))
            num_pagina.click()
            # Attendi che il tag li abbia l'attributo active (non funziona per il motivo nel commento sopra)
            # WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, list_class+' li:nth-child('+str(page)+')'))).get_attribute('active')
            # Attendi che il nome del primo elemento della nuova tabella sia diverso dal nome del primo elemento della tabella aggiornata
            WebDriverWait(driver, 20).until_not(
                EC.text_to_be_present_in_element(
                    (By.XPATH, '/html/body/div[1]/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/table/tbody/tr[1]/td[2]'), nome_primo_fondo
                )
            )
            # Scarica il nuovo dataframe
            element2 = driver.find_element(By.XPATH, table).get_attribute('outerHTML')
            df2 = pd.read_html(element2)[0]
            # Allegalo in coda al primo (df)
            df = pd.concat([df, df2], ignore_index=True)
        return df

    def to_confronto(self, driver, id_lista):
        """
        Raggiungi https://www.quantalys.it/compare/selection_fonds.aspx?univers=Fonds&menu=f da ovunque,
        carica la lista 'list_name' e arriva in confronto 
        https://www.quantalys.it/compare/comparaison_fonds.aspx?ID_Comparaison={ID}.

        Arguments:
            list_name {str} = nome della lista da caricare
        """

        # Confronto / Fondi -> Confronto
        """Tenta di cliccare subito su confronto. Quantalys preme il tasto fondi dopo il login."""
        try:
            driver.find_element(by=By.PARTIAL_LINK_TEXT, value='Confronto').click()
        except NoSuchElementException:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Fondi'))).click()
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Confronto'))).click()
            
        # Seleziona la lista usando il suo identificatore 'id_lista'
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, 'Contenu_Contenu_selectFonds_ctrlTreeListe_ddlDynatree'))
        ).click()
        json_file_liste_string = driver.find_element(
            By.ID, 'Contenu_Contenu_selectFonds_ctrlTreeListe_hidJson'
        ).get_attribute('value')
        json_file_liste_string = json_file_liste_string.replace('false', "'False'") # necessario per il passaggio successivo
        json_file_liste_list = eval(json_file_liste_string) # converte la stringa in una lista di dizionari
        _ = (__ for __ in json_file_liste_list if __['key'] == str(id_lista)) # scegli il dizionario che contiene l'id della lista appena caricata
        nome_lista = next(_)['title'] # ricava il nome della lista
        # seleziona la lista
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, nome_lista))
        ).click()
        # vai
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, 'Contenu_Contenu_selectFonds_ctrlTreeListe_hypValider'))
        ).click()
        # cerca
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'Contenu_Contenu_selectFonds_searchButton'))).click()
        # seleziona tutti
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'Contenu_Contenu_selectFonds_listeFonds_HeaderButton'))).click()
        # confronta
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Contenu_Contenu_btnComparer1"]'))).click()