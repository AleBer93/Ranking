import os
import pandas as pd
import time
with os.add_dll_directory('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\_blpapi'):
    import blpapi
from xbbg import blp
import datetime
import dateutil.relativedelta

class Completo():
    """
    Creazione file completo
    """

    def __init__(self, intermediario, t1, directory_output_liste_complete="C:/Users/Administrator/Desktop/Sbwkrq/Ranking/export_liste_complete_from_Q", file_completo='completo.csv', file_bloomberg='scarico_bloomberg.csv', directory_input_liste='./import_liste_into_Q/'):
        """
        Initialize the file.

        Parameters:
        t0_1Y = data iniziale un anno fa
        t0_3Y = data iniziale tre anni fa
        intermediario = intermediario a cui è destinata l'analisi
        t1 = data finale
        directory_output_liste_complete = percorso in cui trovare i dati scaricati delle liste complete
        file_completo = file da elaborare
        indicatore_bs = colonna contenente l'indicatore B&S
        file_bloomberg = file in cui scaricare le date di avvio dei fondi
        directory_input_liste = percorso in cui salvare le liste
        """
        self.intermediario = intermediario
        self.t1 = t1
        # self.t0_3Y = (datetime.datetime.strptime(t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+3))
        # self.t0_1Y = (datetime.datetime.strptime(t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+1))
        self.directory_output_liste_complete = directory_output_liste_complete
        self.file_completo = file_completo
        self.file_bloomberg = file_bloomberg
        self.directory_input_liste = directory_input_liste

    def concatenazione_file_excel(self):
        """
        Concatena i file excel all'interno della directory_output_liste_complete l'uno sotto l'altro.
        Salva il risultato con il nome file_name.

        Parameters:
        /
        """
        df = pd.concat((pd.read_csv(self.directory_output_liste_complete + '/' + filename, sep = ';', decimal=',', engine='python', encoding = "utf_16_le", error_bad_lines=False, skipfooter=1) for filename in os.listdir(self.directory_output_liste_complete)), ignore_index=True)
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def correzione_micro_russe(self):
        """
        Corregge le righe delle microcategorie Az. Paesi Emerg. Europa e Russia & Az. Paesi Emerg. Europa ex Russia perchè vanno a capo dalla sesta colonna in poi.
        
        Parameters:
        /
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

    def change_datatype(self, **colonne):
        """
        Cambia il tipo di dato alle colonne selezionate del file completo.
        
        Parameters:
        colonne(dict) = dizionario di colonne a cui cambiare il dato. Key=colonna, value=tipo dato(float, int, string)
        """
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        for key, value in colonne.items():
            df[key] = df[key].astype(value, errors='ignore')
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def seleziona_colonne(self, *colonne):
        """
        Seleziona le colonne desiderate dal file_csv con separatore ";" e decimali ","

        Parameters:
        file_csv(str) = file da cui estrarre le colonne
        colonne(tuple) = tuple di colonne da estrarre dal file
        """
        if self.intermediario == 'BPPB' or self.intermediario == 'BPL' or self.intermediario == 'CRV':
            colonne = 'Codice ISIN', 'Valuta', 'Nome del fondo', 'Categoria Quantalys', 'Rischio 1 anno fine mese', 'Rischio 3 anni") fine mese', 'Info 3 anni") fine mese', 'Alpha 3 anni") fine mese', 'SRRI'
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        df = df.loc[:, colonne]
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def merge_files(self, file_excel, left_on='Codice ISIN', right_on='isin', merge='left'):
        """
        Unisce il file completo csv e un secondo file excel o csv con il tipo di merge specificato.

        Parameters:
        file_excel(str) = file da unire al primo
        left_on(str) = colonna del primo df
        right_on(str) = colonna del secondo df
        merge(str) = specifica tipo di merge
        """
        df_1 = pd.read_csv(self.file_completo, sep=";", decimal=',',index_col=None)
        if file_excel.endswith('.xlsx'):
            df_2 = pd.read_excel(file_excel)
        elif file_excel.endswith('.csv'):
            df_2 = pd.read_csv(file_excel)
        df_merged = pd.merge(df_1, df_2, left_on=left_on, right_on=right_on, how=merge)
        print(f"\nIl primo file contiene {len(df_1)} fondi, mentre il secondo ne contiene {len(df_2)} fondi.\
        \nL'unione dei due file contiene {len(df_merged)} fondi.\n")
        df_merged.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def assegna_macro(self):
        """Assegna una macrocategoria ad ogni microcategoria."""
        BPPB_dict = {'Monetari Euro' : 'LIQ', 'Monetari Euro dinamici' : 'LIQ', 'Monet. altre valute europee' : 'LIQ', 'Monetari altre valute europ' : 'LIQ',
            'Obblig. euro gov. breve termine' : 'OBB_BT', 'Obblig. Euro breve term.' : 'OBB_BT', 'Obblig. Euro a scadenza' : 'OBB_BT',
            'Obblig. Euro gov. medio termine' : 'OBB_MLT', 'Obblig. Euro gov. lungo termine' : 'OBB_MLT', 'Obblig. Euro lungo termine' : 'OBB_MLT', 
            'Obblig. Euro medio term.' : 'OBB_MLT', 'Obblig. Euro gov.' : 'OBB_MLT', 'Obblig. Euro all maturities' : 'OBB_MLT', 'Obblig. Europa High Yield' : 'OBB_MLT',
            'Obblig. Europa' : 'OBB_MLT', 'Obblig. Sterlina inglese' : 'OBB_MLT', 'Obblig. Franco svizzero' : 'OBB_MLT', 'Obblig. Indicizz. Inflation Linked' : 'OBB_MLT', 
            'Obblig. Euro corporate' : 'OBB_CORP', 'Obblig. Euro high yield' : 'OBB_CORP',
            'Obblig. paesi emerg. Asia' : 'OBB_EM', 'Obblig. paesi emerg. Europa' : 'OBB_EM', 'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 'Obblig. Paesi Emerg.' : 'OBB_EM', 'Obblig. paesi emerg. a scadenza' : 'OBB_EM',
            'Obblig. Paesi Emerg. Hard Currency' : 'OBB_EM', 'Obblig. Paesi Emerg. Local Currency' : 'OBB_EM',
            'Obblig. Dollaro US breve term.' : 'OBB_GLOB', 'Obblig. USD medio-lungo term.' : 'OBB_GLOB', 'Obblig. Dollaro US medio-lungo term.' : 'OBB_GLOB',
            'Obblig. USD corporate' : 'OBB_GLOB', 'Obblig. Dollaro US corporate' : 'OBB_GLOB', 'Obblig. Doll. US all maturities' : 'OBB_GLOB',
            'Obblig. Dollaro US all mat' : 'OBB_GLOB', 'Obblig. Dollaro US high yield' : 'OBB_GLOB', 'Obblig. Asia' : 'OBB_GLOB', 'Obblig. globale' : 'OBB_GLOB',
            'Obblig. globale corporate' : 'OBB_GLOB', 'Obblig. globale high yield' : 'OBB_GLOB', 'Obblig. Yen' : 'OBB_GLOB', 'Obblig. altre valute' : 'OBB_GLOB',
            "Obblig. Indicizz. all'inflaz. USD" : 'OBB_GLOB', 'Obblig. Global Inflation Linked' : 'OBB_GLOB', 'Monetari Dollaro USA' : 'OBB_GLOB',
            'Monet. ex Europa altre valute' : 'OBB_GLOB', 'Monetari ex Europa altre valute' : 'OBB_GLOB',
            'Az. Europa' : 'AZ_EUR', 'Az. Area Euro' : 'AZ_EUR', 'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Area Euro Growth' : 'AZ_EUR',
            'Az. Area Euro Value' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 'Az. Europa Value' : 'AZ_EUR',
            'Az. Belgio' : 'AZ_EUR', 'Az. Francia' : 'AZ_EUR', 'Az. Francia small cap' : 'AZ_EUR', 'Az. Germania' : 'AZ_EUR', 'Az. Germania small cap' : 'AZ_EUR',
            'Az. Spagna' : 'AZ_EUR', 'Az. Paesi Bassi' : 'AZ_EUR', 'Az. Italia' : 'AZ_EUR', 'Az. UK' : 'AZ_EUR', 'Az. UK small cap' : 'AZ_EUR', 'Az. Svizzera' : 'AZ_EUR',
            'Az.Svizzera small cap' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Europa altri paesi' : 'AZ_EUR',
            'Azionario USA' : 'AZ_NA', 'Az. USA' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 'Az. USA Growth' : 'AZ_NA', 'Az. USA Value' : 'AZ_NA', 'Az. Canada' : 'AZ_NA',
            'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 'Az. Giappone small cap' : 'AZ_PAC', 'Az. Pacifico' : 'AZ_PAC',
            'Az. Brasile' : 'AZ_EM', 'Az. Cina' : 'AZ_EM', 'Az. India' : 'AZ_EM', 'Az. Russia' : 'AZ_EM', 'Az. Altri paesi emerg.' : 'AZ_EM',
            'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 'Az. Paesi Emerg. Europa ex Russia' : 'AZ_EM', 'Az. paesi emerg. Asia' : 'AZ_EM', 'Az. BRIC' : 'AZ_EM',
            'Az. Grande Cina' : 'AZ_EM', 'Az. paesi emerg. America Latina' : 'AZ_EM', 'Az. paesi emerg. altre zone' : 'AZ_EM', 'Az. paesi emerg. Mondo' : 'AZ_EM',
            'Commodities a leva' : 'OPP', 'Commodities Bear' : 'OPP', 'Commodities' : 'OPP', 'Obblig. Convertib. Euro' : 'OPP', 'Obblig. Convertib. Europa' : 'OPP', 
            'Obblig. Convertib. Dollaro US' : 'OPP', 'Obblig. Convertib. Glob.' : 'OPP', 'Az. real estate Europa' : 'OPP', 'Az. Biotech' : 'OPP',
            'Az. beni di consumo' : 'OPP', 'Az. ambiente' : 'OPP', 'Az. energia, materie prime, oro' : 'OPP', 'Az. energia. materie prime. oro' : 'OPP',
            'Az. energia materie prime oro' : 'OPP', 'Az. real estate Mondo' : 'OPP', 'Az. industria' : 'OPP', 'Az. salute   farmaceutico' : 'OPP',
            'Az. salute – farmaceutico' : 'OPP', 'Az. salute - farmaceutico' : 'OPP', 'Az. Servizi di pubblica utilita' : 'OPP', 'Az. servizi finanziari' : 'OPP',
            'Az. tecnologia' : 'OPP', 'Az. telecomunicazioni' : 'OPP', 'Az. Oro' : 'OPP', 'Az. Bear' : 'OPP', 'Obblig. Bear' : 'OPP', 'Valuta Long/Short' : 'OPP',
            'Altri' : 'OPP',
            'Bilanc. Prud. Europa' : 'FLEX', 'Bilanc. Prud. Global Euro' : 'FLEX', 'Bilanc. Prud. Dollaro US' : 'FLEX', 'Bilanc. Prud. Global' : 'FLEX',
            'Bilanc. Prud. altre valute' : 'FLEX', 'Bilanc. Equilib. Europa' : 'FLEX', 'Bilanc. Equil. Global Euro' : 'FLEX', 'Bilanc. Equil. Dollaro US' : 'FLEX',
            'Bilanc. Equil. Global' : 'FLEX', 'Bilanc. Equil. altre valute' : 'FLEX', 'Bilanc. Aggress. Europa' : 'FLEX', 'Bilanc. Aggress. Global Euro' : 'FLEX', 
            'Bilanc. aggress. Dollaro US' : 'FLEX', 'Bilanc. Aggress. Global' : 'FLEX', 'Bilanc. Aggress. altre valute' : 'FLEX', 'Flessibili Europa' : 'FLEX', 
            'Fless. Global Euro' : 'FLEX', 'Flessibili prudenti Europa' : 'FLEX', 'Flessibili Dollaro US' : 'FLEX', 'Flessibili prudenti globale' : 'FLEX',
            'Fless. Global' : 'FLEX', 'Fondi a scadenza pred. Euro' : 'FLEX', 'Fondi a scadenza pred. altre valute' : 'FLEX', 'Perf. ass. Dividendi' : 'FLEX', 
            'Perf. Ass. Arbitr.Fus.-acquis. Euro' : 'FLEX', 'Perf. assoluta strategia valute' : 'FLEX', 'Perf. assoluta Market Neutral Euro' : 'FLEX', 
            'Perf. ass. Long/Short eq.' : 'FLEX', 'Perf. assoluta tassi' : 'FLEX', 'Perf. assoluta volatilita' : 'FLEX', 'Perf. assoluta multi-strategia' : 'FLEX', 
            'Perf. assoluta (GBP)' : 'FLEX', 'Perf. ass. USD' : 'FLEX', 'Fondi  a garanzia o a formula Euro' : 'FLEX', 'Az. globale' : 'FLEX', 
            'Az. globale small cap' : 'FLEX', 'Az. globale Growth' : 'FLEX', 'Az. globale Value' : 'FLEX',
            }
        BPL_dict = {'Monetari Euro' : 'LIQ', 'Monetari Euro dinamici' : 'LIQ', 
            'Monet. ex Europa altre valute' : 'LIQ_FOR', 'Monetari ex Europa altre valute' : 'LIQ_FOR', 'Monet. altre valute europee' : 'LIQ_FOR', 
            'Monetari altre valute europ' : 'LIQ_FOR', 'Monetari Dollaro USA' : 'LIQ_FOR', 
            'Obblig. euro gov. breve termine' : 'OBB_BT', 'Obblig. Euro breve term.' : 'OBB_BT', 
            'Obblig. Euro gov. medio termine' : 'OBB_MLT', 'Obblig. Euro gov. lungo termine' : 'OBB_MLT', 'Obblig. Euro lungo termine' : 'OBB_MLT', 
            'Obblig. Euro medio term.' : 'OBB_MLT', 'Obblig. Euro gov.' : 'OBB_MLT', 'Obblig. Euro all maturities' : 'OBB_MLT',  'Obblig. Euro a scadenza' : 'OBB_MLT', 
            'Obblig. Indicizz. Inflation Linked' : 'OBB_MLT', 'Obblig. Convertib. Euro' : 'OBB_MLT', 'Fondi a scadenza pred. Euro' : 'OBB_MLT', 
            'Obblig. Europa' : 'OBB_EUR', 'Obblig. Sterlina inglese' : 'OBB_EUR', 'Obblig. Franco svizzero' : 'OBB_EUR', 'Obblig. Convertib. Europa' : 'OBB_EUR', 
            'Obblig. Euro corporate' : 'OBB_CORP', 
            'Obblig. paesi emerg. Asia' : 'OBB_EM', 'Obblig. paesi emerg. Europa' : 'OBB_EM',  'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 'Obblig. Paesi Emerg.' : 'OBB_EM', 
            'Obblig. paesi emerg. a scadenza' : 'OBB_EM', 'Obblig. Paesi Emerg. Hard Currency' : 'OBB_EM', 'Obblig. Paesi Emerg. Local Currency' : 'OBB_EM',
            'Obblig. Dollaro US breve term.' : 'OBB_USA', 'Obblig. USD medio-lungo term.' : 'OBB_USA', 'Obblig. Dollaro US medio-lungo term.' : 'OBB_USA', 
            'Obblig. USD corporate' : 'OBB_USA', 'Obblig. Dollaro US corporate' : 'OBB_USA', 'Obblig. Doll. US all maturities' : 'OBB_USA',
            'Obblig. Dollaro US all mat' : 'OBB_USA', 'Obblig. Convertib. Dollaro US' : 'OBB_USA', "Obblig. Indicizz. all'inflaz. USD" : 'OBB_USA',
            'Obblig. Asia' : 'OBB_GLOB', 'Obblig. globale' : 'OBB_GLOB', 'Obblig. globale corporate' : 'OBB_GLOB', 'Obblig. Yen' : 'OBB_GLOB', 
            'Obblig. altre valute' : 'OBB_GLOB', 'Obblig. Global Inflation Linked' : 'OBB_GLOB', 'Obblig. Convertib. Glob.' : 'OBB_GLOB', 
            'Fondi a scadenza pred. altre valute' : 'OBB_GLOB', 
            'Obblig. Euro high yield' : 'OBB_GLOB_HY', 'Obblig. Europa High Yield' : 'OBB_GLOB_HY', 'Obblig. Dollaro US high yield' : 'OBB_GLOB_HY', 
            'Obblig. globale high yield' : 'OBB_GLOB_HY',
            'Az. Europa' : 'AZ_EUR', 'Az. Area Euro' : 'AZ_EUR', 'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Area Euro Growth' : 'AZ_EUR',
            'Az. Area Euro Value' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 'Az. Europa Value' : 'AZ_EUR',
            'Az. Belgio' : 'AZ_EUR', 'Az. Francia' : 'AZ_EUR', 'Az. Francia small cap' : 'AZ_EUR', 'Az. Germania' : 'AZ_EUR', 'Az. Germania small cap' : 'AZ_EUR',
            'Az. Spagna' : 'AZ_EUR', 'Az. Paesi Bassi' : 'AZ_EUR', 'Az. Italia' : 'AZ_EUR', 'Az. UK' : 'AZ_EUR', 'Az. UK small cap' : 'AZ_EUR', 'Az. Svizzera' : 'AZ_EUR',
            'Az.Svizzera small cap' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Europa altri paesi' : 'AZ_EUR',
            'Azionario USA' : 'AZ_NA', 'Az. USA' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 'Az. USA Growth' : 'AZ_NA', 'Az. USA Value' : 'AZ_NA', 'Az. Canada' : 'AZ_NA',
            'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 'Az. Giappone small cap' : 'AZ_PAC', 'Az. Pacifico' : 'AZ_PAC',
            'Az. Brasile' : 'AZ_EM', 'Az. Cina' : 'AZ_EM', 'Az. India' : 'AZ_EM', 'Az. Russia' : 'AZ_EM', 'Az. Altri paesi emerg.' : 'AZ_EM', 
            'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 'Az. Paesi Emerg. Europa ex Russia' : 'AZ_EM', 'Az. paesi emerg. Asia' : 'AZ_EM', 'Az. BRIC' : 'AZ_EM', 
            'Az. Grande Cina' : 'AZ_EM', 'Az. paesi emerg. America Latina' : 'AZ_EM', 'Az. paesi emerg. altre zone' : 'AZ_EM', 'Az. paesi emerg. Mondo' : 'AZ_EM',
            'Az. globale' : 'AZ_GLOB', 'Az. globale small cap' : 'AZ_GLOB', 'Az. globale Growth' : 'AZ_GLOB', 'Az. globale Value' : 'AZ_GLOB',
            'Commodities a leva' : 'OPP', 'Commodities Bear' : 'OPP', 'Commodities' : 'OPP', 'Az. real estate Europa' : 'OPP', 'Az. Biotech' : 'OPP', 
            'Az. beni di consumo' : 'OPP', 'Az. ambiente' : 'OPP', 'Az. energia, materie prime, oro' : 'OPP', 'Az. energia. materie prime. oro' : 'OPP', 
            'Az. energia materie prime oro' : 'OPP', 'Az. real estate Mondo' : 'OPP', 'Az. industria' : 'OPP', 'Az. salute   farmaceutico' : 'OPP', 
            'Az. salute – farmaceutico' : 'OPP', 'Az. salute - farmaceutico' : 'OPP', 'Az. Servizi di pubblica utilita' : 'OPP', 'Az. servizi finanziari' : 'OPP', 
            'Az. tecnologia' : 'OPP', 'Az. telecomunicazioni' : 'OPP', 'Az. Oro' : 'OPP', 'Az. Bear' : 'OPP', 'Obblig. Bear' : 'OPP', 'Altri' : 'OPP',
            'Perf. ass. Dividendi' : 'OPP', 'Perf. Ass. Arbitr.Fus.-acquis. Euro' : 'OPP', 'Perf. assoluta strategia valute' : 'OPP',
            'Perf. assoluta Market Neutral Euro' : 'OPP', 'Perf. ass. Long/Short eq.' : 'OPP', 'Perf. assoluta tassi' : 'OPP', 'Perf. assoluta volatilita' : 'OPP',
            'Perf. assoluta multi-strategia' : 'OPP', 'Perf. assoluta (GBP)' : 'OPP', 'Perf. ass. USD' : 'OPP', 'Fondi  a garanzia o a formula Euro' : 'OPP', 
            'Valuta Long/Short' : 'OPP',
            'Bilanc. Prud. Europa' : 'BIL', 'Bilanc. Prud. Global Euro' : 'BIL', 'Bilanc. Prud. Dollaro US' : 'BIL', 'Bilanc. Prud. Global' : 'BIL', 
            'Bilanc. Prud. altre valute' : 'BIL', 'Bilanc. Equilib. Europa' : 'BIL', 'Bilanc. Equil. Global Euro' : 'BIL', 'Bilanc. Equil. Dollaro US' : 'BIL', 
            'Bilanc. Equil. Global' : 'BIL', 'Bilanc. Equil. altre valute' : 'BIL', 'Bilanc. Aggress. Europa' : 'BIL', 'Bilanc. Aggress. Global Euro' : 'BIL', 
            'Bilanc. aggress. Dollaro US' : 'BIL', 'Bilanc. Aggress. Global' : 'BIL', 'Bilanc. Aggress. altre valute' : 'BIL',
            'Flessibili Europa' : 'FLEX', 'Fless. Global Euro' : 'FLEX', 'Flessibili prudenti Europa' : 'FLEX', 'Flessibili Dollaro US' : 'FLEX', 
            'Flessibili prudenti globale' : 'FLEX', 'Fless. Global' : 'FLEX',
            }      
        CRV_dict = {'Monetari Euro' : 'LIQ', 'Monetari Euro dinamici' : 'LIQ', 'Monet. altre valute europee' : 'LIQ', 'Monetari altre valute europ' : 'LIQ',
            'Obblig. euro gov. breve termine' : 'OBB_BT', 'Obblig. Euro breve term.' : 'OBB_BT', 
            'Obblig. Euro gov. medio termine' : 'OBB_MLT', 'Obblig. Euro gov. lungo termine' : 'OBB_MLT', 'Obblig. Euro lungo termine' : 'OBB_MLT', 
            'Obblig. Euro medio term.' : 'OBB_MLT', 'Obblig. Euro gov.' : 'OBB_MLT', 'Obblig. Euro all maturities' : 'OBB_MLT', 'Obblig. Europa' : 'OBB_MLT',
            'Obblig. Sterlina inglese' : 'OBB_MLT', 'Obblig. Franco svizzero' : 'OBB_MLT', 'Obblig. Indicizz. Inflation Linked' : 'OBB_MLT', 
            'Obblig. Euro a scadenza' : 'OBB_MLT',
            'Obblig. Euro corporate' : 'OBB_CORP',
            'Obblig. paesi emerg. Asia' : 'OBB_EM', 'Obblig. paesi emerg. Europa' : 'OBB_EM', 'Obblig. Paesi Emerg. Europa' : 'OBB_EM', 
            'Obblig. Paesi Emerg.' : 'OBB_EM', 'Obblig. paesi emerg. a scadenza' : 'OBB_EM', 'Obblig. Paesi Emerg. Hard Currency' : 'OBB_EM', 
            'Obblig. Paesi Emerg. Local Currency' : 'OBB_EM',
            'Obblig. Dollaro US breve term.' : 'OBB_GLOB', 'Obblig. USD medio-lungo term.' : 'OBB_GLOB', 'Obblig. Dollaro US medio-lungo term.' : 'OBB_GLOB', 
            'Obblig. USD corporate' : 'OBB_GLOB', 'Obblig. Dollaro US corporate' : 'OBB_GLOB', 'Obblig. Doll. US all maturities' : 'OBB_GLOB', 
            'Obblig. Dollaro US all mat' : 'OBB_GLOB', 'Obblig. Asia' : 'OBB_GLOB', 'Obblig. globale' : 'OBB_GLOB', 'Obblig. globale corporate' : 'OBB_GLOB', 
            'Obblig. Yen' : 'OBB_GLOB', 'Obblig. altre valute' : 'OBB_GLOB', "Obblig. Indicizz. all'inflaz. USD" : 'OBB_GLOB', 
            'Obblig. Global Inflation Linked' : 'OBB_GLOB', 'Monetari Dollaro USA' : 'OBB_GLOB', 'Monet. ex Europa altre valute' : 'OBB_GLOB', 
            'Monetari ex Europa altre valute' : 'OBB_GLOB',
            'Obblig. Euro high yield' : 'OBB_GLOB_HY', 'Obblig. Europa High Yield' : 'OBB_GLOB_HY', 'Obblig. Dollaro US high yield' : 'OBB_GLOB_HY', 
            'Obblig. globale high yield' : 'OBB_GLOB_HY',
            'Az. Europa' : 'AZ_EUR', 'Az. Area Euro' : 'AZ_EUR', 'Az. Area Euro small cap' : 'AZ_EUR', 'Az. Area Euro Growth' : 'AZ_EUR', 
            'Az. Area Euro Value' : 'AZ_EUR', 'Az. Europa small cap' : 'AZ_EUR', 'Az. Europa Growth' : 'AZ_EUR', 'Az. Europa Value' : 'AZ_EUR',
            'Az. Belgio' : 'AZ_EUR', 'Az. Francia' : 'AZ_EUR', 'Az. Francia small cap' : 'AZ_EUR', 'Az. Germania' : 'AZ_EUR', 'Az. Germania small cap' : 'AZ_EUR',
            'Az. Spagna' : 'AZ_EUR', 'Az. Paesi Bassi' : 'AZ_EUR', 'Az. Italia' : 'AZ_EUR', 'Az. UK' : 'AZ_EUR', 'Az. UK small cap' : 'AZ_EUR', 
            'Az. Svizzera' : 'AZ_EUR', 'Az.Svizzera small cap' : 'AZ_EUR', 'Az. paesi nordici' : 'AZ_EUR', 'Az. Europa altri paesi' : 'AZ_EUR',
            'Azionario USA' : 'AZ_NA', 'Az. USA' : 'AZ_NA', 'Az. USA small cap' : 'AZ_NA', 'Az. USA Growth' : 'AZ_NA', 'Az. USA Value' : 'AZ_NA', 
            'Az. Canada' : 'AZ_NA',
            'Az. Asia Pacifico ex Giapp.' : 'AZ_PAC', 'Az. Giappone' : 'AZ_PAC', 'Az. Giappone small cap' : 'AZ_PAC', 'Az. Pacifico' : 'AZ_PAC',
            'Az. Brasile' : 'AZ_EM', 'Az. Cina' : 'AZ_EM', 'Az. India' : 'AZ_EM', 'Az. Russia' : 'AZ_EM', 'Az. Altri paesi emerg.' : 'AZ_EM', 
            'Az. Paesi Emerg. Europa e Russia' : 'AZ_EM', 'Az. Paesi Emerg. Europa ex Russia' : 'AZ_EM', 'Az. paesi emerg. Asia' : 'AZ_EM', 'Az. BRIC' : 'AZ_EM', 
            'Az. Grande Cina' : 'AZ_EM', 'Az. paesi emerg. America Latina' : 'AZ_EM', 'Az. paesi emerg. altre zone' : 'AZ_EM', 'Az. paesi emerg. Mondo' : 'AZ_EM',
            'Az. globale' : 'AZ_GLOB', 'Az. globale small cap' : 'AZ_GLOB', 'Az. globale Growth' : 'AZ_GLOB', 'Az. globale Value' : 'AZ_GLOB',
            'Bilanc. Prud. Europa' : 'FLEX', 'Bilanc. Prud. Global Euro' : 'FLEX', 'Bilanc. Prud. Dollaro US' : 'FLEX', 'Bilanc. Prud. Global' : 'FLEX', 
            'Bilanc. Prud. altre valute' : 'FLEX', 'Bilanc. Equilib. Europa' : 'FLEX', 'Bilanc. Equil. Global Euro' : 'FLEX', 'Bilanc. Equil. Dollaro US' : 'FLEX', 
            'Bilanc. Equil. Global' : 'FLEX', 'Bilanc. Equil. altre valute' : 'FLEX', 'Bilanc. Aggress. Europa' : 'FLEX', 'Bilanc. Aggress. Global Euro' : 'FLEX', 
            'Bilanc. aggress. Dollaro US' : 'FLEX', 'Bilanc. Aggress. Global' : 'FLEX', 'Bilanc. Aggress. altre valute' : 'FLEX', 'Flessibili Europa' : 'FLEX', 
            'Fless. Global Euro' : 'FLEX', 'Flessibili prudenti Europa' : 'FLEX', 'Flessibili Dollaro US' : 'FLEX', 'Flessibili prudenti globale' : 'FLEX', 
            'Fless. Global' : 'FLEX',
            'Commodities a leva' : 'OPP', 'Commodities Bear' : 'OPP', 'Commodities' : 'OPP', 'Obblig. Convertib. Euro' : 'OPP', 'Obblig. Convertib. Europa' : 'OPP', 
            'Obblig. Convertib. Dollaro US' : 'OPP', 'Obblig. Convertib. Glob.' : 'OPP', 'Az. real estate Europa' : 'OPP', 'Az. Biotech' : 'OPP', 
            'Az. beni di consumo' : 'OPP', 'Az. ambiente' : 'OPP', 'Az. energia, materie prime, oro' : 'OPP', 'Az. energia. materie prime. oro' : 'OPP', 
            'Az. energia materie prime oro' : 'OPP', 'Az. real estate Mondo' : 'OPP', 'Az. industria' : 'OPP', 'Az. salute   farmaceutico' : 'OPP', 
            'Az. salute – farmaceutico' : 'OPP', 'Az. salute - farmaceutico' : 'OPP', 'Az. Servizi di pubblica utilita' : 'OPP', 'Az. servizi finanziari' : 'OPP', 
            'Az. tecnologia' : 'OPP', 'Az. telecomunicazioni' : 'OPP', 'Az. Oro' : 'OPP', 'Az. Bear' : 'OPP', 'Obblig. Bear' : 'OPP', 'Valuta Long/Short' : 'OPP', 
            'Altri' : 'OPP', 'Perf. ass. Dividendi' : 'OPP', 'Perf. Ass. Arbitr.Fus.-acquis. Euro' : 'OPP', 'Perf. assoluta strategia valute' : 'OPP', 
            'Perf. assoluta Market Neutral Euro' : 'OPP', 'Perf. ass. Long/Short eq.' : 'OPP', 'Perf. assoluta tassi' : 'OPP',
            'Perf. assoluta volatilita' : 'OPP', 'Perf. assoluta multi-strategia' : 'OPP', 'Perf. assoluta (GBP)' : 'OPP', 'Perf. ass. USD' : 'OPP', 
            'Fondi  a garanzia o a formula Euro' : 'OPP', 'Fondi a scadenza pred. Euro' : 'OPP', 'Fondi a scadenza pred. altre valute' : 'OPP',
            }
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        if self.intermediario == 'BPPB':
            df['macro_categoria'] = df['Categoria Quantalys'].map(BPPB_dict)
        elif self.intermediario == 'BPL':
            df['macro_categoria'] = df['Categoria Quantalys'].map(BPL_dict)
        elif self.intermediario == 'CRV':
            df['macro_categoria'] = df['Categoria Quantalys'].map(CRV_dict)
        print(f"Ci sono {df['macro_categoria'].isnull().sum()} fondi a cui non è stata assegnata una macro categoria.")
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def sconta_commissioni(self):
        """Sconta le commissioni dei fondi in base alla loro macro categoria"""
        if self.intermediario == 'CRV':
            df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
            sconti = {'LIQ' : 0.85, 'OBB_BT' : 0.35, 'OBB_MLT' : 0.35, 'OBB_CORP' : 0.35, 'OBB_EM' : 0.35, 'OBB_GLOB' : 0.35,
                'OBB_GLOB_HY' : 0.35, 'AZ_EUR' : 0.30, 'AZ_NA' : 0.30, 'AZ_PAC' : 0.30, 'AZ_EM' : 0.30, 'AZ_GLOB' : 0.30, 'FLEX' : 0.60,
                'OPP' : 0.50}
            df['commissione'] = df['commissione']*df['macro_categoria'].apply(lambda x : sconti[x])
            df.to_csv(self.file_completo, sep=";", decimal=',', index=False)
        else:
            pass

    def scarico_datadiavvio(self):
        # """ scarica da SQL il dataframe con tutte le date di avvio disponibili, e fai un merge con il file completo, poi """
        # from sqlalchemy import create_engine, MetaData, Table
        # from sqlalchemy.types import Float, DateTime
        # DATABASE_URL = 'postgres+psycopg2://postgres:bloomberg893@localhost:5432/ranking'
        # engine = create_engine(DATABASE_URL)
        # connection = engine.connect()
        # df_id = pd.read_sql("SELECT * FROM inception_date", connection)
        # print("\nSto scaricando le dati di avvio dei fondi da Bloomberg...")
        # df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        # df_merge_id = pd.merge(df, df_id, how='left', left_on='Codice ISIN', right_on='isin_code')
        # fondi_non_presenti = df_merge_id.loc[df_merge_id['fund_incept_dt'].isna(), ['Codice ISIN', 'valuta']]
        # print(fondi_non_presenti)
        # df_bl = blp.bdp('/isin/LU0048578792' + fondi_non_presenti['Codice ISIN'], flds="fund_incept_dt") #/isin/IT0001029823
        # print(df_bl)
        # df_bl.reset_index(inplace=True)
        # df_bl['isin_code'] = df_bl['index'].str[6:]
        # df_bl.reset_index(drop=True, inplace=True)
        # print(df_bl)
        # df_merged = pd.merge(df_merge_id, df_bl, left_on='Codice ISIN', right_on='isin_code', how='left')
        # df_merged.to_csv(self.file_completo, sep=";")
        # df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        # # TODO : SCARICA DA SQL LA DATA DI AVVIO E DA BLOOMBERG LE RIMANENTI. AGGIORNA QUELLE NON PRESENTI SU SQL
        # df.to_sql('persone_fisiche', con=engine, if_exists='replace', index=False, dtype={'data_questionario' : DateTime()})

        """
        Scarica la data di avvio dei fondi nel file_bloomberg utilizzando la libreria di Bloomberg.
        Aggiungi la data di avvio al file completo.
        """
        print("\nSto scaricando le dati di avvio dei fondi da Bloomberg...")
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        df_bl = blp.bdp('/isin/' + df['Codice ISIN'], flds="fund_incept_dt") #/isin/IT0001029823
        df_bl.reset_index(inplace=True)
        df_bl['isin_code'] = df_bl['index'].str[6:]
        df_bl.reset_index(drop=True, inplace=True)
        df_merged = pd.merge(df, df_bl, left_on='Codice ISIN', right_on='isin_code', how='left')
        print('scaricate!')
        fondi_senza_data_di_avvio = df_merged.loc[df_merged['fund_incept_dt'].isna(), ['Codice ISIN', 'Valuta', 'Nome del fondo']]
        print(f"\nCi sono {df_merged['fund_incept_dt'].isnull().sum()} fondi senza data di avvio:\n{fondi_senza_data_di_avvio}\n")
        df_merged.to_csv(self.file_completo, sep=";", decimal=',', index=False)
        df_bl.to_csv(self.file_bloomberg, sep=";")
        df = pd.read_csv('completo.csv', sep=";", decimal=',', index_col=None)
        while any(df_merged['fund_incept_dt'].isna()):
            print("Ci sono delle date mancanti, è necessario aggiornarle per l'analisi successiva,")
            _ = input(f'apri il file {self.file_completo}, aggiungi le date, poi premi enter\n')
            df_merged = pd.read_csv('completo.csv', sep=";", decimal=',', index_col=None)

    def indicatore_BS(self):
        # TODO : fai uno scarico da quantalys con benchamrk di default per tutti quei fondi che hanno alpha o IR pari a 0. L'alfa scaricato da quantalys è in percentuale...
        # quantalys assegna un valore pari a 0 all'information ratio se l'alpha è un numero del tipo 0.00*
        """
        Calcola l'indicatore B&S correggendo l'IR per i costi spalmati sugli anni di detenzione medi di un fondo.
        Formula = IR - (IR * fee) / (anni_detenzione * alpha)
        Formula = (IR * TEV - (fee / anni_detenzione)) / TEV
        Le colonne considerate ai fini del calcolo sono: 'Info 3 anni") fine mese', 'Alpha 3 anni") fine mese', 'commissione'
        """
        classi_a_benchmark_BPPB = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM']
        classi_a_benchmark_BPL = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_EUR', 'OBB_CORP', 'OBB_GLOB', 'OBB_USA', 'OBB_EM', 'OBB_GLOB_HY']
        classi_a_benchmark_CRV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM', 'OBB_GLOB_HY']
        if self.intermediario == 'BPPB':
            anni_detenzione = 3
            classi = classi_a_benchmark_BPPB
        elif self.intermediario == 'BPL':
            anni_detenzione = 5
            classi = classi_a_benchmark_BPL
        elif self.intermediario == 'CRV':
            anni_detenzione = 3
            classi = classi_a_benchmark_CRV
            return None
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        while any(df['Info 3 anni") fine mese']==0) or any(df['Alpha 3 anni") fine mese']==0):
            print("Ci sono dei fondi con alpha o information ratio uguale a 0, è necessario aggiornarli per l'analisi successiva,")
            _ = input(f'apri il file {self.file_completo}, aggiorna i dati, poi premi enter\n')
            df = pd.read_csv('completo.csv', sep=";", decimal=',', index_col=None)
        df['fund_incept_dt'] = pd.to_datetime(df['fund_incept_dt'], dayfirst=True)
        t0_3Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+3)).strftime('%d/%m/%Y') # data iniziale tre anni fa
        df.loc[(df['macro_categoria'].isin(classi)) & (df['fund_incept_dt'] < t0_3Y), 'BS_3_anni'] = df['Info 3 anni") fine mese'] - (df['Info 3 anni") fine mese'] * df['commissione']) / (int(anni_detenzione) * df['Alpha 3 anni") fine mese'])
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def calcolo_best_worst(self):
        """Calcolo best e worst per le categorie a benchmark, per fondi con più di tre anni e con indicatore B&S presente."""
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        # print(df['fund_incept_dt'].dtypes) # da oggetto
        df['fund_incept_dt'] = pd.to_datetime(df['fund_incept_dt'], dayfirst=True)
        #df['fund_incept_dt'] = df['fund_incept_dt'].astype('datetime64[ns]')
        # print(df['fund_incept_dt'].dtypes) # a datetime
        classi_a_benchmark_BPPB = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM']
        classi_a_benchmark_BPL = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_EUR', 'OBB_CORP', 'OBB_GLOB', 'OBB_USA', 'OBB_EM', 'OBB_GLOB_HY']
        classi_a_benchmark_CRV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM', 'OBB_GLOB_HY']
        t0_3Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+3)).strftime('%d/%m/%Y') # data iniziale tre anni fa
        # df2 = df[df['fund_incept_dt'] >= t0_3Y]
        # df2.to_csv('aaa.csv', sep=";", decimal=',', index=False)
        if self.intermediario == 'BPPB':
            for macro in classi_a_benchmark_BPPB:
                for micro in df.loc[df['macro_categoria'] == macro, 'Categoria Quantalys'].unique():
                    mediana = df.loc[(df['macro_categoria'] == macro) & (df['Categoria Quantalys'] == micro) & (df['fund_incept_dt'] < t0_3Y) & (df['BS_3_anni'].notnull()), 'BS_3_anni'].median()
                    df.loc[(df['macro_categoria'] == macro) & (df['Categoria Quantalys'] == micro) & (df['fund_incept_dt'] < t0_3Y) & (df['BS_3_anni'].notnull()), 'Best_Worst'] = df.loc[(df['macro_categoria'] == macro) & (df['Categoria Quantalys'] == micro) & (df['fund_incept_dt'] < t0_3Y) & (df['BS_3_anni'].notnull()), 'BS_3_anni'].apply(lambda x: 'worst' if x < mediana else 'best')
        elif self.intermediario == 'BPL':
            for macro in classi_a_benchmark_BPL:
                for micro in df.loc[df['macro_categoria'] == macro, 'Categoria Quantalys'].unique():
                    mediana = df.loc[(df['macro_categoria'] == macro) & (df['Categoria Quantalys'] == micro) & (df['fund_incept_dt'] < t0_3Y) & (df['BS_3_anni'].notnull()), 'BS_3_anni'].median()
                    df.loc[(df['macro_categoria'] == macro) & (df['Categoria Quantalys'] == micro) & (df['fund_incept_dt'] < t0_3Y) & (df['BS_3_anni'].notnull()), 'Best_Worst'] = df.loc[(df['macro_categoria'] == macro) & (df['Categoria Quantalys'] == micro) & (df['fund_incept_dt'] < t0_3Y) & (df['BS_3_anni'].notnull()), 'BS_3_anni'].apply(lambda x: 'worst' if x < mediana else 'best')
        elif self.intermediario == 'CRV':
            pass
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def sfdr(self):
        """Scarica il numero dell'articolo della disciplina europea SFDR"""
        print("\nSto scaricando l'articolo della disciplina SFDR...")
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        df_bl = blp.bdp('/isin/' + df['Codice ISIN'], flds="sfdr_classification")
        df_bl.reset_index(inplace=True)
        df_bl['isin_code'] = df_bl['index'].str[6:]
        df_bl.reset_index(drop=True, inplace=True)
        df_merged = pd.merge(df, df_bl, left_on='Codice ISIN', right_on='isin_code', how='left')
        df_merged["sfdr_classification"] = df_merged["sfdr_classification"].fillna(0)
        df_merged["sfdr_classification"] = pd.to_numeric(df_merged["sfdr_classification"], errors='coerce').astype(int)
        df_merged["sfdr_classification"].replace(0, '', inplace=True)
        print('scaricate!')
        df_merged.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def discriminazione_flessibili(self):
        """
        Assegna l'etichetta 'bassa_vola' se la volatilità a 3 anni è inferiore a 0.05 oppure 'media_alta_vola', ove disponibile,
        altrimenti se la volatilità a 1 anno è inferiore a 0.05 oppure 'media_alta_vola', ove disponibile,
        altrimenti se l'indicatore SRRI è inferiore a 3 oppure 'media_alta_vola', ove disponbile,
        altrimenti asssegna l'etichetta 'bassa_vola' ai fondi senza dati sul rischio.
        """
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        if self.intermediario == 'BPPB' or self.intermediario == 'CRV':
            df['categoria_flessibili'] = df.loc[(df['macro_categoria'] == 'FLEX') & (df['Rischio 3 anni") fine mese'].notnull()), 'Rischio 3 anni") fine mese'].apply(lambda x: 'bassa_vola' if x < 0.05 else 'media_alta_vola')
            df.loc[df['categoria_flessibili'].isnull(), 'categoria_flessibili'] = df.loc[(df['macro_categoria'] == 'FLEX') & (df['Rischio 1 anno fine mese'].notnull()), 'Rischio 1 anno fine mese'].apply(lambda x: 'bassa_vola' if x < 0.05 else 'media_alta_vola')
            df.loc[df['categoria_flessibili'].isnull(), 'categoria_flessibili'] = df.loc[(df['macro_categoria'] == 'FLEX') & (df['SRRI'].notnull()), 'SRRI'].apply(lambda x: 'bassa_vola' if x < 3 else 'media_alta_vola')
            print(f"\nI seguenti fondi flessibili non sono stati classificati:\n {df.loc[(df['macro_categoria'] == 'FLEX') & (df['categoria_flessibili'].isnull()), ['Codice ISIN', 'Nome del fondo', 'Categoria Quantalys']]}\n---Gli verrà assegnata la categoria bassa_volatilità.\n")
            df.loc[(df['macro_categoria'] == 'FLEX') & (df['categoria_flessibili'].isnull()), 'categoria_flessibili'] = 'bassa_vola'
        elif self.intermediario == 'BPL':
            pass
        df.to_csv(self.file_completo, sep=";", decimal=',', index=False)
    
    def seleziona_e_rinomina_colonne(self, file_excel):
        """
        Seleziona solo le colonne utili del file_excel con la funzione sopra definita seleziona_colonne.
        Rinomina le colonne del file_excel.

        Parameters:
        colonne(tuple) = tuple contenente i nuovi nomi da assegnare alle colonne precedentemente selezionate
        """
        if self.intermediario == 'BPPB':
            col_sel = ['Codice ISIN', 'Valuta', 'Nome del fondo', 'Categoria Quantalys', 'macro_categoria', 'fund_incept_dt',
                'commissione', 'BS_3_anni', 'Best_Worst', 'categoria_flessibili', 'fondo_a_finestra']
            col_ren = ['ISIN', 'valuta', 'nome', 'micro_categoria', 'macro_categoria', 'data_di_avvio', 'commissione', 'B&S_3Y',
                'Best_Worst', 'categoria_flessibili', 'fondo_a_finestra']
        elif self.intermediario == 'BPL':
            col_sel = ['Codice ISIN', 'Valuta', 'Nome del fondo', 'Categoria Quantalys', 'macro_categoria', 'fund_incept_dt',
                'commissione', 'BS_3_anni', 'Best_Worst']
            col_ren = ['ISIN', 'valuta', 'nome', 'micro_categoria', 'macro_categoria', 'data_di_avvio',
                'commissione', 'B&S_3Y', 'Best_Worst']
        elif self.intermediario == 'CRV':
            col_sel = ['Codice ISIN', 'Valuta', 'Nome del fondo', 'Categoria Quantalys', 'macro_categoria', 'fund_incept_dt',
                'categoria_flessibili', 'commissione']
            col_ren = ['ISIN', 'valuta', 'nome', 'micro_categoria', 'macro_categoria', 'data_di_avvio', 'categoria_flessibili',
                'commissione']
        df = pd.read_csv(file_excel, sep=";", decimal=',', index_col=None)
        df = df[col_sel]
        df.columns = col_ren
        df.to_csv(file_excel, sep=";", decimal=',', index=False)

    def creazione_liste_input(self):
        """
        Crea file csv, con dimensioni massime pari a ???199 elementi, da importare in Quantalys.it.
        Directory in cui vengono salvati i file : './import_liste_into_Q/'

        Parameters:
        /
        """
        df_com = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        if not os.path.exists(self.directory_input_liste):
            os.makedirs(self.directory_input_liste)
        for categoria in df_com['macro_categoria'].unique():
            chunks = len(df_com.loc[df_com['macro_categoria'] == categoria])//500 +1 # blocchi da 500 elementi
            for chunk in range(chunks):
                df = df_com.loc[df_com['macro_categoria'] == categoria, ['ISIN', 'valuta']]
                df = df.iloc[0 + 499 * chunk : 499 + 499 * chunk]
                df.columns = ['codice isin', 'divisa']
                df.to_csv(self.directory_input_liste + categoria + '_' + str(chunk) + '.csv', sep=";", index=False)


if __name__ == '__main__':
    start = time.time()
    _ = Completo(intermediario='BPPB', t1='31/01/2022')
    _.concatenazione_file_excel()
    _.correzione_micro_russe()
    _.change_datatype(SRRI = float)
    _.seleziona_colonne()
    _.merge_files('catalogo_fondi.xlsx')
    _.assegna_macro()
    _.sconta_commissioni()
    _.scarico_datadiavvio()
    _.indicatore_BS()
    _.calcolo_best_worst()
    _.sfdr()
    # _.discriminazione_flessibili()
    # _.seleziona_e_rinomina_colonne('completo.csv')
    # _.creazione_liste_input()
    end = time.time()
    print("Elapsed time: ", end - start, 'seconds')