import datetime
import os
import time
import zipfile
from pathlib import Path

import dateutil.relativedelta
import numpy as np
import pandas as pd
import win32com.client
from openpyxl import load_workbook  # Per caricare un libro
from openpyxl.styles import numbers  # Per cambiare i formati dei numeri
from openpyxl.styles import (Alignment, Border, Font,  # Per cambiare lo stile
                             PatternFill, Side)


class Ranking():

    def __str__(self):
        return "Ranking dei fondi"

    def __init__(self, intermediario, t1):
        """
        Arguments:
            intermediario {str} = intermediario a cui è destinata l'analisi
            t1 {datetime} = data di calcolo indici alla fine del mese
            file_catalogo {str} = file formattato
            file_completo {str} = file da elaborare
            file_ranking {str} = file in cui fare la rankizzazione
            directory_output_liste {WindowsPath} = percorso in cui scaricare i dati delle liste
            directory_input_liste_best {WindowsPath} = percorso in cui scaricare i dati delle liste con i soli best in class
            file_zip {str} = file zip in cui salvare il file di ranking e le note
        """
        self.intermediario = intermediario
        self.t1 = t1
        directory = Path().cwd()
        self.directory = directory
        self.directory_output_liste = self.directory.joinpath('docs', 'export_liste_from_Q')
        self.directory_input_liste_best = self.directory.joinpath('docs', 'import_liste_best_into_Q')
        self.file_catalogo = 'catalogo_fondi.xlsx'
        self.file_completo = 'completo.csv'
        self.file_ranking = 'ranking.xlsx'
        self.file_zip = 'rank.zip'
        self.soluzioni_BPPB = {'LIQ' : 1, 'OBB_BT' : 1, 'OBB_MLT' : 1, 'OBB_CORP' : 1, 'OBB_GLOB' : 1, 'OBB_EM' : 1, 'OBB_GLOB_HY' : 1, 
            'AZ_EUR' : 3, 'AZ_NA' : 3, 'AZ_PAC' : 3, 'AZ_EM' : 3}
        self.soluzioni_BPL = {'LIQ' : 3, 'OBB_BT' : 3, 'OBB_MLT' : 3, 'OBB_EUR' : 3, 'OBB_CORP' : 3, 'OBB_GLOB' : 3, 'OBB_USA' : 3, 
            'OBB_EM' : 3, 'OBB_GLOB_HY' : 3, 'AZ_EUR' : 3, 'AZ_NA' : 3, 'AZ_PAC' : 3, 'AZ_EM' : 3, 'AZ_GLOB' : 3}

    def ranking_per_grado(self, metodo):
        """
        Assegna un punteggio in ordine decrescente ai fondi delle micro categorie direzionali in base al loro indicatore corretto
        a 3 anni e ad 1 anno, discriminando in base al grado di gestione.
        """
        classi_a_benchmark_BPPB_metodo_doppio = {'LIQ' : 'Monetari Euro', 'OBB_BT' : 'Obblig. Euro breve term.', 
            'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale', 
            'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield', 'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 
            'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo'}
        classi_a_benchmark_BPL_metodo_doppio = {'LIQ' : 'Monetari Euro', 'OBB_BT' : 'Obblig. Euro breve term.', 
            'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_EUR' : 'Obblig. Europa', 'OBB_CORP' : 'Obblig. Euro corporate', 
            'OBB_GLOB' : 'Obblig. globale', 'OBB_USA' : 'Obblig. Dollaro US all mat', 'OBB_EM' : 'Obblig. Paesi Emerg.', 
            'OBB_GLOB_HY' : 'Obblig. globale high yield', 'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 
            'AZ_EM' : 'Az. paesi emerg. Mondo', 'AZ_GLOB' : 'Az. globale'}
        t0_3Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+3)).strftime('%d/%m/%Y') # data iniziale tre anni fa
        t0_1Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+1)).strftime('%d/%m/%Y') # data iniziale un anno fa
        if metodo == 'singolo':
            return None
        elif metodo == 'doppio':
            df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
            df['data_di_avvio'] = pd.to_datetime(df['data_di_avvio'], dayfirst=True)
            if self.intermediario == 'BPPB':
                macro_micro = classi_a_benchmark_BPPB_metodo_doppio
                soluzioni = self.soluzioni_BPPB
            elif self.intermediario == 'BPL':
                macro_micro = classi_a_benchmark_BPL_metodo_doppio
                soluzioni = self.soluzioni_BPL
            elif self.intermediario == 'CRV':
                return None

            for macro in list(macro_micro.keys()):
                for micro in df.loc[df['macro_categoria'] == macro, 'micro_categoria'].unique():
                    if micro in list(macro_micro.values()):
                        if soluzioni[macro] == 1:
                            for etichetta in ['best', 'worst']:
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'semi_attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'semi_attivo') & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False)
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'attivo') & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'semi_attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'].max()
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'molto_attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'molto_attivo') & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'].max()
                            for etichetta in ['best', 'worst']:
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'semi_attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'semi_attivo') & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False)
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'attivo') & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'semi_attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'].max()
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'molto_attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'molto_attivo') & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'].max()
                        elif soluzioni[macro] == 2:
                            for etichetta in ['best', 'worst']:
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'attivo') & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False)
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'semi_attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'semi_attivo') & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'].max()
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'molto_attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'molto_attivo') & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'semi_attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'].max()
                            for etichetta in ['best', 'worst']:
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'attivo') & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False)
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'semi_attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'semi_attivo') & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'].max()
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'molto_attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'molto_attivo') & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'semi_attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'].max()
                        elif soluzioni[macro] == 3:
                            for etichetta in ['best', 'worst']:
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])) & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])) & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False)
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'molto_attivo') & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'] == 'molto_attivo') & (df['Best_Worst_3Y'] == etichetta), 'BS_3_anni'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_3Y) & (df['BS_3_anni'].notnull()) & (df['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])) & (df['Best_Worst_3Y'] == etichetta), 'ranking_per_grado_3Y'].max()
                            for etichetta in ['best', 'worst']:
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])) & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])) & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False)
                                df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'molto_attivo') & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'] = df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'] == 'molto_attivo') & (df['Best_Worst_1Y'] == etichetta), 'BS_1_anno'].rank(method='first', na_option='keep', ascending=False) + df.loc[(df['macro_categoria'] == macro) & (df['micro_categoria'] == micro) & (df['data_di_avvio'] < t0_1Y) & (df['BS_1_anno'].notnull()) & (df['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])) & (df['Best_Worst_1Y'] == etichetta), 'ranking_per_grado_1Y'].max()

            df.to_csv(self.file_completo, sep=";", decimal=',', index=False)

    def merge_completo_liste(self):
        """
        Aggiunge gli indici scaricati da Quantalys.it nel file completo
        Passando dal percorso Fondi -> Confronto, alcuni fondi estinti o assorbiti vengono eslusi dalla lista caricata.
        Questo porta ad avere i dati di un numero di fondi inferiore a quelli caricati nelle liste.
        # TODO: Devo verificare se i fondi che non sono stati scaricati da Quantalys, non siano stati classificati come best nel processo
        precedente altrimenti sorgerebbe un problema.
        """
        df = pd.read_csv(self.file_completo, sep=";", decimal=',', index_col=None)
        print('sto aggiungendo gli indici delle liste al file completo...')

        df['Information_Ratio_3Y'] = np.nan
        df['TEV_3Y'] = np.nan
        df['Information_Ratio_1Y'] = np.nan
        df['TEV_1Y'] = np.nan
        df['Sortino_3Y'] = np.nan
        df['DSR_3Y'] = np.nan
        df['Sortino_1Y'] = np.nan
        df['DSR_1Y'] = np.nan
        df['Perf_3Y'] = np.nan
        df['Vol_3Y'] = np.nan
        df['Perf_1Y'] = np.nan
        df['Vol_1Y'] = np.nan
        df['Sharpe_3Y'] = np.nan
        df['Sharpe_1Y'] = np.nan

        if self.intermediario == 'BPPB':
            IR_TEV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM', 'OBB_GLOB_HY']
            SOR_DSR = ['FLEX']
            SHA_VOL = ['OPP']
            PER_VOL = ['LIQ']
        elif self.intermediario == 'BPL':
            IR_TEV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_EUR', 'OBB_CORP', 'OBB_GLOB', 'OBB_USA', 'OBB_EM', 'OBB_GLOB_HY']
            SOR_DSR = ['BIL', 'FLEX']
            SHA_VOL = ['OPP']
            PER_VOL = ['LIQ', 'LIQ_FOR']
        elif self.intermediario == 'CRV':
            IR_TEV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM', 'OBB_GLOB_HY']
            SOR_DSR = ['FLEX']
            SHA_VOL = ['OPP']
            PER_VOL = ['LIQ']

        # Aggiunge gli indici delle liste al file completo 
        for filename in os.listdir(self.directory_output_liste):
            if filename[:-9] in IR_TEV:
                i = 10
                while True:
                    try:
                        lista_indici = pd.read_csv(self.directory_output_liste.joinpath(filename), sep = ';', header=0, skiprows=2, skipfooter=i, engine='python', encoding='unicode_escape')
                    except pd.errors.ParserError:
                        # se i fondi sono pochi la tabella delle correlazioni viene riempita al completo e le righe da saltare in fondo sono più di 14
                        i = i + 1
                    else:
                        if lista_indici.Nome.isna().any():
                            i = i + 1
                            continue
                        else:
                            break
                lista_indici = lista_indici.replace(',', '.', regex=True).astype({'Information ratio': float, 'TEV' : float})
                if filename[-6:-4] == '3Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Information_Ratio_3Y'] = df['Information ratio'].fillna(df['Information_Ratio_3Y'])
                    df['TEV_3Y'] = df['TEV'].fillna(df['TEV_3Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Information ratio', 'TEV'], axis=1, inplace=True)
                    df = df.astype({'Information_Ratio_3Y' : float, 'TEV_3Y' : float})
                elif filename[-6:-4] == '1Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Information_Ratio_1Y'] = df['Information ratio'].fillna(df['Information_Ratio_1Y'])
                    df['TEV_1Y'] = df['TEV'].fillna(df['TEV_1Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Information ratio', 'TEV'], axis=1, inplace=True)
                    df = df.astype({'Information_Ratio_1Y' : float, 'TEV_1Y' : float})
            elif filename[:-9] in SOR_DSR:
                i = 10
                while True:
                    try:
                        lista_indici = pd.read_csv(self.directory_output_liste.joinpath(filename), sep = ';', header=0, skiprows=2, skipfooter=i, engine='python', encoding='unicode_escape')
                    except pd.errors.ParserError:
                        # se i fondi sono pochi la tabella delle correlazioni viene riempita al completo e le righe da saltare in fondo sono più di 14
                        i = i + 1
                    else:
                        if lista_indici.Nome.isna().any():
                            i = i + 1
                            continue
                        else:
                            break
                lista_indici = lista_indici.replace(',', '.', regex=True).astype({'Sortino ratio': float, 'DSR' : float})
                if filename[-6:-4] == '3Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Sortino_3Y'] = df['Sortino ratio'].fillna(df['Sortino_3Y'])
                    df['DSR_3Y'] = df['DSR'].fillna(df['DSR_3Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Sortino ratio', 'DSR'], axis=1, inplace=True)
                elif filename[-6:-4] == '1Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Sortino_1Y'] = df['Sortino ratio'].fillna(df['Sortino_1Y'])
                    df['DSR_1Y'] = df['DSR'].fillna(df['DSR_1Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Sortino ratio', 'DSR'], axis=1, inplace=True)
            elif filename[:-9] in SHA_VOL:
                i = 10
                while True:
                    try:
                        lista_indici = pd.read_csv(self.directory_output_liste.joinpath(filename), sep = ';', header=0, skiprows=2, skipfooter=i, engine='python', encoding='unicode_escape')
                    except pd.errors.ParserError:
                        # se i fondi sono pochi la tabella delle correlazioni viene riempita al completo e le righe da saltare in fondo sono più di 14
                        i = i + 1
                    else:
                        if lista_indici.Nome.isna().any():
                            i = i + 1
                            continue
                        else:
                            break
                lista_indici = lista_indici.replace(',', '.', regex=True).astype({'Sharpe ratio': float, 'Volatilità' : float})
                #print(filename)
                if filename[-6:-4] == '3Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Sharpe_3Y'] = df['Sharpe ratio'].fillna(df['Sharpe_3Y'])
                    df['Vol_3Y'] = df['Volatilità'].fillna(df['Vol_3Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Sharpe ratio', 'Volatilità'], axis=1, inplace=True)
                elif filename[-6:-4] == '1Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Sharpe_1Y'] = df['Sharpe ratio'].fillna(df['Sharpe_1Y'])
                    df['Vol_1Y'] = df['Volatilità'].fillna(df['Vol_1Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Sharpe ratio', 'Volatilità'], axis=1, inplace=True)
            elif filename[:-9] in PER_VOL:
                i = 10
                while True:
                    try:
                        lista_indici = pd.read_csv(self.directory_output_liste.joinpath(filename), sep = ';', header=0, skiprows=2, skipfooter=i, engine='python', encoding='unicode_escape')
                    except pd.errors.ParserError:
                        # se i fondi sono pochi la tabella delle correlazioni viene riempita al completo e le righe da saltare in fondo sono più di 14
                        i = i + 1
                    else:
                        if lista_indici.Nome.isna().any():
                            i = i + 1
                            continue
                        else:
                            break
                lista_indici = lista_indici.replace(',', '.', regex=True).astype({'Perf Ann.': float, 'Volatilità' : float})
                #print(filename)
                if filename[-6:-4] == '3Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Perf_3Y'] = df['Perf Ann.'].fillna(df['Perf_3Y'])
                    df['Vol_3Y'] = df['Volatilità'].fillna(df['Vol_3Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Perf Ann.', 'Volatilità'], axis=1, inplace=True)
                elif filename[-6:-4] == '1Y':
                    df = df.merge(lista_indici, how='outer', left_on='ISIN', right_on='Codice ISIN')
                    df['Perf_1Y'] = df['Perf Ann.'].fillna(df['Perf_1Y'])
                    df['Vol_1Y'] = df['Volatilità'].fillna(df['Vol_1Y'])
                    df.drop(['Codice ISIN', 'Nome', 'Valuta', 'Perf Ann.', 'Volatilità'], axis=1, inplace=True)
        df.to_excel(self.file_ranking, index=False)

    def discriminazione_flessibili_e_bilanciati(self):
        """
        Discrimina i flessibili  e i bilanciati in classi diverse a seconda della loro volatilità (maggiore o minore di 0.05), o alla loro appartenenza alle classi.
        """
        df = pd.read_excel(self.file_ranking, index_col=None)
        if self.intermediario == 'BPPB':
            print("\nsto discriminando i flessibili in base alla loro volatilità...")
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df['categoria_flessibili'].map({'bassa_vola' : 'FLEX_BVOL', 'media_alta_vola' : 'FLEX_MAVOL'}, na_action='ignore')
        elif self.intermediario == 'BPL':
            print("\nsto discriminando i flessibili e i bilanciati in base alla loro classe di appartenenza...")
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df['micro_categoria'].map({'Flessibili prudenti globale' : 'FLEX_PR', 'Flessibili prudenti Europa' : 'FLEX_PR', 'Flessibili Europa' : 'FLEX_DIN', 'Flessibili Dollaro US' : 'FLEX_DIN', 'Fless. Global Euro' : 'FLEX_DIN', 'Fless. Global' : 'FLEX_DIN',}, na_action='ignore')
            df.loc[df['macro_categoria'] == 'BIL', 'macro_categoria'] = df['micro_categoria'].map({'Bilanc. Prud. Europa' : 'BIL_MBVOL', 'Bilanc. Prud. Dollaro US' : 'BIL_MBVOL', 'Bilanc. Prud. Global Euro' : 'BIL_MBVOL', 'Bilanc. Prud. Global' : 'BIL_MBVOL', 'Bilanc. Prud. altre valute' : 'BIL_MBVOL',  'Bilanc. Equilib. Europa' : 'BIL_MBVOL', 'Bilanc. Equil. Dollaro US' : 'BIL_MBVOL', 'Bilanc. Equil. Global Euro' : 'BIL_MBVOL', 'Bilanc. Equil. Global' : 'BIL_MBVOL', 'Bilanc. Equil. altre valute' : 'BIL_MBVOL', 'Bilanc. Aggress. Europa' : 'BIL_AVOL', 'Bilanc. aggress. Dollaro US' : 'BIL_AVOL', 'Bilanc. Aggress. Global Euro' : 'BIL_AVOL', 'Bilanc. Aggress. Global' : 'BIL_AVOL', 'Bilanc. Aggress. altre valute' : 'BIL_AVOL'}, na_action='ignore')
        elif self.intermediario == 'CRV':
            print("\nsto discriminando i flessibili in base alla loro volatilità...")
            df.loc[df['macro_categoria'] == 'FLEX', 'macro_categoria'] = df['categoria_flessibili'].map({'bassa_vola' : 'FLEX_PR', 'media_alta_vola' : 'FLEX_DIN'}, na_action='ignore')
        df.to_excel(self.file_ranking, index=False)

    def rank(self, metodo=''):
        # Aggiungi il criterio best/worst ad un anno. Discrimina per attività del fondo, metti gli ESG a fianco, rimuovi colonne inutili.

        # TODO : fattorizza per tipo di indicatore
        # TODO : fallo all'interno di un wrapper come il metodo aggiunta_colonne
        """
        Crea il file di ranking con tanti fogli quante sono le macro asset class.
        """
        t0_3Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+3)).strftime('%Y/%m/%d') # data iniziale tre anni fa
        t0_1Y = (datetime.datetime.strptime(self.t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(years=+1)).strftime("%Y/%m/%d") # data iniziale un anno fa
        # Creazione file ranking diviso per macro
        print('\nsto facendo la rankizzazione')
        df = pd.read_excel(self.file_ranking, index_col=None)
        df['data_di_avvio'] = pd.to_datetime(df['data_di_avvio'], dayfirst=True)
        writer = pd.ExcelWriter(self.file_ranking,  engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
        
        if self.intermediario == 'BPPB':
            anni_detenzione = 3
            IR_TEV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM']
            SOR_DSR = ['FLEX_BVOL', 'FLEX_MAVOL'] # Ora i flessibili sono discriminati
            SHA_VOL = ['OPP']
            PER_VOL = ['LIQ']
            if metodo == 'singolo':
                micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 
                    'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_CORP' : 'Obblig. Euro corporate', 
                    'OBB_GLOB' : 'Obblig. globale', 'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield'}
            elif metodo == 'doppio':
                micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 
                    'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_CORP' : 'Obblig. Euro corporate', 
                    'OBB_GLOB' : 'Obblig. globale', 'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield', 'LIQ' : 'Monetari Euro'}
            soluzioni = self.soluzioni_BPPB
        elif self.intermediario =='BPL':
            anni_detenzione = 5
            IR_TEV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_EUR', 'OBB_CORP', 'OBB_GLOB', 'OBB_USA', 'OBB_EM', 'OBB_GLOB_HY']
            SOR_DSR = ['BIL_MBVOL', 'BIL_AVOL', 'FLEX_PR', 'FLEX_DIN'] # Ora i flessibili e i bilanciati sono discriminati
            SHA_VOL = ['OPP']
            PER_VOL = ['LIQ', 'LIQ_FOR']
            if metodo == 'singolo':
                micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 
                    'AZ_GLOB' : 'Az. globale', 'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_EUR' : 'Obblig. Europa', 
                    'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale', 'OBB_USA' : 'Obblig. Dollaro US all mat', 
                    'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield'}
            elif metodo == 'doppio':
                micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 
                    'AZ_GLOB' : 'Az. globale', 'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_EUR' : 'Obblig. Europa', 
                    'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale', 'OBB_USA' : 'Obblig. Dollaro US all mat', 
                    'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield', 'LIQ' : 'Monetari Euro'}
            soluzioni = self.soluzioni_BPL
        elif self.intermediario =='CRV':
            anni_detenzione = 3
            IR_TEV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM', 'OBB_GLOB_HY']
            SOR_DSR = ['FLEX_PR', 'FLEX_DIN'] # Ora i flessibili sono discriminati
            SHA_VOL = ['OPP']
            PER_VOL = ['LIQ']
            micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 
                'AZ_GLOB' : 'Az. globale', 'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 
                'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale', 'OBB_EM' : 'Obblig. Paesi Emerg.', 
                'OBB_GLOB_HY' : 'Obblig. globale high yield'}

        
        for macro in df.loc[:, 'macro_categoria'].unique():
            # Crea un foglio per ogni macro categoria
            foglio = df.loc[df['macro_categoria'] == macro].copy()
            if macro in IR_TEV:
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL': # metodo best-worst
                    if metodo == 'singolo':
                        # Rank IR_1Y
                        foglio['ranking_IR_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Information_Ratio_1Y'].notnull()), 'Information_Ratio_1Y'].rank(method='first', na_option='bottom', ascending=False)
                        # Quartile IR_1Y
                        foglio['quartile_IR_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Information_Ratio_1Y'].notnull()), 'Information_Ratio_1Y'].apply(lambda x: 'best' if x > foglio['Information_Ratio_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile IR_1Y
                        foglio['terzile_IR_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Information_Ratio_1Y'].notnull()), 'Information_Ratio_1Y'].apply(lambda x: 'best' if x > foglio['Information_Ratio_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                        # Creazione IR_corretto_1Y
                        foglio['IR_corretto_1Y'] = ((df['Information_Ratio_1Y'] * (df['TEV_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['TEV_1Y'] / 100)
                        # Rank IR_corretto_1Y
                        foglio['ranking_IR_1Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False)
                        # Quartile IR_1Y corretto
                        foglio['quartile_IR_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].apply(lambda x: 'best' if x > foglio['IR_corretto_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile IR_1Y corretto
                        foglio['terzile_IR_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].apply(lambda x: 'best' if x > foglio['IR_corretto_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')

                        # Rank IR_3Y
                        foglio['ranking_IR_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'].notnull()) & (foglio['Information_Ratio_3Y'].notnull()), 'Information_Ratio_3Y'].rank(method='first', na_option='keep', ascending=False)
                        # Aggiunta nota per i fondi che possiedono dati a tre anni pur non avendo tre anni di vita, e ad un anno non avendo un anno di vita
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & foglio['Best_Worst'].isnull(), 'note'] = 'Ha 3 anni, ma non è in classifica.'
                        foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Information_Ratio_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                        # Quartile IR_3Y
                        foglio['quartile_IR_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'].notnull()) & (foglio['Information_Ratio_3Y'].notnull()), 'Information_Ratio_3Y'].apply(lambda x: 'best' if x > foglio['Information_Ratio_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile IR_3Y
                        foglio['terzile_IR_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'].notnull()) & (foglio['Information_Ratio_3Y'].notnull()), 'Information_Ratio_3Y'].apply(lambda x: 'best' if x > foglio['Information_Ratio_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                        # Creazione IR_corretto_3Y
                        foglio['IR_corretto_3Y'] = ((df['Information_Ratio_3Y'] * (df['TEV_3Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['TEV_3Y'] / 100)
                        # Rank IR_corretto_3Y
                        foglio['ranking_IR_3Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'].notnull()) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False)
                        # Quartile IR_3Y corretto
                        foglio['quartile_IR_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'].notnull()) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].apply(lambda x: 'best' if x > foglio['IR_corretto_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile IR_3Y corretto
                        foglio['terzile_IR_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'].notnull()) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].apply(lambda x: 'best' if x > foglio['IR_corretto_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                        
                        # Ranking finale
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'best') & (foglio['IR_corretto_3Y'].notnull()) & (foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]), 'ranking_finale'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'best') & (foglio['IR_corretto_3Y'].notnull()) & (foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False)
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'best') & (foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]), 'ranking_finale'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'best') & (foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'worst') & (foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]), 'ranking_finale'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'worst') & (foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'worst') & (foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]), 'ranking_finale'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Best_Worst'] == 'worst') & (foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                
                    elif metodo == 'doppio':
                        # Aggiunta nota per i fondi che possiedono dati a tre anni pur non avendo tre anni di vita, e ad un anno non avendo un anno di vita
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & foglio['Best_Worst_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non è in classifica.'
                        foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Information_Ratio_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                        # Creazione IR_corretto_3Y
                        foglio['IR_corretto_3Y'] = ((df['Information_Ratio_3Y'] * (df['TEV_3Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['TEV_3Y'] / 100)
                        foglio['IR_corretto_1Y'] = ((df['Information_Ratio_1Y'] * (df['TEV_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['TEV_1Y'] / 100)
                        # Ranking finale
                        if soluzioni[macro] == 1:
                            # Fondi best blend - Gerarchia : semi_attivo, attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi best non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi worst non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                        elif soluzioni[macro] == 2:
                            # Fondi best blend - Gerarchia : semi_attivo, attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi best non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi worst non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                        elif soluzioni[macro] == 3:
                            # Fondi best blend - Gerarchia : (semi_attivo & attivo), molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi best non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            # Fondi worst blend - Gerarchia : (semi_attivo & attivo), molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi worst non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()

                elif self.intermediario == 'CRV': # metodo normalizzazione
                    # Creazione IR_corretto_1Y
                    foglio['IR_corretto_1Y'] = ((df['Information_Ratio_1Y'] * (df['TEV_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['TEV_1Y'] / 100)
                    # Creazione IR_corretto_3Y
                    foglio['IR_corretto_3Y'] = ((df['Information_Ratio_3Y'] * (df['TEV_3Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['TEV_3Y'] / 100)
                    # Note
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Information_Ratio_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
                    foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & (foglio['Information_Ratio_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                    # Ranking finale
                    minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'])
                    massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale_3Y'] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'] / (massimo_3Y - minimo_3Y)
                    minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'])
                    massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale_1Y'] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'] / (massimo_1Y - minimo_1Y)
                    foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
                    foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')

                # Seleziona colonne utili
                if self.intermediario == 'BPPB':
                    if metodo == 'singolo':
                        foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'Best_Worst', 'micro_categoria', 'ranking_finale', 'Information_Ratio_3Y', 'ranking_IR_3Y', 'quartile_IR_3Y',
                            'terzile_IR_3Y', 'TEV_3Y', 'commissione', 'IR_corretto_3Y', 'ranking_IR_3Y_corretto', 'quartile_IR_corretto_3Y', 'terzile_IR_corretto_3Y', 
                            'Information_Ratio_1Y', 'ranking_IR_1Y', 'quartile_IR_1Y', 'terzile_IR_1Y', 'TEV_1Y', 'commissione', 'IR_corretto_1Y', 'ranking_IR_1Y_corretto',
                            'quartile_IR_corretto_1Y', 'terzile_IR_corretto_1Y', 'SFDR', 'fondo_a_finestra', 'note']]
                    elif metodo == 'doppio':
                        pass
                elif self.intermediario == 'BPL':
                    if metodo == 'singolo':
                        foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'Best_Worst', 'micro_categoria', 'ranking_finale', 'Information_Ratio_3Y', 'ranking_IR_3Y', 'quartile_IR_3Y',
                            'terzile_IR_3Y', 'TEV_3Y', 'commissione', 'IR_corretto_3Y', 'ranking_IR_3Y_corretto', 'quartile_IR_corretto_3Y', 'terzile_IR_corretto_3Y', 
                            'Information_Ratio_1Y', 'ranking_IR_1Y', 'quartile_IR_1Y', 'terzile_IR_1Y', 'TEV_1Y', 'commissione', 'IR_corretto_1Y', 'ranking_IR_1Y_corretto',
                            'quartile_IR_corretto_1Y', 'terzile_IR_corretto_1Y', 'note']]
                    elif metodo == 'doppio':
                        foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Best_Worst_3Y', 'grado_gestione_3Y', 
                            'Best_Worst_1Y', 'grado_gestione_1Y', 'ranking_per_grado_3Y', 'ranking_per_grado_1Y', 'ranking_finale', 'Information_Ratio_3Y',
                            'TEV_3Y', 'commissione', 'IR_corretto_3Y', 'Information_Ratio_1Y', 'TEV_1Y', 'commissione', 'IR_corretto_1Y', 'note']]
                elif self.intermediario == 'CRV':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'podio', 'ranking_finale', 'ranking_finale_3Y',
                        'ranking_finale_1Y', 'Information_Ratio_3Y', 'TEV_3Y', 'commissione', 'IR_corretto_3Y', 'Information_Ratio_1Y', 
                        'TEV_1Y', 'commissione', 'IR_corretto_1Y', 'note']]
                
                # Cambio formato data
                foglio['data_di_avvio'] = foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
                # Ordinamento finale
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                    foglio.sort_values('ranking_finale', ascending=True, inplace=True)
                elif self.intermediario == 'CRV':
                    foglio.sort_values('ranking_finale', ascending=False, inplace=True)
                    # Etichetta ND per i fondi senza dati
                    foglio['ranking_finale_1Y'] = foglio['ranking_finale_1Y'].fillna('ND')
                    foglio['ranking_finale_3Y'] = foglio['ranking_finale_3Y'].fillna('ND')
                    foglio['ranking_finale'] = foglio['ranking_finale'].fillna('ND')
                # Reindex
                foglio.reset_index(drop=True, inplace=True)
                
                # Crea foglio
                foglio.to_excel(writer, sheet_name=macro)

            elif macro in SOR_DSR:
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                    # Rank SO_1Y
                    foglio['ranking_SO_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Sortino_1Y'].notnull()), 'Sortino_1Y'].rank(method='first', na_option='bottom', ascending=False)
                    # Quartile SO_1Y
                    foglio['quartile_SO_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Sortino_1Y'].notnull()), 'Sortino_1Y'].apply(lambda x: 'best' if x > foglio['Sortino_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SO_1Y
                    foglio['terzile_SO_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Sortino_1Y'].notnull()), 'Sortino_1Y'].apply(lambda x: 'best' if x > foglio['Sortino_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                    # Creazione SO_corretto_1Y
                    foglio['SO_corretto_1Y'] = ((df['Sortino_1Y'] * (df['DSR_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['DSR_1Y'] / 100)
                    # Rank SO_corretto_1Y
                    foglio['ranking_SO_1Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False)
                    # Quartile SO_1Y corretto
                    foglio['quartile_SO_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'].apply(lambda x: 'best' if x > foglio['SO_corretto_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SO_1Y corretto
                    foglio['terzile_SO_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'].apply(lambda x: 'best' if x > foglio['SO_corretto_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')

                    # Rank SO_3Y
                    foglio['ranking_SO_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'Sortino_3Y'].rank(method='first', na_option='keep', ascending=False)
                    # Aggiunta nota per i fondi che possiedono dati a tre anni pur non avendo tre anni di vita, e ad un anno non avendo un anno di vita
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & foglio['Sortino_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
                    foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Sortino_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                    # Quartile SO_3Y TOGLI IL BEST_WORST
                    foglio['quartile_SO_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'Sortino_3Y'].apply(lambda x: 'best' if x > foglio['Sortino_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SO_3Y
                    foglio['terzile_SO_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'Sortino_3Y'].apply(lambda x: 'best' if x > foglio['Sortino_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                    # Creazione SO_corretto_3Y
                    foglio['SO_corretto_3Y'] = ((df['Sortino_3Y'] * (df['DSR_3Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['DSR_3Y'] / 100)
                    # Rank SO_corretto_3Y
                    foglio['ranking_SO_3Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False)
                    # Quartile SO_3Y corretto
                    foglio['quartile_SO_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'].apply(lambda x: 'best' if x > foglio['SO_corretto_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SO_3Y corretto
                    foglio['terzile_SO_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'].apply(lambda x: 'best' if x > foglio['SO_corretto_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                
                elif self.intermediario == 'CRV': # metodo normalizzazione
                    # Creazione SO_corretto_1Y
                    foglio['SO_corretto_1Y'] = ((df['Sortino_1Y'] * (df['DSR_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['DSR_1Y'] / 100)
                    # Creazione SO_corretto_3Y
                    foglio['SO_corretto_3Y'] = ((df['Sortino_3Y'] * (df['DSR_3Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['DSR_3Y'] / 100)
                    # Note
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sortino_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
                    foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                    # Ranking finale
                    minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'])
                    massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'ranking_finale_3Y'] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'] / (massimo_3Y - minimo_3Y)
                    minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'])
                    massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'ranking_finale_1Y'] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'] / (massimo_1Y - minimo_1Y)
                    foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
                    foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')

                # Seleziona colonne utili
                if self.intermediario == 'BPPB':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sortino_3Y', 'ranking_SO_3Y', 'quartile_SO_3Y',
                        'terzile_SO_3Y', 'DSR_3Y', 'commissione', 'SO_corretto_3Y', 'ranking_SO_3Y_corretto', 'quartile_SO_corretto_3Y', 'terzile_SO_corretto_3Y', 
                        'Sortino_1Y', 'ranking_SO_1Y', 'quartile_SO_1Y', 'terzile_SO_1Y', 'DSR_1Y', 'commissione', 'SO_corretto_1Y', 'ranking_SO_1Y_corretto',
                        'quartile_SO_corretto_1Y', 'terzile_SO_corretto_1Y', 'SFDR', 'fondo_a_finestra', 'note']]
                elif self.intermediario == 'BPL':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sortino_3Y', 'ranking_SO_3Y', 'quartile_SO_3Y',
                        'terzile_SO_3Y', 'DSR_3Y', 'commissione', 'SO_corretto_3Y', 'ranking_SO_3Y_corretto', 'quartile_SO_corretto_3Y', 'terzile_SO_corretto_3Y', 
                        'Sortino_1Y', 'ranking_SO_1Y', 'quartile_SO_1Y', 'terzile_SO_1Y', 'DSR_1Y', 'commissione', 'SO_corretto_1Y', 'ranking_SO_1Y_corretto',
                        'quartile_SO_corretto_1Y', 'terzile_SO_corretto_1Y', 'note']]
                elif self.intermediario == 'CRV':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'podio', 'ranking_finale', 'ranking_finale_3Y', 'ranking_finale_1Y', 'Sortino_3Y',
                        'DSR_3Y', 'commissione', 'SO_corretto_3Y', 'Sortino_1Y', 'DSR_1Y', 'commissione', 'SO_corretto_1Y', 'note']]
                
                # Cambio formato data
                foglio['data_di_avvio'] = foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
                # Ordinamento finale
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                    foglio.sort_values('ranking_SO_3Y_corretto', ascending=True, inplace=True)
                elif self.intermediario == 'CRV':
                    foglio.sort_values('ranking_finale', ascending=False, inplace=True)
                    # Etichetta ND per i fondi senza dati
                    foglio['ranking_finale_1Y'] = foglio['ranking_finale_1Y'].fillna('ND')
                    foglio['ranking_finale_3Y'] = foglio['ranking_finale_3Y'].fillna('ND')
                    foglio['ranking_finale'] = foglio['ranking_finale'].fillna('ND')
                # Reindex
                foglio.reset_index(drop=True, inplace=True)

                # Crea foglio
                foglio.to_excel(writer, sheet_name=macro)
            
            elif macro in SHA_VOL:
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                    # Rank SH_1Y
                    foglio['ranking_SH_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Sharpe_1Y'].notnull()), 'Sharpe_1Y'].rank(method='first', na_option='bottom', ascending=False)
                    # Quartile SH_1Y
                    foglio['quartile_SH_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Sharpe_1Y'].notnull()), 'Sharpe_1Y'].apply(lambda x: 'best' if x > foglio['Sharpe_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SH_1Y
                    foglio['terzile_SH_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Sharpe_1Y'].notnull()), 'Sharpe_1Y'].apply(lambda x: 'best' if x > foglio['Sharpe_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                    # Creazione SH_corretto_1Y
                    foglio['SH_corretto_1Y'] = ((df['Sharpe_1Y'] * (df['Vol_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['Vol_1Y'] / 100)
                    # Rank SH_corretto_1Y
                    foglio['ranking_SH_1Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False)
                    # Quartile SH_1Y corretto
                    foglio['quartile_SH_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'].apply(lambda x: 'best' if x > foglio['SH_corretto_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SH_1Y corretto
                    foglio['terzile_SH_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'].apply(lambda x: 'best' if x > foglio['SH_corretto_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')

                    # Rank SH_3Y
                    foglio['ranking_SH_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'Sharpe_3Y'].rank(method='first', na_option='keep', ascending=False)
                    # Aggiunta nota per i fondi che possiedono dati a tre anni pur non avendo tre anni di vita, e ad un anno non avendo un anno di vita
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & foglio['Sharpe_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
                    foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Sharpe_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                    # Quartile SH_3Y TOGLI IL BEST_WORST
                    foglio['quartile_SH_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'Sharpe_3Y'].apply(lambda x: 'best' if x > foglio['Sharpe_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SH_3Y
                    foglio['terzile_SH_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'Sharpe_3Y'].apply(lambda x: 'best' if x > foglio['Sharpe_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                    # Creazione SH_corretto_3Y
                    foglio['SH_corretto_3Y'] = ((df['Sharpe_3Y'] * (df['Vol_3Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['Vol_3Y'] / 100)
                    # Rank SH_corretto_3Y
                    foglio['ranking_SH_3Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False)
                    # Quartile SH_3Y corretto
                    foglio['quartile_SH_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'].apply(lambda x: 'best' if x > foglio['SH_corretto_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                    # Terzile SH_3Y corretto
                    foglio['terzile_SH_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'].apply(lambda x: 'best' if x > foglio['SH_corretto_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')

                elif self.intermediario == 'CRV': # metodo normalizzazione
                    # Creazione SH_corretto_1Y
                    foglio['SH_corretto_1Y'] = ((df['Sharpe_1Y'] * (df['Vol_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['Vol_1Y'] / 100)
                    # Creazione SH_corretto_3Y
                    foglio['SH_corretto_3Y'] = ((df['Sharpe_3Y'] * (df['Vol_3Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['Vol_3Y'] / 100)
                    # Note
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Sharpe_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
                    foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                    # Ranking finale
                    minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'])
                    massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'ranking_finale_3Y'] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'] / (massimo_3Y - minimo_3Y)
                    minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'])
                    massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'ranking_finale_1Y'] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'] / (massimo_1Y - minimo_1Y)
                    foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
                    foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')
                
                # Seleziona colonne utili
                if self.intermediario == 'BPPB':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sharpe_3Y', 'ranking_SH_3Y', 'quartile_SH_3Y',
                        'terzile_SH_3Y', 'Vol_3Y', 'commissione', 'SH_corretto_3Y', 'ranking_SH_3Y_corretto', 'quartile_SH_corretto_3Y', 'terzile_SH_corretto_3Y', 
                        'Sharpe_1Y', 'ranking_SH_1Y', 'quartile_SH_1Y', 'terzile_SH_1Y', 'Vol_1Y', 'commissione', 'SH_corretto_1Y', 'ranking_SH_1Y_corretto',
                        'quartile_SH_corretto_1Y', 'terzile_SH_corretto_1Y', 'SFDR', 'fondo_a_finestra', 'note']]
                elif self.intermediario == 'BPL':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sharpe_3Y', 'ranking_SH_3Y', 'quartile_SH_3Y',
                        'terzile_SH_3Y', 'Vol_3Y', 'commissione', 'SH_corretto_3Y', 'ranking_SH_3Y_corretto', 'quartile_SH_corretto_3Y', 'terzile_SH_corretto_3Y', 
                        'Sharpe_1Y', 'ranking_SH_1Y', 'quartile_SH_1Y', 'terzile_SH_1Y', 'Vol_1Y', 'commissione', 'SH_corretto_1Y', 'ranking_SH_1Y_corretto',
                        'quartile_SH_corretto_1Y', 'terzile_SH_corretto_1Y', 'note']]
                elif self.intermediario == 'CRV':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'podio', 'ranking_finale', 'ranking_finale_3Y', 'ranking_finale_1Y', 'Sharpe_3Y',
                        'Vol_3Y', 'commissione', 'SH_corretto_3Y', 'Sharpe_1Y', 'Vol_1Y', 'commissione', 'SH_corretto_1Y', 'note']]
                
                # Cambio formato data
                foglio['data_di_avvio'] = foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
                # Ordinamento finale
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                    foglio.sort_values('ranking_SH_3Y_corretto', ascending=True, inplace=True)
                elif self.intermediario == 'CRV':
                    foglio.sort_values('ranking_finale', ascending=False, inplace=True)
                    # Etichetta ND per i fondi senza dati
                    foglio['ranking_finale_1Y'] = foglio['ranking_finale_1Y'].fillna('ND')
                    foglio['ranking_finale_3Y'] = foglio['ranking_finale_3Y'].fillna('ND')
                    foglio['ranking_finale'] = foglio['ranking_finale'].fillna('ND')
                # Reindex
                foglio.reset_index(drop=True, inplace=True)

                # Crea foglio
                foglio.to_excel(writer, sheet_name=macro)

            elif macro in PER_VOL:
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                    if metodo == 'singolo' or (metodo == 'doppio' and macro == 'LIQ_FOR'):
                        # Rank PERF_1Y
                        foglio['ranking_PERF_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Perf_1Y'].notnull()), 'Perf_1Y'].rank(method='first', na_option='bottom', ascending=False)
                        # Quartile PERF_1Y
                        foglio['quartile_PERF_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Perf_1Y'].notnull()), 'Perf_1Y'].apply(lambda x: 'best' if x > foglio['Perf_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile PERF_1Y
                        foglio['terzile_PERF_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['Perf_1Y'].notnull()), 'Perf_1Y'].apply(lambda x: 'best' if x > foglio['Perf_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                        # Creazione PERF_corretto_1Y
                        foglio['PERF_corretto_1Y'] = (df['Perf_1Y'] / 100) - (df['commissione'] / anni_detenzione)
                        # Rank PERF_corretto_1Y
                        foglio['ranking_PERF_1Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False)
                        # Quartile PERF_1Y corretto
                        foglio['quartile_PERF_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].apply(lambda x: 'best' if x > foglio['PERF_corretto_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile PERF_1Y corretto
                        foglio['terzile_PERF_corretto_1Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].apply(lambda x: 'best' if x > foglio['PERF_corretto_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')

                        # Rank PERF_3Y
                        foglio['ranking_PERF_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Perf_3Y'].notnull()), 'Perf_3Y'].rank(method='first', na_option='keep', ascending=False)
                        # Aggiunta nota per i fondi che possiedono dati a tre anni pur non avendo tre anni di vita, e ad un anno non avendo un anno di vita
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & foglio['Perf_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
                        foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Perf_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                        # Quartile PERF_3Y TOGLI IL BEST_WORST
                        foglio['quartile_PERF_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Perf_3Y'].notnull()), 'Perf_3Y'].apply(lambda x: 'best' if x > foglio['Perf_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile PERF_3Y
                        foglio['terzile_PERF_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Perf_3Y'].notnull()), 'Perf_3Y'].apply(lambda x: 'best' if x > foglio['Perf_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
                        # Creazione PERF_corretto_3Y (la volatilità è già in percentuale)
                        foglio['PERF_corretto_3Y'] = (df['Perf_3Y'] / 100) - (df['commissione'] / anni_detenzione)
                        # Rank PERF_corretto_3Y
                        foglio['ranking_PERF_3Y_corretto'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False)
                        # Quartile PERF_3Y corretto
                        foglio['quartile_PERF_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].apply(lambda x: 'best' if x > foglio['PERF_corretto_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
                        # Terzile PERF_3Y corretto
                        foglio['terzile_PERF_corretto_3Y'] = foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].apply(lambda x: 'best' if x > foglio['PERF_corretto_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')

                    elif metodo == 'doppio':
                        # Aggiunta nota per i fondi che possiedono dati a tre anni pur non avendo tre anni di vita, e ad un anno non avendo un anno di vita
                        foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & foglio['Best_Worst_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non è in classifica.'
                        foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Information_Ratio_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                        # Creazione PERF_corretto_3Y
                        foglio['PERF_corretto_3Y'] = (df['Perf_3Y'] / 100) - (df['commissione'] / anni_detenzione)
                        foglio['PERF_corretto_1Y'] = (df['Perf_1Y'] / 100) - (df['commissione'] / anni_detenzione)
                        # Ranking finale
                        if soluzioni[macro] == 1:
                            # Fondi best blend - Gerarchia : semi_attivo, attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi best non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi worst non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                        elif soluzioni[macro] == 2:
                            # Fondi best blend - Gerarchia : semi_attivo, attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi best non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi worst non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                        elif soluzioni[macro] == 3:
                            # Fondi best blend - Gerarchia : (semi_attivo & attivo), molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi best non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            # Fondi worst blend - Gerarchia : (semi_attivo & attivo), molto_attivo
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == micro_blend_classi_a_benchmark[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + foglio['ranking_finale'].max()
                            # Fondi worst non blend
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()
                            foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != micro_blend_classi_a_benchmark[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + foglio['ranking_finale'].max()

                elif self.intermediario == 'CRV': # metodo normalizzazione
                    # Creazione PERF_corretto_1Y
                    foglio['PERF_corretto_1Y'] = (df['Perf_1Y'] / 100) - (df['commissione'] / anni_detenzione)
                    # Creazione PERF_corretto_3Y
                    foglio['PERF_corretto_3Y'] = (df['Perf_3Y'] / 100) - (df['commissione'] / anni_detenzione)
                    # Note
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['Perf_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
                    foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & (foglio['Perf_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
                    # Ranking finale
                    minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'])
                    massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale_3Y'] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'] / (massimo_3Y - minimo_3Y)
                    minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'])
                    massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'])
                    foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale_1Y'] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'] / (massimo_1Y - minimo_1Y)
                    foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
                    foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')
                
                # Seleziona colonne utili
                if self.intermediario == 'BPPB':
                    if metodo == 'singolo':
                        foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Perf_3Y', 'ranking_PERF_3Y', 'quartile_PERF_3Y',
                            'terzile_PERF_3Y', 'Vol_3Y', 'commissione', 'PERF_corretto_3Y', 'ranking_PERF_3Y_corretto', 'quartile_PERF_corretto_3Y', 'terzile_PERF_corretto_3Y', 
                            'Perf_1Y', 'ranking_PERF_1Y', 'quartile_PERF_1Y', 'terzile_PERF_1Y', 'Vol_1Y', 'commissione', 'PERF_corretto_1Y', 'ranking_PERF_1Y_corretto',
                            'quartile_PERF_corretto_1Y', 'terzile_PERF_corretto_1Y', 'SFDR', 'fondo_a_finestra', 'note']]
                    elif metodo == 'doppio':
                        pass
                elif self.intermediario == 'BPL':
                    if metodo == 'singolo' or (metodo == 'doppio' and macro == 'LIQ_FOR'):
                        foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Perf_3Y', 'ranking_PERF_3Y', 'quartile_PERF_3Y',
                            'terzile_PERF_3Y', 'Vol_3Y', 'commissione', 'PERF_corretto_3Y', 'ranking_PERF_3Y_corretto', 'quartile_PERF_corretto_3Y', 'terzile_PERF_corretto_3Y', 
                            'Perf_1Y', 'ranking_PERF_1Y', 'quartile_PERF_1Y', 'terzile_PERF_1Y', 'Vol_1Y', 'commissione', 'PERF_corretto_1Y', 'ranking_PERF_1Y_corretto',
                            'quartile_PERF_corretto_1Y', 'terzile_PERF_corretto_1Y', 'note']]
                    elif metodo == 'doppio':
                        foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Best_Worst_3Y', 'grado_gestione_3Y', 'Best_Worst_1Y', 
                            'grado_gestione_1Y', 'ranking_per_grado_3Y', 'ranking_per_grado_1Y', 'ranking_finale', 'Perf_3Y', 'Vol_3Y', 'commissione', 
                            'PERF_corretto_3Y', 'Perf_1Y', 'Vol_1Y', 'commissione', 'PERF_corretto_1Y', 'note']]
                elif self.intermediario == 'CRV':
                    foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'podio', 'ranking_finale', 'ranking_finale_3Y', 'ranking_finale_1Y', 'Perf_3Y', 
                        'Vol_3Y', 'commissione', 'PERF_corretto_3Y', 'Perf_1Y', 'Vol_1Y', 'commissione', 'PERF_corretto_1Y', 'note']]

                # Cambio formato data
                foglio['data_di_avvio'] = foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
                # Ordinamento finale
                if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                    if metodo == 'singolo' or (metodo == 'doppio' and macro == 'LIQ_FOR'):
                        foglio.sort_values('ranking_PERF_3Y_corretto', ascending=True, inplace=True)
                    elif metodo == 'doppio':
                        foglio.sort_values('ranking_finale', ascending=True, inplace=True)
                elif self.intermediario == 'CRV':
                    foglio.sort_values('ranking_finale', ascending=False, inplace=True)
                    # Etichetta ND per i fondi senza dati
                    foglio['ranking_finale_1Y'] = foglio['ranking_finale_1Y'].fillna('ND')
                    foglio['ranking_finale_3Y'] = foglio['ranking_finale_3Y'].fillna('ND')
                    foglio['ranking_finale'] = foglio['ranking_finale'].fillna('ND')
                # Reindex
                foglio.reset_index(drop=True, inplace=True)

                # Crea foglio
                foglio.to_excel(writer, sheet_name=macro)

        writer.save()

    def aggiunta_prodotti_non_presenti(self):
        """Aggiunta foglio con i prodotti non presenti sulla piattaforma"""
        df = pd.read_csv(self.directory.joinpath('docs', 'prodotti_non_presenti.csv'), sep=';', index_col=None)
        if self.intermediario == 'CRV':
            df['ranking_finale'] = 'ND'
        with pd.ExcelWriter(self.file_ranking,  engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='NON_IN_PIATTAFORMA')
            
    def aggiunta_colonne(self, *colonne):
        """Aggiungi eventuali colonne presenti nel file_catalogo alla fine dei fogli del file di ranking
        
        Arguments:
            colonne {*args} = colonne da aggiungere al file di ranking
        """
        # TODO : con un if fissa l'argument colonne pari a nome nel caso di CRV, e pari a fondo_a_finestra nel caso di BPPB
        if self.intermediario == 'CRV':
            colonne = ['nome']
        else:
            colonne = []
        # Carica file_catalogo
        df_input = pd.read_excel(self.directory.joinpath(self.file_catalogo), index_col=None)
        # Carica file di ranking senza specificare un foglio per ottenerli tutti
        df_globale = pd.read_excel(self.file_ranking, sheet_name=None) # read all sheets
        with pd.ExcelWriter(self.file_ranking, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            for sheet in df_globale.keys():
                df = pd.read_excel(self.file_ranking, sheet_name=sheet, index_col=0)
                for colonna in colonne:
                    # Se la colonna da inserire ha lo stesso nome di una colonna già presente modificala
                    if colonna in df.columns:
                        colonna_label = colonna + '_2'
                    else:
                        colonna_label = colonna
                    df_input.rename(columns = {colonna : colonna_label}, inplace=True)
                    df_input_2 = df_input.loc[df_input['isin'].isin(df['ISIN']), ['isin', colonna_label]]
                    # Unisci i due file togliendo la colonna 'isin'
                    df = pd.merge(left=df, right=df_input_2, how='left', left_on='ISIN', right_on='isin').drop('isin', axis=1)
                df.to_excel(writer, sheet_name=sheet)

    def rank_formatted(self, metodo):
        """
        Formatta il file self.ranking per mettere in evidenza le classi blend e l'ordinamento definitivo.
        Formatta le intestazioni del file.
        """
        wb = load_workbook(filename='ranking.xlsx') # carica il file
        # Colora le micro blend
        print('\nsto formattando il file di ranking...')
        if self.intermediario == 'BPPB': # TODO aggiungi la globale high yield
            if metodo == 'singolo':
                micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 'OBB_BT' : 'Obblig. Euro breve term.',
                    'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale', 'OBB_EM' : 'Obblig. Paesi Emerg.'}
                micro_blend_classi_non_a_benchmark = ['FLEX_BVOL', 'FLEX_MAVOL', 'OPP', 'LIQ']
            elif metodo == 'doppio':
                pass
        elif self.intermediario == 'BPL':
            if metodo == 'singolo':
                micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 'AZ_GLOB' : 'Az. globale',
                    'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_EUR' : 'Obblig. Europa', 'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale',
                    'OBB_USA' : 'Obblig. Dollaro US all mat', 'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield'}
                micro_blend_classi_non_a_benchmark = ['BIL_MBVOL', 'BIL_AVOL', 'FLEX_PR', 'FLEX_DIN', 'OPP', 'LIQ', 'LIQ_FOR']
            elif metodo == 'doppio':
                micro_blend_classi_a_benchmark = {'LIQ' : 'Monetari Euro', 'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 'AZ_GLOB' : 'Az. globale',
                    'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_EUR' : 'Obblig. Europa', 'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale',
                    'OBB_USA' : 'Obblig. Dollaro US all mat', 'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield'}
        elif self.intermediario == 'CRV':
            micro_blend_classi_a_benchmark = {'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico', 'AZ_EM' : 'Az. paesi emerg. Mondo', 'AZ_GLOB' : 'Az. globale',
                'OBB_BT' : 'Obblig. Euro breve term.', 'OBB_MLT' : 'Obblig. Euro all maturities', 'OBB_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale',
                'OBB_EM' : 'Obblig. Paesi Emerg.', 'OBB_GLOB_HY' : 'Obblig. globale high yield'}
            micro_blend_classi_non_a_benchmark = ['FLEX_PR', 'FLEX_DIN', 'OPP', 'LIQ']

        for sheet in wb.sheetnames: 
            if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
                if metodo == 'singolo':
                    if sheet in micro_blend_classi_a_benchmark.keys():
                        foglio = wb[sheet] # attiva foglio
                        for cell in foglio['G']:
                            if cell.value == micro_blend_classi_a_benchmark[sheet]: # filtra per micro blend
                                cell.fill = PatternFill(fgColor="f4b084", fill_type='solid') # colora le micro blend
                        for cell in foglio['H']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='215967', fill_type='solid') # colora il ranking finale
                        for cell in foglio['J']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='92CDDC', fill_type='solid') # colora il ranking IR_3Y
                        for cell in foglio['P']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='31869B', fill_type='solid') # colora il ranking IR_3Y corretto
                        for cell in foglio['T']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='DAEEF3', fill_type='solid') # colora il ranking IR_1Y
                        for cell in foglio['Z']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='B7DEE8', fill_type='solid') # colora il ranking IR_1Y corretto
                        for cell in foglio['O']:
                            cell.number_format = '0.0000'
                        for cell in foglio['Y']:
                            cell.number_format = '0.0000'
                        for cell in foglio['N']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
                        for cell in foglio['X']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
                    elif sheet in micro_blend_classi_non_a_benchmark:
                        foglio = wb[sheet] # attiva foglio
                        for cell in foglio['H']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='92CDDC', fill_type='solid') # colora il ranking PERF_3Y
                        for cell in foglio['N']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='31869B', fill_type='solid') # colora il ranking PERF_3Y corretto
                        for cell in foglio['R']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='DAEEF3', fill_type='solid') # colora il ranking PERF_1Y
                        for cell in foglio['X']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='B7DEE8', fill_type='solid') # colora il ranking PERF_1Y corretto
                        for cell in foglio['M']:
                            cell.number_format = '0.0000'
                        for cell in foglio['W']:
                            cell.number_format = '0.0000'
                        for cell in foglio['L']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
                        for cell in foglio['V']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
                elif metodo == 'doppio':
                    if sheet in micro_blend_classi_a_benchmark.keys():
                        foglio = wb[sheet] # attiva foglio
                        for cell in foglio['F']:
                            if cell.value == micro_blend_classi_a_benchmark[sheet]: # filtra per micro blend
                                cell.fill = PatternFill(fgColor="f4b084", fill_type='solid') # colora le micro blend
                        for cell in foglio['H']:
                            if cell.value == micro_blend_classi_a_benchmark[sheet]: # filtra per micro blend
                                cell.fill = PatternFill(fgColor="f4b084", fill_type='solid') # colora le micro blend        
                        for cell in foglio['I']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='215967', fill_type='solid') # colora il ranking finale
                        for cell in foglio['J']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='92CDDC', fill_type='solid') # colora il ranking IR_3Y
                        for cell in foglio['P']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='31869B', fill_type='solid') # colora il ranking IR_3Y corretto
                        for cell in foglio['T']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='DAEEF3', fill_type='solid') # colora il ranking IR_1Y
                        for cell in foglio['Z']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='B7DEE8', fill_type='solid') # colora il ranking IR_1Y corretto
                        for cell in foglio['O']:
                            cell.number_format = '0.0000'
                        for cell in foglio['Y']:
                            cell.number_format = '0.0000'
                        for cell in foglio['N']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
                        for cell in foglio['X']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
                    elif sheet in micro_blend_classi_non_a_benchmark:
                        foglio = wb[sheet] # attiva foglio
                        for cell in foglio['H']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='92CDDC', fill_type='solid') # colora il ranking PERF_3Y
                        for cell in foglio['N']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='31869B', fill_type='solid') # colora il ranking PERF_3Y corretto
                        for cell in foglio['R']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='DAEEF3', fill_type='solid') # colora il ranking PERF_1Y
                        for cell in foglio['X']:
                            cell.alignment = Alignment(horizontal='center')
                            cell.fill = PatternFill(fgColor='B7DEE8', fill_type='solid') # colora il ranking PERF_1Y corretto
                        for cell in foglio['M']:
                            cell.number_format = '0.0000'
                        for cell in foglio['W']:
                            cell.number_format = '0.0000'
                        for cell in foglio['L']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
                        for cell in foglio['V']:
                            cell.number_format = numbers.FORMAT_PERCENTAGE_00
            elif self.intermediario == 'CRV':
                if sheet in micro_blend_classi_a_benchmark.keys() or sheet in micro_blend_classi_non_a_benchmark:
                    foglio = wb[sheet] # attiva foglio
                    for cell in foglio['H']: # colora il ranking finale
                        if cell.value == 'ND':
                            cell.fill = PatternFill(fgColor="595959", fill_type='solid')
                            cell.alignment = Alignment(horizontal='right')
                        else:
                            try:
                                if float(cell.value) <= 3.0: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="cd7f32", fill_type='solid')
                                    cell.number_format = '0.0000'
                                elif float(cell.value) <= 6.0: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="c0c0c0", fill_type='solid')
                                    cell.number_format = '0.0000'
                                elif float(cell.value) <= 9.1: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="cda434", fill_type='solid')
                                    cell.number_format = '0.0000'
                            except:
                                pass
                    for cell in foglio['I']: # colora il ranking finale 3Y
                        if cell.value == 'ND':
                            cell.fill = PatternFill(fgColor="595959", fill_type='solid')
                            cell.alignment = Alignment(horizontal='right')
                        else:
                            try:
                                if float(cell.value) <= 3.0: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="cd7f32", fill_type='solid')
                                    cell.number_format = '0.0000'
                                elif float(cell.value) <= 6.0: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="c0c0c0", fill_type='solid')
                                    cell.number_format = '0.0000'
                                elif float(cell.value) <= 9.1: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="cda434", fill_type='solid')
                                    cell.number_format = '0.0000'
                            except:
                                pass
                    for cell in foglio['J']: # colora il ranking finale 1Y
                        if cell.value == 'ND':
                            cell.fill = PatternFill(fgColor="595959", fill_type='solid')
                            cell.alignment = Alignment(horizontal='right')
                        else:
                            try:
                                if float(cell.value) <= 3.0: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="cd7f32", fill_type='solid')
                                    cell.number_format = '0.0000'
                                elif float(cell.value) <= 6.0: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="c0c0c0", fill_type='solid')
                                    cell.number_format = '0.0000'
                                elif float(cell.value) <= 9.1: # 1-3 bronzo 4-6 argento 7-9 oro
                                    cell.fill = PatternFill(fgColor="cda434", fill_type='solid')
                                    cell.number_format = '0.0000'
                            except:
                                pass
                    for cell in foglio['N']:
                        cell.number_format = '0.0000'
                    for cell in foglio['R']:
                        cell.number_format = '0.0000'
                    for cell in foglio['M']:
                        cell.number_format = numbers.FORMAT_PERCENTAGE_00
                    for cell in foglio['Q']:
                        cell.number_format = numbers.FORMAT_PERCENTAGE_00

        # Colora e cambia stile alle intestazioni
        for sheet in wb.sheetnames:
            foglio = wb[sheet]
            for column in foglio[1]:
                if column.value != '':
                    column.font = Font(name='Palatino Linotype', bold=True, italic=True) 
                    column.fill = PatternFill(fgColor="bfbfbf", fill_type='solid')
        
        # Ordina fogli
        if self.intermediario == 'BPPB':
            ordine = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_EM', 'OBB_GLOB', 'OBB_GLOB_HY', 'FLEX_BVOL', 'FLEX_MAVOL', 'OPP', 
                'LIQ']
            # for _ in wb._sheets:
            #     print(str(_)[12:-2])
            wb._sheets.sort(key=lambda i: ordine.index(str(i)[12:-2]))
        elif self.intermediario == 'BPL':
            ordine = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_EUR', 'OBB_CORP', 'OBB_GLOB', 'OBB_USA', 'OBB_EM', 'OBB_GLOB_HY',
                'BIL_MBVOL', 'BIL_AVOL', 'FLEX_PR', 'FLEX_DIN', 'OPP', 'LIQ', 'LIQ_FOR']
            wb._sheets.sort(key=lambda i: ordine.index(str(i)[12:-2]))
        elif self.intermediario == 'CRV':
            ordine = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_EM', 'OBB_GLOB', 'OBB_GLOB_HY', 'FLEX_PR', 'FLEX_DIN', 
                'OPP', 'LIQ']
            wb._sheets.sort(key=lambda i: ordine.index(str(i)[12:-2]))

        wb.save(self.file_ranking)

    def autofit(self):
        """
        Imposta la miglior lunghezza per le colonne selezionate.
        # TODO : accetta più di un foglio
        # TODO: accetta anche lettere per selezionare le colonne.
        # TODO: se columns è vuoto, autofit tutte le colonne.

        Parameters:
            sheet {string} = foglio excel da formattare
            columns {list} = lista contenente il numero o le lettere delle colonne da formattare. if not columns: formatta tutte le colonne del foglio
            min_width {list} = lista contenente la lunghezza massima in pixels della colonna, che l'autofit potrebbe non superare (usa None se non serve su una data colonna)
            max_width {list} = lista contenente la lunghezza massima in pixels della colonna, che l'autofit potrebbe superare (usa None se non serve su una data colonna)
        """
        if self.intermediario == 'BPPB' or self.intermediario == 'BPL':
            columns = range(1, 32)
        elif self.intermediario == 'CRV':
            columns = range(1, 21)
        xls_file = win32com.client.Dispatch("Excel.Application")
        xls_file.visible = False
        wb = xls_file.Workbooks.Open(Filename=self.directory.joinpath('ranking.xlsx').__str__())
        # openpyxl_wb = load_workbook(filename='ranking.xlsx') # carica il file
        for ws in wb.Sheets:
            for num, value in enumerate(columns):
                if value > 0: # la colonna 0 e le negative non esistono
                    ws.Columns(value).AutoFit()
                else:
                    continue
            wb.Save()
        xls_file.DisplayAlerts = False
        wb.Close(SaveChanges=True, Filename=self.file_ranking)
        xls_file.Quit()

    def creazione_liste_best_input(self):
        """
        Crea file csv da importare in Quantalys.it contenenti i best di ogni macro categoria direzionale.
        Directory in cui vengono salvati i file : './docs/import_liste_best_into_Q'
        """
        df_rank = pd.read_excel(self.file_ranking, index_col=None, sheet_name=1)
        classi_a_benchmark_BPPB = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM']
        classi_a_benchmark_BPL = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_EUR', 'OBB_CORP', 'OBB_GLOB', 'OBB_USA', 'OBB_EM', 'OBB_GLOB_HY']
        classi_a_benchmark_CRV = ['AZ_EUR', 'AZ_NA', 'AZ_PAC', 'AZ_EM', 'AZ_GLOB', 'OBB_BT', 'OBB_MLT', 'OBB_CORP', 'OBB_GLOB', 'OBB_EM', 'OBB_GLOB_HY']
        if self.intermediario == 'BPL':
            classi_a_benchmark = classi_a_benchmark_BPL
            # per ora solo per BPL
            if not os.path.exists(self.directory_input_liste_best):
                os.makedirs(self.directory_input_liste_best)
            print('sto creando le liste contenenti i fondi best di ogni macro categoria...')
            for classe in classi_a_benchmark:
                df_rank = pd.read_excel(self.file_ranking, index_col=None, sheet_name=classe)
                df = df_rank.loc[df_rank['Best_Worst'] == "best", ['ISIN', 'valuta']]
                df.columns = ['codice isin', 'divisa']
                df.to_csv(self.directory_input_liste_best.joinpath(classe + '_best' + '.csv'), sep=";", index=False)
        else:
            pass

    def zip_file(self):
        # TODO : Metti i file dei best in class nel file zip
        """
        Crea un file zip contenente il file_ranking e le note.
        """
        print(f'\nsto creando il file zip da inviare a {self.intermediario}...')
        rankZip = zipfile.ZipFile(self.file_zip, 'w')
        rankZip.write(self.file_ranking, compress_type=zipfile.ZIP_DEFLATED)
        rankZip.close()


if __name__ == '__main__':
    start = time.perf_counter()
    _ = Ranking(intermediario='BPL', t1='31/03/2022')
    # _.ranking_per_grado('doppio')
    _.merge_completo_liste()
    _.discriminazione_flessibili_e_bilanciati()
    _.rank('doppio')
    # _.aggiunta_colonne() # TODO: testa per intermediari diversi da CRV # 'nome' se CRV, 'fondo_a_finestra' se BPPB
    # _.rank_formatted()
    # _.aggiunta_prodotti_non_presenti() # TODO: testa per intermediari diversi da CRV
    # _.autofit()
    # _.creazione_liste_best_input()
    # _.zip_file()
    end = time.perf_counter()
    print("Elapsed time: ", round(end - start, 2), 'seconds')