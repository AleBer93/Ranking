import pandas as pd
import math
import numpy as np
import dateutil.relativedelta
import datetime

class Metodi_ranking():

    def __init__(self, foglio) -> None:
        self.foglio = foglio

    def codice_vecchio(self):
        ### Metodo doppio IR ###
        # # Creazione IR_corretto_3Y
        # if self.intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
        #     foglio['IR_corretto_3Y'] = ((df['Information_Ratio_3Y'] * (df['TEV_3Y'] / 100) ) - (df['commissione'] / df['anni_detenzione'])) / (df['TEV_3Y'] / 100)
        #     foglio['IR_corretto_1Y'] = ((df['Information_Ratio_1Y'] * (df['TEV_1Y'] / 100) ) - (df['commissione'] / df['anni_detenzione'])) / (df['TEV_1Y'] / 100)
        # else:
        #     foglio['IR_corretto_3Y'] = ((df['Information_Ratio_3Y'] * (df['TEV_3Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)) / (df['TEV_3Y'] / 100)
        #     foglio['IR_corretto_1Y'] = ((df['Information_Ratio_1Y'] * (df['TEV_1Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)) / (df['TEV_1Y'] / 100)
        # # Note
        # foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & foglio['Best_Worst_1Y'].isnull(), 'note'] = 'Ha 1 anno, ma non è in classifica ad un anno.'
        # # foglio.loc[(foglio['data_di_avvio'] > t0_1Y) & foglio['Information_Ratio_1Y'].notnull(), 'note'] = 'Non ha 1 anno, ma possiede dati a un anno.' Nota fuorviante
        # foglio.loc[(foglio['data_di_avvio'] > self.t0_1Y) & foglio['Information_Ratio_1Y'].notnull(), ['Information_Ratio_1Y', 'TEV_1Y', 'IR_corretto_1Y']] = np.nan
        # foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & foglio['Best_Worst_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non è in classifica a tre anni.'
        # # foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Information_Ratio_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.' Nota fuorviante
        # foglio.loc[(foglio['data_di_avvio'] > self.t0_3Y) & foglio['Information_Ratio_3Y'].notnull(), ['Information_Ratio_3Y', 'TEV_3Y', 'IR_corretto_3Y']] = np.nan
        # # Ranking finale
        # if self.soluzioni[macro] == 1:
        #     # Fondi best blend - Gerarchia : semi_attivo, attivo, molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        # elif self.soluzioni[macro] == 2:
        #     # Fondi best blend - Gerarchia : attivo, semi_attivo, molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        # elif self.soluzioni[macro] == 3:
        #     # Fondi best blend - Gerarchia : (semi_attivo & attivo), molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : (semi_attivo & attivo), molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        # elif self.soluzioni[macro] == 4:
        #     # Fondi best blend - Gerarchia : semi_attivo & attivo & molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : semi_attivo & attivo & molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        
        ### in fondo a SOR_DSR ###
        # Seleziona colonne utili
        # if self.metodo == 'singolo':
        #     foglio = foglio[
        #         ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sortino_3Y', 'ranking_SO_3Y',
        #         'quartile_SO_3Y', 'terzile_SO_3Y', 'DSR_3Y', 'commissione', 'SO_corretto_3Y', 'ranking_SO_3Y_corretto',
        #         'quartile_SO_corretto_3Y', 'terzile_SO_corretto_3Y', 'Sortino_1Y', 'ranking_SO_1Y', 'quartile_SO_1Y',
        #         'terzile_SO_1Y', 'DSR_1Y', 'commissione', 'SO_corretto_1Y', 'ranking_SO_1Y_corretto',
        #         'quartile_SO_corretto_1Y', 'terzile_SO_corretto_1Y', 'SFDR', 'note']
        #     ]
        # elif self.metodo == 'doppio':
        #     foglio = foglio[['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sortino_3Y', 'DSR_3Y',
        #         'commissione', 'SO_corretto_3Y', 'ranking_SO_3Y_corretto', 'Sortino_1Y', 'DSR_1Y', 'commissione',
        #         'SO_corretto_1Y', 'ranking_SO_1Y_corretto', 'SFDR', 'note']
        #     ]
        # elif self.metodo == 'linearizzazione':
        #     foglio = foglio[
        #         ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'podio', 'ranking_finale',
        #         'ranking_finale_3Y', 'ranking_finale_1Y', 'Sortino_3Y', 'DSR_3Y', 'commissione', 'SO_corretto_3Y',
        #         'Sortino_1Y', 'DSR_1Y', 'commissione', 'SO_corretto_1Y', 'note']
        #     ]
        
        # Cambio formato data
        # foglio['data_di_avvio'] = foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
        # # Ordinamento finale
        # if self.metodo == 'singolo' or self.metodo == 'doppio':
        #     foglio.sort_values('ranking_SO_3Y_corretto', ascending=True, inplace=True)
        # elif self.metodo == 'linearizzazione':
        #     foglio.sort_values('ranking_finale', ascending=False, inplace=True)
        #     # Etichetta ND per i fondi senza dati
        #     foglio['ranking_finale_1Y'] = foglio['ranking_finale_1Y'].fillna('ND')
        #     foglio['ranking_finale_3Y'] = foglio['ranking_finale_3Y'].fillna('ND')
        #     foglio['ranking_finale'] = foglio['ranking_finale'].fillna('ND')
        # # Reindex
        # foglio.reset_index(drop=True, inplace=True)

        ### in fondo a SHA_VOL ###
        # Seleziona colonne utili
        # if self.metodo == 'singolo':
        #     foglio = foglio[
        #         ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sharpe_3Y', 'ranking_SH_3Y',
        #         'quartile_SH_3Y', 'terzile_SH_3Y', 'Vol_3Y', 'commissione', 'SH_corretto_3Y', 'ranking_SH_3Y_corretto',
        #         'quartile_SH_corretto_3Y', 'terzile_SH_corretto_3Y', 'Sharpe_1Y', 'ranking_SH_1Y', 'quartile_SH_1Y',
        #         'terzile_SH_1Y', 'Vol_1Y', 'commissione', 'SH_corretto_1Y', 'ranking_SH_1Y_corretto',
        #         'quartile_SH_corretto_1Y', 'terzile_SH_corretto_1Y', 'SFDR', 'note']
        #     ]
        # elif self.metodo == 'doppio':
        #     foglio = foglio[
        #         ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Sharpe_3Y', 'Vol_3Y', 'commissione',
        #         'SH_corretto_3Y', 'ranking_SH_3Y_corretto', 'Sharpe_1Y', 'Vol_1Y', 'commissione', 'SH_corretto_1Y',
        #         'ranking_SH_1Y_corretto', 'SFDR', 'note']
        #     ]
        # elif self.metodo == 'linearizzazione':
        #     foglio = foglio[
        #         ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'podio', 'ranking_finale',
        #         'ranking_finale_3Y', 'ranking_finale_1Y', 'Sharpe_3Y', 'Vol_3Y', 'commissione', 'SH_corretto_3Y',
        #         'Sharpe_1Y', 'Vol_1Y', 'commissione', 'SH_corretto_1Y', 'note']
        #     ]
        
        # Cambio formato data
        # foglio['data_di_avvio'] = foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
        # # Ordinamento finale
        # if self.metodo == 'singolo' or self.metodo == 'doppio':
        #     foglio.sort_values('ranking_SH_3Y_corretto', ascending=True, inplace=True)
        # elif self.metodo == 'linearizzazione':
        #     foglio.sort_values('ranking_finale', ascending=False, inplace=True)
        #     # Etichetta ND per i fondi senza dati
        #     foglio['ranking_finale_1Y'] = foglio['ranking_finale_1Y'].fillna('ND')
        #     foglio['ranking_finale_3Y'] = foglio['ranking_finale_3Y'].fillna('ND')
        #     foglio['ranking_finale'] = foglio['ranking_finale'].fillna('ND')
        # # Reindex
        # foglio.reset_index(drop=True, inplace=True)

        ### Metodo doppio PERF ###
        # # Creazione PERF_corretto_3Y
        # if self.intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
        #     foglio['PERF_corretto_3Y'] = (df['Perf_3Y'] / 100) - (df['commissione'] / df['anni_detenzione'])
        #     foglio['PERF_corretto_1Y'] = (df['Perf_1Y'] / 100) - (df['commissione'] / df['anni_detenzione'])
        # else:
        #     foglio['PERF_corretto_3Y'] = (df['Perf_3Y'] / 100) - (df['commissione'] / self.anni_detenzione)
        #     foglio['PERF_corretto_1Y'] = (df['Perf_1Y'] / 100) - (df['commissione'] / self.anni_detenzione)
        # # Note
        # foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & foglio['Best_Worst_1Y'].isnull(), 'note'] = 'Ha 1 anno, ma non è in classifica ad un anno.'
        # # foglio.loc[(foglio['data_di_avvio'] > t0_1Y) & foglio['Perf_1Y'].notnull(), 'note'] = 'Non ha 1 anno, ma possiede dati a un anno.' Nota fuorviante
        # foglio.loc[(foglio['data_di_avvio'] > self.t0_1Y) & foglio['Perf_1Y'].notnull(), ['Perf_1Y', 'Vol_1Y', 'PERF_corretto_1Y']] = np.nan
        # # Aggiunta nota per i fondi che possiedono dati a tre anni pur non avendo tre anni di vita, e ad un anno non avendo un anno di vita
        # foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & foglio['Best_Worst_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non è in classifica a tre anni.'
        # # foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Perf_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.' Nota fuorviante
        # foglio.loc[(foglio['data_di_avvio'] > self.t0_3Y) & foglio['Perf_3Y'].notnull(), ['Perf_3Y', 'Vol_3Y', 'PERF_corretto_3Y']] = np.nan
        # # Ranking finale
        # if self.soluzioni[macro] == 1:
        #     # Fondi best blend - Gerarchia : semi_attivo, attivo, molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        # elif self.soluzioni[macro] == 2:
        #     # Fondi best blend - Gerarchia : attivo, semi_attivo, molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        # elif self.soluzioni[macro] == 3:
        #     # Fondi best blend - Gerarchia : (semi_attivo & attivo), molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : (semi_attivo & attivo), molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        # elif self.soluzioni[macro] == 4:
        #     # Fondi best blend - Gerarchia : semi_attivo & attivo & molto_attivo
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi best non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'best') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] != 'best') & (foglio['Best_Worst_1Y'] == 'best') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     # Fondi worst blend - Gerarchia : semi_attivo & attivo & molto_attivo
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] == self.classi_metodo_doppio[macro]) &  (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
        #     # Fondi worst non blend
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'] == 'worst') & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        #     ultimo_elemento_ordinato = foglio['ranking_finale'].max()
        #     if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
        #     foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale'] = foglio.loc[(foglio['micro_categoria'] != self.classi_metodo_doppio[macro]) & (foglio['Best_Worst_3Y'].isnull()) & (foglio['Best_Worst_1Y'] == 'worst') & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        pass

    def singolo(self, t0_1Y, t0_3Y, anni_detenzione):
        # TODO: da sistemare da capo
        self.foglio['ranking_IR_1Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio['Information_Ratio_1Y'].notnull()), 'Information_Ratio_1Y'
        ].rank(method='first', na_option='bottom', ascending=False)
        # Quartile IR_1Y
        self.foglio['quartile_IR_1Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio['Information_Ratio_1Y'].notnull()), 'Information_Ratio_1Y'
        ].apply(lambda x: 'best' if x > self.foglio['Information_Ratio_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
        # Terzile IR_1Y
        self.foglio['terzile_IR_1Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio['Information_Ratio_1Y'].notnull()), 'Information_Ratio_1Y'
        ].apply(lambda x: 'best' if x > self.foglio['Information_Ratio_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')
        # Creazione IR_corretto_1Y
        self.foglio['IR_corretto_1Y'] = (
            (df['Information_Ratio_1Y'] * (df['TEV_1Y'] / 100) ) - (df['commissione'] / anni_detenzione)) / (df['TEV_1Y'] / 100)
        # Rank IR_corretto_1Y
        self.foglio['ranking_IR_1Y_corretto'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'
        ].rank(method='first', na_option='bottom', ascending=False)
        # Quartile IR_1Y corretto
        self.foglio['quartile_IR_corretto_1Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'
        ].apply(lambda x: 'best' if x > self.foglio['IR_corretto_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
        # Terzile IR_1Y corretto
        self.foglio['terzile_IR_corretto_1Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'
        ].apply(lambda x: 'best' if x > self.foglio['IR_corretto_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')

        # Rank IR_3Y
        self.foglio['ranking_IR_3Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'].notnull())
            & (self.foglio['Information_Ratio_3Y'].notnull()), 'Information_Ratio_3Y'
        ].rank(method='first', na_option='keep', ascending=False)
        # Note
        self.foglio.loc[(self.foglio['data_di_avvio'] < t0_3Y) & self.foglio['Best_Worst'].isnull(), 'note'] = 'Ha 3 anni, ma non è in classifica.'
        self.foglio.loc[(self.foglio['data_di_avvio'] > t0_3Y) & self.foglio['Information_Ratio_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
        # Quartile IR_3Y
        self.foglio['quartile_IR_3Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'].notnull())
            & (self.foglio['Information_Ratio_3Y'].notnull()), 'Information_Ratio_3Y'
        ].apply(lambda x: 'best' if x > self.foglio['Information_Ratio_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
        # Terzile IR_3Y
        self.foglio['terzile_IR_3Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'].notnull())
            & (self.foglio['Information_Ratio_3Y'].notnull()), 'Information_Ratio_3Y'
        ].apply(lambda x: 'best' if x > self.foglio['Information_Ratio_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
        # Creazione IR_corretto_3Y
        self.foglio['IR_corretto_3Y'] = (
            (df['Information_Ratio_3Y'] * (df['TEV_3Y'] / 100)) - (df['commissione'] / anni_detenzione)) / (df['TEV_3Y'] / 100)
        # Rank IR_corretto_3Y
        self.foglio['ranking_IR_3Y_corretto'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'].notnull())
            & (self.foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'
        ].rank(method='first', na_option='bottom', ascending=False)
        # Quartile IR_3Y corretto
        self.foglio['quartile_IR_corretto_3Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'].notnull())
            & (self.foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'
        ].apply(lambda x: 'best' if x > self.foglio['IR_corretto_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
        # Terzile IR_3Y corretto
        self.foglio['terzile_IR_corretto_3Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'].notnull())
            & (self.foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'
        ].apply(lambda x: 'best' if x > self.foglio['IR_corretto_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
        
        # Ranking finale
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'best') & (self.foglio['IR_corretto_3Y'].notnull())
            & (self.foglio['micro_categoria'] == self.classi_metodo_singolo[macro]), 'ranking_finale'
        ] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'best') & (self.foglio['IR_corretto_3Y'].notnull())
            & (self.foglio['micro_categoria'] == self.classi_metodo_singolo[macro]), 'IR_corretto_3Y'
            ].rank(method='first', na_option='bottom', ascending=False)
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'best')
            & (self.foglio['micro_categoria'] != self.classi_metodo_singolo[macro]), 'ranking_finale'
        ] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'best')
            & (self.foglio['micro_categoria'] != self.classi_metodo_singolo[macro]), 'IR_corretto_3Y'
            ].rank(method='first', na_option='bottom', ascending=False) + self.foglio['ranking_finale'].max()
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'worst')
            & (self.foglio['micro_categoria'] == self.classi_metodo_singolo[macro]), 'ranking_finale'
        ] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'worst')
            & (self.foglio['micro_categoria'] == self.classi_metodo_singolo[macro]), 'IR_corretto_3Y'
            ].rank(method='first', na_option='bottom', ascending=False) + self.foglio['ranking_finale'].max()
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'worst')
            & (self.foglio['micro_categoria'] != self.classi_metodo_singolo[macro]), 'ranking_finale'
        ] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio['Best_Worst'] == 'worst')
            & (self.foglio['micro_categoria'] != self.classi_metodo_singolo[macro]), 'IR_corretto_3Y'
            ].rank(method='first', na_option='bottom', ascending=False) + self.foglio['ranking_finale'].max()
        
        return self.foglio

    def classifica(self, indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione):
        if indicatore == 'IR':
            primo_indicatore_tre_anni = 'Information_Ratio_3Y'
            secondo_indicatore_tre_anni = 'TEV_3Y'
            primo_indicatore_un_anno = 'Information_Ratio_1Y'
            secondo_indicatore_un_anno = 'TEV_1Y'
            indicatore_corretto_tre_anni = 'IR_corretto_3Y'
            indicatore_corretto_un_anno  = 'IR_corretto_1Y'
            # Creazione indicatore corretto 1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            # Creazione indicatore corretto 3Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
            else:    
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
        if indicatore == 'SO':
            primo_indicatore_tre_anni = 'Sortino_3Y'
            secondo_indicatore_tre_anni = 'DSR_3Y'
            primo_indicatore_un_anno = 'Sortino_1Y'
            secondo_indicatore_un_anno = 'DSR_1Y'
            indicatore_corretto_tre_anni = 'SO_corretto_3Y'
            indicatore_corretto_un_anno  = 'SO_corretto_1Y'
            # Creazione indicatore corretto 1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            # Creazione indicatore corretto 3Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
            else:    
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
            # Codice vecchio
            # # Rank SO_1Y
            # foglio['ranking_SO_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['Sortino_1Y'].notnull()), 'Sortino_1Y'
            # ].rank(method='first', na_option='bottom', ascending=False)
            # # Quartile SO_1Y
            # foglio['quartile_SO_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['Sortino_1Y'].notnull()), 'Sortino_1Y'
            # ].apply(lambda x: 'best' if x > foglio['Sortino_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SO_1Y
            # foglio['terzile_SO_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['Sortino_1Y'].notnull()), 'Sortino_1Y'
            # ].apply(lambda x: 'best' if x > foglio['Sortino_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')
            # # Creazione SO_corretto_1Y
            # if self.intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
            #     foglio['SO_corretto_1Y'] = (
            #         (df['Sortino_1Y'] * (df['DSR_1Y'] / 100) ) - (df['commissione'] / df['anni_detenzione'])
            #     ) / (df['DSR_1Y'] / 100)
            # else:
            #     foglio['SO_corretto_1Y'] = (
            #         (df['Sortino_1Y'] * (df['DSR_1Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)
            #     ) / (df['DSR_1Y'] / 100)
            # # Rank SO_corretto_1Y
            # foglio['ranking_SO_1Y_corretto'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'
            # ].rank(method='first', na_option='bottom', ascending=False)
            # # Quartile SO_1Y corretto
            # foglio['quartile_SO_corretto_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'
            # ].apply(lambda x: 'best' if x > foglio['SO_corretto_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SO_1Y corretto
            # foglio['terzile_SO_corretto_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'
            # ].apply(lambda x: 'best' if x > foglio['SO_corretto_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')

            # # Rank SO_3Y
            # foglio['ranking_SO_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'Sortino_3Y'
            # ].rank(method='first', na_option='keep', ascending=False)
            
            # # Quartile SO_3Y TOGLI IL BEST_WORST
            # foglio['quartile_SO_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'Sortino_3Y'
            # ].apply(lambda x: 'best' if x > foglio['Sortino_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SO_3Y
            # foglio['terzile_SO_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'Sortino_3Y'
            # ].apply(lambda x: 'best' if x > foglio['Sortino_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
            # # Creazione SO_corretto_3Y
            # if self.intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
            #     foglio['SO_corretto_3Y'] = (
            #         (df['Sortino_3Y'] * (df['DSR_3Y'] / 100) ) - (df['commissione'] / df['anni_detenzione'])
            #     ) / (df['DSR_3Y'] / 100)
            # else:    
            #     foglio['SO_corretto_3Y'] = (
            #         (df['Sortino_3Y'] * (df['DSR_3Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)
            #     ) / (df['DSR_3Y'] / 100)
            # # Rank SO_corretto_3Y
            # foglio['ranking_SO_3Y_corretto'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'
            # ].rank(method='first', na_option='bottom', ascending=False)
            # # Quartile SO_3Y corretto
            # foglio['quartile_SO_corretto_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'
            # ].apply(lambda x: 'best' if x > foglio['SO_corretto_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SO_3Y corretto
            # foglio['terzile_SO_corretto_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'
            # ].apply(lambda x: 'best' if x > foglio['SO_corretto_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
            
            # if self.metodo == 'singolo':
            #     # Note
            #     foglio.loc[
            #         (foglio['data_di_avvio'] < self.t0_3Y) & foglio['Sortino_3Y'].isnull(), 'note'
            #     ] = 'Ha 3 anni, ma non possiede dati a tre anni.'
            #     foglio.loc[
            #         (foglio['data_di_avvio'] > self.t0_3Y) & foglio['Sortino_3Y'].notnull(), 'note'
            #     ] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            # elif self.metodo == 'doppio':
            #     # Note
            #     foglio.loc[
            #         (foglio['data_di_avvio'] < self.t0_1Y) & foglio['Sortino_1Y'].isnull(), 'note'
            #     ] = 'Ha 1 anno, ma non è in classifica ad un anno.'
            #     # foglio.loc[(foglio['data_di_avvio'] > t0_1Y) & foglio['Sortino_1Y'].notnull(), 'note'] = 'Non ha 1 anno, ma possiede dati a un anno.'
            #     foglio.loc[
            #         (foglio['data_di_avvio'] > self.t0_1Y) & foglio['Sortino_1Y'].notnull(), ['Sortino_1Y', 'DSR_1Y', 'SO_corretto_1Y']
            #     ] = np.nan
            #     foglio.loc[
            #         (foglio['data_di_avvio'] < self.t0_3Y) & foglio['Sortino_3Y'].isnull(), 'note'
            #     ] = 'Ha 3 anni, ma non è in classifica a tre anni.'
            #     # foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Sortino_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            #     foglio.loc[
            #         (foglio['data_di_avvio'] > self.t0_3Y) & foglio['Sortino_3Y'].notnull(), ['Sortino_3Y', 'DSR_3Y', 'SO_corretto_3Y']
            #     ] = np.nan
        elif indicatore == 'SH':
            primo_indicatore_tre_anni = 'Sharpe_3Y'
            secondo_indicatore_tre_anni = 'Vol_3Y'
            primo_indicatore_un_anno = 'Sharpe_1Y'
            secondo_indicatore_un_anno = 'Vol_1Y'
            indicatore_corretto_tre_anni = 'SH_corretto_3Y'
            indicatore_corretto_un_anno  = 'SH_corretto_1Y'
            # Creazione indicatore corretto 1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            # Creazione indicatore corretto 3Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
            else:    
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
            # Codice vecchio
            # # Rank SH_1Y
            # foglio['ranking_SH_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['Sharpe_1Y'].notnull()), 'Sharpe_1Y'
            # ].rank(method='first', na_option='bottom', ascending=False)
            # # Quartile SH_1Y
            # foglio['quartile_SH_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['Sharpe_1Y'].notnull()), 'Sharpe_1Y'
            # ].apply(lambda x: 'best' if x > foglio['Sharpe_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SH_1Y
            # foglio['terzile_SH_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['Sharpe_1Y'].notnull()), 'Sharpe_1Y'
            # ].apply(lambda x: 'best' if x > foglio['Sharpe_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')
            # # Creazione SH_corretto_1Y
            # if self.intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
            #     foglio['SH_corretto_1Y'] = (
            #         (df['Sharpe_1Y'] * (df['Vol_1Y'] / 100) ) - (df['commissione'] / df['anni_detenzione'])
            #     ) / (df['Vol_1Y'] / 100)
            # else:
            #     foglio['SH_corretto_1Y'] = (
            #         (df['Sharpe_1Y'] * (df['Vol_1Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)
            #     ) / (df['Vol_1Y'] / 100)
            # # Rank SH_corretto_1Y
            # foglio['ranking_SH_1Y_corretto'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'
            # ].rank(method='first', na_option='bottom', ascending=False)
            # # Quartile SH_1Y corretto
            # foglio['quartile_SH_corretto_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'
            # ].apply(lambda x: 'best' if x > foglio['SH_corretto_1Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SH_1Y corretto
            # foglio['terzile_SH_corretto_1Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'
            # ].apply(lambda x: 'best' if x > foglio['SH_corretto_1Y'].quantile(0.33, interpolation = 'linear') else 'worst')

            # # Rank SH_3Y
            # foglio['ranking_SH_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'Sharpe_3Y'
            # ].rank(method='first', na_option='keep', ascending=False)
            # # Quartile SH_3Y
            # foglio['quartile_SH_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'Sharpe_3Y'
            # ].apply(lambda x: 'best' if x > foglio['Sharpe_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SH_3Y
            # foglio['terzile_SH_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'Sharpe_3Y'
            # ].apply(lambda x: 'best' if x > foglio['Sharpe_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
            # # Creazione SH_corretto_3Y
            # if self.intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
            #     foglio['SH_corretto_3Y'] = (
            #         (df['Sharpe_3Y'] * (df['Vol_3Y'] / 100) ) - (df['commissione'] / df['anni_detenzione'])
            #     ) / (df['Vol_3Y'] / 100)
            # else:
            #     foglio['SH_corretto_3Y'] = (
            #         (df['Sharpe_3Y'] * (df['Vol_3Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)
            #     ) / (df['Vol_3Y'] / 100)
            # # Rank SH_corretto_3Y
            # foglio['ranking_SH_3Y_corretto'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'
            # ].rank(method='first', na_option='bottom', ascending=False)
            # # Quartile SH_3Y corretto
            # foglio['quartile_SH_corretto_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'
            # ].apply(lambda x: 'best' if x > foglio['SH_corretto_3Y'].quantile(0.25, interpolation = 'linear') else 'worst')
            # # Terzile SH_3Y corretto
            # foglio['terzile_SH_corretto_3Y'] = foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'
            # ].apply(lambda x: 'best' if x > foglio['SH_corretto_3Y'].quantile(0.33, interpolation = 'linear') else 'worst')
            
            # if self.metodo == 'singolo':
            #     # Note
            #     foglio.loc[(
            #         foglio['data_di_avvio'] < self.t0_3Y) & foglio['Sharpe_3Y'].isnull(), 'note'
            #     ] = 'Ha 3 anni, ma non possiede dati a tre anni.'
            #     foglio.loc[
            #         (foglio['data_di_avvio'] > self.t0_3Y) & foglio['Sharpe_3Y'].notnull(), 'note'
            #     ] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            # elif self.metodo == 'doppio':
            #     # Note
            #     foglio.loc[
            #         (foglio['data_di_avvio'] < self.t0_1Y) & foglio['Sharpe_1Y'].isnull(), 'note'
            #     ] = 'Ha 1 anno, ma non è in classifica ad un anno.'
            #     # foglio.loc[(foglio['data_di_avvio'] > t0_1Y) & foglio['Sharpe_1Y'].notnull(), 'note'] = 'Non ha 1 anno, ma possiede dati a un anno.'
            #     foglio.loc[
            #         (foglio['data_di_avvio'] > self.t0_1Y) & foglio['Sharpe_1Y'].notnull(), ['Sharpe_1Y', 'Vol_1Y', 'SH_corretto_1Y']
            #     ] = np.nan
            #     foglio.loc[
            #         (foglio['data_di_avvio'] < self.t0_3Y) & foglio['Sharpe_3Y'].isnull(), 'note'
            #     ] = 'Ha 3 anni, ma non è in classifica a tre anni.'
            #     # foglio.loc[(foglio['data_di_avvio'] > t0_3Y) & foglio['Sharpe_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            #     foglio.loc[
            #         (foglio['data_di_avvio'] > self.t0_3Y) & foglio['Sharpe_3Y'].notnull(), ['Sharpe_3Y', 'Vol_3Y', 'SH_corretto_3Y']
            #     ] = np.nan
        elif indicatore == 'PERF':
            primo_indicatore_tre_anni = 'Perf_3Y'
            secondo_indicatore_tre_anni = 'Vol_3Y'
            primo_indicatore_un_anno = 'Perf_1Y'
            secondo_indicatore_un_anno = 'Vol_1Y'
            indicatore_corretto_tre_anni = 'PERF_corretto_3Y'
            indicatore_corretto_un_anno  = 'PERF_corretto_1Y'
            # Creazione PERF_corretto_3Y e PERF_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
            else:
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / anni_detenzione)
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / anni_detenzione)
                
        # Rank finale 1Y
        self.foglio['ranking_finale_1Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno
        ].rank(method='first', na_option='bottom', ascending=False)

        # Rank finale 3Y
        self.foglio['ranking_finale_3Y'] = self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni
        ].rank(method='first', na_option='bottom', ascending=False)
        
        # Note
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & self.foglio[primo_indicatore_un_anno].isnull(), 'note'
        ] = 'Ha 1 anno, ma non è in classifica ad un anno.'
        # self.foglio.loc[(self.foglio['data_di_avvio'] > t0_1Y) & self.foglio[primo_indicatore_un_anno].notnull(), 'note'] = 'Non ha 1 anno, ma possiede dati a un anno.'
        self.foglio.loc[
            (self.foglio['data_di_avvio'] > t0_1Y) & self.foglio[primo_indicatore_un_anno].notnull(), [primo_indicatore_un_anno, secondo_indicatore_un_anno, indicatore_corretto_un_anno]
        ] = np.nan
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & self.foglio[primo_indicatore_tre_anni].isnull(), 'note'
        ] = 'Ha 3 anni, ma non è in classifica a tre anni.'
        # self.foglio.loc[(self.foglio['data_di_avvio'] > t0_3Y) & self.foglio['Sortino_3Y'].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
        self.foglio.loc[
            (self.foglio['data_di_avvio'] > t0_3Y) & self.foglio[primo_indicatore_tre_anni].notnull(), [primo_indicatore_tre_anni, secondo_indicatore_tre_anni, indicatore_corretto_tre_anni]
        ] = np.nan

        # Cambio formato data
        self.foglio['data_di_avvio'] = self.foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
        # Ordinamento finale
        self.foglio.sort_values('ranking_finale_1Y', ascending=True, inplace=True)
        self.foglio.sort_values('ranking_finale_3Y', ascending=True, inplace=True)
        # Reindex
        self.foglio.reset_index(drop=True, inplace=True)
        # Seleziona colonne utili

        self.foglio = self.foglio[
            ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', primo_indicatore_tre_anni, secondo_indicatore_tre_anni,
            'commissione', indicatore_corretto_tre_anni, 'ranking_finale_3Y', primo_indicatore_un_anno, secondo_indicatore_un_anno,
            'commissione', indicatore_corretto_un_anno, 'ranking_finale_1Y', 'SFDR', 'note']
        ]

        return self.foglio

    def classifica_con_linearizzazione(self, indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione):
        self.foglio = self.classifica(indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione)
        if indicatore == 'SO':
            primo_indicatore_tre_anni = 'Sortino_3Y'
            secondo_indicatore_tre_anni = 'DSR_3Y'
            primo_indicatore_un_anno = 'Sortino_1Y'
            secondo_indicatore_un_anno = 'DSR_1Y'
            indicatore_corretto_tre_anni = 'SO_corretto_3Y'
            indicatore_corretto_un_anno  = 'SO_corretto_1Y'
            # Rinomina la colonna
            self.foglio.rename(columns = {'ranking_finale_3Y': 'classifica_finale'}, inplace = True)
        elif indicatore == 'SH':
            primo_indicatore_tre_anni = 'Sharpe_3Y'
            secondo_indicatore_tre_anni = 'Vol_3Y'
            primo_indicatore_un_anno = 'Sharpe_1Y'
            secondo_indicatore_un_anno = 'Vol_1Y'
            indicatore_corretto_tre_anni = 'SH_corretto_3Y'
            indicatore_corretto_un_anno  = 'SH_corretto_1Y'
            # Rinomina la colonna
            self.foglio.rename(columns = {'ranking_finale_3Y': 'classifica_finale'}, inplace = True)
        # Cambio formato data
        self.foglio['data_di_avvio'] = pd.to_datetime(self.foglio['data_di_avvio'], dayfirst=True)
        # Ranking finale
        minimo_3Y = min(self.foglio.loc[self.foglio['classifica_finale'].notnull(), 'classifica_finale'])
        massimo_3Y = max(self.foglio.loc[self.foglio['classifica_finale'].notnull(), 'classifica_finale'])
        self.foglio.loc[
            self.foglio['classifica_finale'].notnull(), 'ranking_finale'
        ] = 1 - (8 / (massimo_3Y - minimo_3Y)) + (8 * (massimo_3Y + minimo_3Y - self.foglio.loc[
            self.foglio['classifica_finale'].notnull(), 'classifica_finale'
            ]) / (massimo_3Y - minimo_3Y))
        # Cambio formato data
        self.foglio['data_di_avvio'] = self.foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
        # Seleziona colonne utili
        self.foglio = self.foglio[
            ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', primo_indicatore_tre_anni, secondo_indicatore_tre_anni,
            'commissione', indicatore_corretto_tre_anni, 'classifica_finale', 'ranking_finale', 
            primo_indicatore_un_anno, secondo_indicatore_un_anno, 'commissione', indicatore_corretto_un_anno, 'ranking_finale_1Y', 'note']
        ]

        return self.foglio

    def doppio(self, macro, indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione, soluzioni, classi_metodo_doppio):
        """Questo metodo può essere utilizzato da classi a benchmark, aggiungi un assert nel metodo di ranking!

        Arguments:
            macro {_type_} -- _description_
            indicatore {_type_} -- _description_
            t0_1Y {_type_} -- _description_
            t0_3Y {_type_} -- _description_
            intermediario {_type_} -- _description_
            anni_detenzione {_type_} -- _description_
            soluzioni {_type_} -- _description_
            classi_metodo_doppio {_type_} -- _description_

        Returns:
            _type_ -- _description_
        """
        if indicatore == 'IR':
            primo_indicatore_tre_anni = 'Information_Ratio_3Y'
            secondo_indicatore_tre_anni = 'TEV_3Y'
            primo_indicatore_un_anno = 'Information_Ratio_1Y'
            secondo_indicatore_un_anno = 'TEV_1Y'
            indicatore_corretto_tre_anni = 'IR_corretto_3Y'
            indicatore_corretto_un_anno  = 'IR_corretto_1Y'
            # Creazione IR_corretto_3Y e IR_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) -
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
        elif indicatore == 'PERF':
            primo_indicatore_tre_anni = 'Perf_3Y'
            secondo_indicatore_tre_anni = 'Vol_3Y'
            primo_indicatore_un_anno = 'Perf_1Y'
            secondo_indicatore_un_anno = 'Vol_1Y'
            indicatore_corretto_tre_anni = 'PERF_corretto_3Y'
            indicatore_corretto_un_anno  = 'PERF_corretto_1Y'
            # Creazione PERF_corretto_3Y e PERF_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
            else:
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / anni_detenzione)
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / anni_detenzione)
        # Note
        self.foglio.loc[(self.foglio['data_di_avvio'] < t0_1Y) & self.foglio['Best_Worst_1Y'].isnull(), 'note'] = 'Ha 1 anno, ma non è in classifica ad un anno.'
        # self.foglio.loc[(self.foglio['data_di_avvio'] > t0_1Y) & self.foglio[primo_indicatore_un_anno].notnull(), 'note'] = 'Non ha 1 anno, ma possiede dati a un anno.' Nota fuorviante
        self.foglio.loc[(self.foglio['data_di_avvio'] > t0_1Y) & self.foglio[primo_indicatore_un_anno].notnull(), [primo_indicatore_un_anno, secondo_indicatore_un_anno, indicatore_corretto_un_anno]] = np.nan
        self.foglio.loc[(self.foglio['data_di_avvio'] < t0_3Y) & self.foglio['Best_Worst_3Y'].isnull(), 'note'] = 'Ha 3 anni, ma non è in classifica a tre anni.'
        # self.foglio.loc[(self.foglio['data_di_avvio'] > t0_3Y) & self.foglio[primo_indicatore_tre_anni].notnull(), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.' Nota fuorviante
        self.foglio.loc[(self.foglio['data_di_avvio'] > t0_3Y) & self.foglio[primo_indicatore_tre_anni].notnull(), [primo_indicatore_tre_anni, secondo_indicatore_tre_anni, indicatore_corretto_tre_anni]] = np.nan
        # Ranking finale
        if soluzioni[macro] == 1:
            # Fondi best blend - Gerarchia : semi_attivo, attivo, molto_attivo
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi best non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi worst non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        elif soluzioni[macro] == 2:
            # Fondi best blend - Gerarchia : attivo, semi_attivo, molto_attivo
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi best non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            # Fondi worst blend - Gerarchia : attivo, semi_attivo, molto_attivo
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'semi_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'semi_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi worst non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        elif soluzioni[macro] == 3:
            # Fondi best blend - Gerarchia : (semi_attivo & attivo), molto_attivo
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi best non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            # Fondi worst blend - Gerarchia : (semi_attivo & attivo), molto_attivo
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'] == 'molto_attivo'), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'].isin(['attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'] == 'molto_attivo'), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi worst non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        elif soluzioni[macro] == 4:
            # Fondi best blend - Gerarchia : semi_attivo & attivo & molto_attivo
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True)
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi best non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'best') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] != 'best') & (self.foglio['Best_Worst_1Y'] == 'best') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            # Fondi worst blend - Gerarchia : semi_attivo & attivo & molto_attivo
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_3Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_3Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] == classi_metodo_doppio[macro]) &  (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio['grado_gestione_1Y'].isin(['molto_attivo', 'attivo', 'semi_attivo'])), 'ranking_per_grado_1Y'].rank(method='first', na_option='bottom', ascending=True) + ultimo_elemento_ordinato
            # Fondi worst non blend
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'] == 'worst') & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
            ultimo_elemento_ordinato = self.foglio['ranking_finale'].max()
            if math.isnan(ultimo_elemento_ordinato): ultimo_elemento_ordinato = 0
            self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale'] = self.foglio.loc[(self.foglio['micro_categoria'] != classi_metodo_doppio[macro]) & (self.foglio['Best_Worst_3Y'].isnull()) & (self.foglio['Best_Worst_1Y'] == 'worst') & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno].rank(method='first', na_option='bottom', ascending=False) + ultimo_elemento_ordinato
        
        # Cambio formato data
        self.foglio['data_di_avvio'] = self.foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
        # Ordinamento finale
        self.foglio.sort_values('ranking_finale', ascending=True, inplace=True)
        # Reindex
        self.foglio.reset_index(drop=True, inplace=True)
        # Seleziona colonne utili
        self.foglio = self.foglio[
            ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Best_Worst_3Y', 'grado_gestione_3Y', 
            'Best_Worst_1Y', 'grado_gestione_1Y', 'ranking_per_grado_3Y', 'ranking_per_grado_1Y', 'ranking_finale',
            primo_indicatore_tre_anni, secondo_indicatore_tre_anni, 'commissione', indicatore_corretto_tre_anni,
            primo_indicatore_un_anno, secondo_indicatore_un_anno, 'commissione', indicatore_corretto_un_anno, 'SFDR', 'note']
        ]

        return self.foglio

    def linearizzazione(self, indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione):
        if indicatore == 'IR':
            primo_indicatore_tre_anni = 'Information_Ratio_3Y'
            secondo_indicatore_tre_anni = 'TEV_3Y'
            primo_indicatore_un_anno = 'Information_Ratio_1Y'
            secondo_indicatore_un_anno = 'TEV_1Y'
            indicatore_corretto_tre_anni = 'IR_corretto_3Y'
            indicatore_corretto_un_anno  = 'IR_corretto_1Y'
            # Creazione IR_corretto_3Y e IR_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            # Codice vecchio
            # # Creazione IR_corretto_1Y
            # foglio['IR_corretto_1Y'] = ((df['Information_Ratio_1Y'] * (df['TEV_1Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)) / (df['TEV_1Y'] / 100)
            # # Creazione IR_corretto_3Y
            # foglio['IR_corretto_3Y'] = ((df['Information_Ratio_3Y'] * (df['TEV_3Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)) / (df['TEV_3Y'] / 100)
            # # Note
            # foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Information_Ratio_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
            # foglio.loc[(foglio['data_di_avvio'] > self.t0_3Y) & (foglio['Information_Ratio_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            # # Ranking finale
            # minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'])
            # massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'])
            # foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'ranking_finale_3Y'
            # ] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['IR_corretto_3Y'].notnull()), 'IR_corretto_3Y'
            #     ] / (massimo_3Y - minimo_3Y)
            # minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'])
            # massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'])
            # foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'ranking_finale_1Y'
            # ] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['IR_corretto_1Y'].notnull()), 'IR_corretto_1Y'
            #     ] / (massimo_1Y - minimo_1Y)
            # foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
            # foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')
        elif indicatore == 'SO':
            primo_indicatore_tre_anni = 'Sortino_3Y'
            secondo_indicatore_tre_anni = 'DSR_3Y'
            primo_indicatore_un_anno = 'Sortino_1Y'
            secondo_indicatore_un_anno = 'DSR_1Y'
            indicatore_corretto_tre_anni = 'SO_corretto_3Y'
            indicatore_corretto_un_anno  = 'SO_corretto_1Y'
            # Creazione SO_corretto_3Y e SO_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            # Codice vecchio
            # # Creazione SO_corretto_1Y
            # foglio['SO_corretto_1Y'] = (
            #     (df['Sortino_1Y'] * (df['DSR_1Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)
            # ) / (df['DSR_1Y'] / 100)
            # # Creazione SO_corretto_3Y
            # foglio['SO_corretto_3Y'] = (
            #     (df['Sortino_3Y'] * (df['DSR_3Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)
            # ) / (df['DSR_3Y'] / 100)
            # # Note
            # foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sortino_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
            # foglio.loc[(foglio['data_di_avvio'] > self.t0_3Y) & (foglio['Sortino_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            # # Ranking finale
            # minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'])
            # massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'])
            # foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'ranking_finale_3Y'
            # ] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SO_corretto_3Y'].notnull()), 'SO_corretto_3Y'
            #     ] / (massimo_3Y - minimo_3Y)
            # minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'])
            # massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'])
            # foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'ranking_finale_1Y'
            # ] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SO_corretto_1Y'].notnull()), 'SO_corretto_1Y'
            #     ] / (massimo_1Y - minimo_1Y)
            # foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
            # foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')
        elif indicatore == 'SH':
            primo_indicatore_tre_anni = 'Sharpe_3Y'
            secondo_indicatore_tre_anni = 'Vol_3Y'
            primo_indicatore_un_anno = 'Sharpe_1Y'
            secondo_indicatore_un_anno = 'Vol_1Y'
            indicatore_corretto_tre_anni = 'SH_corretto_3Y'
            indicatore_corretto_un_anno  = 'SH_corretto_1Y'
            # Creazione SO_corretto_3Y e SO_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            # Codice vecchio
            # # Creazione SH_corretto_1Y
            # foglio['SH_corretto_1Y'] = ((df['Sharpe_1Y'] * (df['Vol_1Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)) / (df['Vol_1Y'] / 100)
            # # Creazione SH_corretto_3Y
            # foglio['SH_corretto_3Y'] = ((df['Sharpe_3Y'] * (df['Vol_3Y'] / 100) ) - (df['commissione'] / self.anni_detenzione)) / (df['Vol_3Y'] / 100)
            # # Note
            # foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Sharpe_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
            # foglio.loc[(foglio['data_di_avvio'] > self.t0_3Y) & (foglio['Sharpe_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            # # Ranking finale
            # minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'])
            # massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'])
            # foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'ranking_finale_3Y'
            # ] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_3Y) & (foglio['SH_corretto_3Y'].notnull()), 'SH_corretto_3Y'
            #     ] / (massimo_3Y - minimo_3Y)
            # minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'])
            # massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'])
            # foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'ranking_finale_1Y'
            # ] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[
            #     (foglio['data_di_avvio'] < self.t0_1Y) & (foglio['SH_corretto_1Y'].notnull()), 'SH_corretto_1Y'
            #     ] / (massimo_1Y - minimo_1Y)
            # foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
            # foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')
        elif indicatore == 'PERF':
            primo_indicatore_tre_anni = 'Perf_3Y'
            secondo_indicatore_tre_anni = 'Vol_3Y'
            primo_indicatore_un_anno = 'Perf_1Y'
            secondo_indicatore_un_anno = 'Vol_1Y'
            indicatore_corretto_tre_anni = 'PERF_corretto_3Y'
            indicatore_corretto_un_anno  = 'PERF_corretto_1Y'
            # Creazione PERF_corretto_3Y e PERF_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
            else:
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / anni_detenzione)
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / anni_detenzione)
            # Codice vecchio
            # # Creazione PERF_corretto_1Y
            # foglio['PERF_corretto_1Y'] = (df['Perf_1Y'] / 100) - (df['commissione'] / self.anni_detenzione)
            # # Creazione PERF_corretto_3Y
            # foglio['PERF_corretto_3Y'] = (df['Perf_3Y'] / 100) - (df['commissione'] / self.anni_detenzione)
            # # Note
            # foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['Perf_3Y'].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
            # foglio.loc[(foglio['data_di_avvio'] > self.t0_3Y) & (foglio['Perf_3Y'].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
            # # Ranking finale
            # minimo_3Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'])
            # massimo_3Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'])
            # foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'ranking_finale_3Y'] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < self.t0_3Y) & (foglio['PERF_corretto_3Y'].notnull()), 'PERF_corretto_3Y'] / (massimo_3Y - minimo_3Y)
            # minimo_1Y = min(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'])
            # massimo_1Y = max(foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'])
            # foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'ranking_finale_1Y'] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * foglio.loc[(foglio['data_di_avvio'] < self.t0_1Y) & (foglio['PERF_corretto_1Y'].notnull()), 'PERF_corretto_1Y'] / (massimo_1Y - minimo_1Y)
            # foglio['ranking_finale'] = foglio['ranking_finale_3Y'].fillna(foglio['ranking_finale_1Y'])
            # foglio['podio'] = foglio['ranking_finale'].apply(lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')
        # Note
        self.foglio.loc[(self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio[primo_indicatore_tre_anni].isnull()), 'note'] = 'Ha 3 anni, ma non possiede dati a tre anni.'
        self.foglio.loc[(self.foglio['data_di_avvio'] > t0_3Y) & (self.foglio[primo_indicatore_tre_anni].notnull()), 'note'] = 'Non ha 3 anni, ma possiede dati a tre anni.'
        # Ranking finale
        minimo_3Y = min(self.foglio.loc[(self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni])
        massimo_3Y = max(self.foglio.loc[(self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni])
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio[indicatore_corretto_tre_anni].notnull()), 'ranking_finale_3Y'
        ] = 1 - 8 * minimo_3Y / (massimo_3Y - minimo_3Y) + 8 * self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_3Y) & (self.foglio[indicatore_corretto_tre_anni].notnull()), indicatore_corretto_tre_anni
            ] / (massimo_3Y - minimo_3Y)
        minimo_1Y = min(self.foglio.loc[(self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno])
        massimo_1Y = max(self.foglio.loc[(self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno])
        self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio[indicatore_corretto_un_anno].notnull()), 'ranking_finale_1Y'
        ] = 1 - 8 * minimo_1Y / (massimo_1Y - minimo_1Y) + 8 * self.foglio.loc[
            (self.foglio['data_di_avvio'] < t0_1Y) & (self.foglio[indicatore_corretto_un_anno].notnull()), indicatore_corretto_un_anno
            ] / (massimo_1Y - minimo_1Y)
        self.foglio['ranking_finale'] = self.foglio['ranking_finale_3Y'].fillna(self.foglio['ranking_finale_1Y'])
        self.foglio['podio'] = self.foglio['ranking_finale'].apply(
            lambda ranking: 'bronzo' if ranking <= 3.0 else 'argento' if ranking <= 6.0 else 'oro' if ranking <= 9.1 else '')
        
        # Cambio formato data
        self.foglio['data_di_avvio'] = self.foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
        # Ordinamento finale
        self.foglio.sort_values('ranking_finale', ascending=False, inplace=True)
        # Etichetta ND per i fondi senza dati
        self.foglio['ranking_finale_1Y'] = self.foglio['ranking_finale_1Y'].fillna('ND')
        self.foglio['ranking_finale_3Y'] = self.foglio['ranking_finale_3Y'].fillna('ND')
        self.foglio['ranking_finale'] = self.foglio['ranking_finale'].fillna('ND')
        # Reindex
        self.foglio.reset_index(drop=True, inplace=True)
        # Seleziona colonne utili
        self.foglio = self.foglio[
            ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'podio', 'ranking_finale', 'ranking_finale_3Y',
            'ranking_finale_1Y', primo_indicatore_tre_anni, secondo_indicatore_tre_anni, 'commissione', indicatore_corretto_tre_anni,
            primo_indicatore_un_anno, secondo_indicatore_un_anno, 'commissione', indicatore_corretto_un_anno, 'note']
        ]

        return self.foglio

    def doppio_con_linearizzazione(self, macro, indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione, soluzioni, classi_metodo_doppio):
        if indicatore == 'IR':
            primo_indicatore_tre_anni = 'Information_Ratio_3Y'
            secondo_indicatore_tre_anni = 'TEV_3Y'
            primo_indicatore_un_anno = 'Information_Ratio_1Y'
            secondo_indicatore_un_anno = 'TEV_1Y'
            indicatore_corretto_tre_anni = 'IR_corretto_3Y'
            indicatore_corretto_un_anno  = 'IR_corretto_1Y'
            # Creazione IR_corretto_3Y e IR_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            else:
                self.foglio[indicatore_corretto_tre_anni] = (
                    (self.foglio[primo_indicatore_tre_anni] * (self.foglio[secondo_indicatore_tre_anni] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_tre_anni] / 100)
                self.foglio[indicatore_corretto_un_anno] = (
                    (self.foglio[primo_indicatore_un_anno] * (self.foglio[secondo_indicatore_un_anno] / 100)) - 
                    (self.foglio['commissione'] / anni_detenzione)
                ) / (self.foglio[secondo_indicatore_un_anno] / 100)
            # Chiama il metodo doppio
            self.foglio = self.doppio(macro, indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione, soluzioni, classi_metodo_doppio)
            # Rinomina la colonna
            self.foglio.rename(columns = {'ranking_finale': 'classifica_finale'}, inplace = True)
        elif indicatore == 'PERF':
            primo_indicatore_tre_anni = 'Perf_3Y'
            secondo_indicatore_tre_anni = 'Vol_3Y'
            primo_indicatore_un_anno = 'Perf_1Y'
            secondo_indicatore_un_anno = 'Vol_1Y'
            indicatore_corretto_tre_anni = 'PERF_corretto_3Y'
            indicatore_corretto_un_anno  = 'PERF_corretto_1Y'
            # Creazione PERF_corretto_3Y e PERF_corretto_1Y
            if intermediario == 'RAI': # Raiffeisen specifica gli anni di detenzione per fondo, non in maniera universale.
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / self.foglio['anni_detenzione'])
            else:
                self.foglio[indicatore_corretto_tre_anni] = (self.foglio[primo_indicatore_tre_anni] / 100) - (self.foglio['commissione'] / anni_detenzione)
                self.foglio[indicatore_corretto_un_anno] = (self.foglio[primo_indicatore_un_anno] / 100) - (self.foglio['commissione'] / anni_detenzione)
            # Chiama il metodo doppio
            self.foglio = self.doppio(macro, indicatore, t0_1Y, t0_3Y, intermediario, anni_detenzione, soluzioni, classi_metodo_doppio)
            # Rinomina la colonna
            self.foglio.rename(columns = {'ranking_finale': 'classifica_finale'}, inplace = True)

        # Cambio formato data
        self.foglio['data_di_avvio'] = pd.to_datetime(self.foglio['data_di_avvio'], dayfirst=True)
        # Ranking finale
        minimo_3Y = min(self.foglio.loc[self.foglio['classifica_finale'].notnull(), 'classifica_finale'])
        massimo_3Y = max(self.foglio.loc[self.foglio['classifica_finale'].notnull(), 'classifica_finale'])
        self.foglio.loc[
            self.foglio['classifica_finale'].notnull(), 'ranking_finale'
        ] = 1 - 8 / (massimo_3Y - minimo_3Y) + 8 * (
            massimo_3Y - 
            self.foglio.loc[
                self.foglio['classifica_finale'].notnull(), 'classifica_finale'
            ] + 
            minimo_3Y
        ) / (massimo_3Y - minimo_3Y)   
        # # Ordinamento finale
        # self.foglio.sort_values('ranking_finale', ascending=False, inplace=True)
        # # Etichetta ND per i fondi senza dati
        # self.foglio['ranking_finale_1Y'] = self.foglio['ranking_finale_1Y'].fillna('ND')
        # self.foglio['ranking_finale_3Y'] = self.foglio['ranking_finale_3Y'].fillna('ND')
        # self.foglio['ranking_finale'] = self.foglio['ranking_finale'].fillna('ND')
        # Reindex
        # self.foglio.reset_index(drop=True, inplace=True)
        # Cambio formato data
        self.foglio['data_di_avvio'] = self.foglio['data_di_avvio'].dt.strftime('%d/%m/%Y')
        # Seleziona colonne utili
        self.foglio = self.foglio[
            ['ISIN', 'valuta', 'nome', 'data_di_avvio', 'micro_categoria', 'Best_Worst_3Y', 'grado_gestione_3Y', 
            'Best_Worst_1Y', 'grado_gestione_1Y', 'ranking_per_grado_3Y', 'ranking_per_grado_1Y', 'classifica_finale',
            'ranking_finale', primo_indicatore_tre_anni,
            secondo_indicatore_tre_anni, 'commissione', indicatore_corretto_tre_anni,
            primo_indicatore_un_anno, secondo_indicatore_un_anno, 'commissione', indicatore_corretto_un_anno, 'note']
        ]

        return self.foglio

t1 = '30/09/2023'
macro = 'AZ_EUR'
t0_3Y = np.datetime64(
        datetime.datetime.strptime(t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+3)
    )
t0_1Y = np.datetime64(
        datetime.datetime.strptime(t1, '%d/%m/%Y') - dateutil.relativedelta.relativedelta(days=-1, years=+1)
    )
intermediario = 'CRV'
anni_detenzione = 3
soluzioni = {
        'LIQ' : 4, 'OBB_EUR_BT' : 4, 'OBB_EUR_MLT' : 4, 'OBB_EUR_CORP' : 4, 'OBB_GLOB' : 4, 'OBB_EM' : 4,
        'OBB_HY' : 4, 'AZ_EUR' : 4, 'AZ_NA' : 4, 'AZ_PAC' : 4, 'AZ_EM' : 4, 'AZ_GLOB' : 4,
    }
classi_metodo_doppio = {
        'LIQ' : 'Monetari Euro', 'OBB_EUR_BT' : 'Obblig. Euro breve term.', 'OBB_EUR_MLT' : 'Obblig. Euro all maturities', 
        'OBB_EUR_CORP' : 'Obblig. Euro corporate', 'OBB_GLOB' : 'Obblig. globale', 'OBB_EM' : 'Obblig. Paesi Emerg.',
        'OBB_HY' : 'Obblig. globale high yield', 'AZ_EUR' : 'Az. Europa', 'AZ_NA' : 'Az. USA', 'AZ_PAC' : 'Az. Pacifico',
        'AZ_EM' : 'Az. paesi emerg. Mondo', 'AZ_GLOB' : 'Az. globale', 
    }

if __name__ == '__main__':
    df = pd.read_excel('ranking.xlsx', index_col=None)
    df['data_di_avvio'] = pd.to_datetime(df['data_di_avvio'], dayfirst=True)
    foglio = df.loc[df['macro_categoria'] == 'AZ_EUR']
    a = Metodi_ranking(foglio)
    b = a.doppio(macro, t0_1Y, t0_3Y, intermediario, anni_detenzione, soluzioni, classi_metodo_doppio)
    print(b)