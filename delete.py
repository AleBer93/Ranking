import pandas as pd
import numpy as np
df = pd.read_csv(r'C:\Users\Alessio\Documents\Sbwkrq\Ranking\docs\export_liste_complete_from_Q\lista_completa_0.csv', sep = ';', 
    decimal=',', engine='python', encoding = "utf_16_le", skipfooter=1)
micro = df['Categoria Quantalys'].unique()
print(micro)