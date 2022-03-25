# import pandas as pd
# # with os.add_dll_directory('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\_blpapi'):
# #     import blpapi
# from xbbg import blp
# df = pd.DataFrame({'fondi' : ['LU0306632687', 'LU0343755343']})
# print(df)
# fondi = ['LU0306632687', 'LU0343755343']
# articoli = blp.bdp(['/isin/' + fondo for fondo in fondi], flds="sfdr_classification")
# articoli.reset_index(inplace=True)
# articoli['isin_code'] = articoli['index'].str[6:]
# print(articoli) # i fondi senza valore non vengono mostrati
# df = pd.merge(left=df, right=articoli, how='left', left_on='fondi', right_on='isin_code').drop(['index', 'isin_code'], axis=1)
# df['sfdr_classification'] = df['sfdr_classification'].fillna(0)
# df['sfdr_classification'] = pd.to_numeric(df['sfdr_classification'], errors='coerce').astype(int)
# df['sfdr_classification'].replace(0, '', inplace=True)
# print(df)
# df.to_excel('lol.xlsx')
while True:
    _ = input('scrivi qualcosa ')
    print(_)
    if _ == 'ok':
        break

for i in range(3):
    print(i)