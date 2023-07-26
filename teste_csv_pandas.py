import pandas as pd

arquivo = 'pdv_canal_vl_junho23_1a18_rcs.csv'

df = pd.read_csv(arquivo, delimiter=';')

print(df.head())