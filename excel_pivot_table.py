import pandas as pd
from openpyxl import load_workbook
from datetime import date, datetime

hoje = datetime.today().strftime('%Y%m%d')

def monta_tabdin_demonstrativo_gross(nome_arquivo)
    excel_file = f'C:\\Users\\PC\\Documents\\Python\\Automação_Tendência\\{nome_arquivo}.xlsb'
    dest_filename = f'C:\\Users\\PC\\Documents\\Python\\Automação_Tendência\\Demonstrativo Gross_{hoje}.xlsx'

    df = pd.read_excel(excel_file, sheet_name='Database', engine='pyxlsb')
    pt_instalacao = df.query('TIPO == "INSTALACAO"').pivot_table(
                                                        values="PROJ", 
                                                        index=["UF"], 
                                                        columns="MERCADO", 
                                                        aggfunc=sum,
                                                        fill_value=0,
                                                        margins=True, margins_name="INSTALACAO",
                                                        )
    pt_migracao = df.query('TIPO == "MIGRACAO"').pivot_table(
                                                        values=["PROD","PROJ"], 
                                                        #index=["UF"], 
                                                        index="MERCADO", 
                                                        aggfunc=sum,
                                                        fill_value=0,
                                                        margins=True, margins_name="MIGRACAO",
                                                        )
    with pd.ExcelWriter(dest_filename) as writer:
        pt_instalacao.to_excel(writer, sheet_name="TabDin",startcol=0, startrow=0)
        pt_migracao.to_excel(writer, sheet_name="TabDin",startcol=6, startrow=0)

monta_tabdin_demonstrativo_gross('Demonstrativo Gross_20220922')        