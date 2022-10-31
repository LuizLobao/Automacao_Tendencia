#VERIFICAR DATA DE ATUALIZAÇÃO DO ARQUIVO
#SE DATA DO ARQUIVO = HOJE, COPIAR PARA A PASTA DE REDE
#CRIAR TABELAS DINAMICAS PARA INSTALAÇÃO E PARA MIGRAÇÃO

import shutil,os,time
from telnetlib import theNULL
from datetime import date, datetime
import pandas as pd
from openpyxl import load_workbook

hoje = datetime.today().strftime('%Y-%m-%d')
AAAAMMDD = datetime.today().strftime('%Y%m%d')
arquivo1 = 'Demonstrativo Gross'


# Verificar data de Modificação antes de Realizar a Cópia do Arquivo
def data_mod_arquivo():
    arquivo = (f'Y:\{arquivo1}.xlsb')
    modificado = time.strftime('%Y-%m-%d', time.gmtime(os.path.getmtime(arquivo)))
    return (modificado)
# Realiza a copia do Arquivo para a pasta de Donwload local renomeando o arquivo
def copia_arquivo_renomeia():
    shutil.copy(r"Y:\\Demonstrativo Gross.xlsb", fr'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsb')
    
def monta_tabdin_demonstrativo_gross():
    excel_file = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsb'
    dest_filename = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsx'

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


datamod = data_mod_arquivo()
tentativas = 1

while (datamod != hoje):
    print('\x1b[1;32;41m' + 'Datas Diferentes. Esperando 5 minutos' + '\x1b[0m')
    time.sleep(300) #5minutos
    tentativas = tentativas + 1
    print(f'Verificando novamente o arquivo na rede...Tentativa número: {tentativas}')
    datamod = data_mod_arquivo()
    print(datamod)

if datamod == hoje:
    print('Iniciando processo de copiar arquivo e criar tabelas dinâmicas')
    copia_arquivo_renomeia()
    monta_tabdin_demonstrativo_gross()
else: 
    print('ESCAPE - NÃO COPIAR')