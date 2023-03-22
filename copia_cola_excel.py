import openpyxl

def copiar_colar_excel(origem, destino, planilha_origem):
    # Abrir arquivo de origem
    wb_origem = openpyxl.load_workbook(origem, data_only=True)
    ws_origem = wb_origem[planilha_origem]

    # Selecionar o range de células a ser copiado
    range_copia = ws_origem['A23':'D49']

    # Abrir arquivo de destino
    wb_destino = openpyxl.Workbook()
    ws_destino = wb_destino.active

    # Copiar e colar como valor
    for row in range_copia:
        nova_linha = []
        for cell in row:
            nova_linha.append(cell.value)
        ws_destino.append(nova_linha)

    # Salvar arquivo de destino
    wb_destino.save(destino)


diretorio_arquivo_origem=r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\TEND_FIBRA_e_UPDATEs.xlsx'
diretorio_arquivo_destino=r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\tend_email.xlsx'
planilha_origem='UPDATE_TENDENCIA_VLL'
planilha_destino='20230320'

#range_celulas='A23:D45'
copiar_colar_excel(diretorio_arquivo_origem, diretorio_arquivo_destino, planilha_origem)