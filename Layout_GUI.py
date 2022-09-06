from importlib.util import spec_from_file_location
from turtle import end_fill
from PySimpleGUI import Window, Button, Text, Image, Column, VSeparator, Push, theme, popup, MenuBar
from datetime import date, datetime
import os,time, monitor_cargas


#TO-DO:
#incluir um POP-UP para abertura de chamado caso haja atraso nos arquivos (a partir das 10hs)

datas_fim, datas_down, datas_ini = monitor_cargas.puxa_dts_cargas()



hoje = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
login = os.getlogin() #Puxa o Login do usuário 


############################## MONTANDO A JANELA DE INTERFACE COM O USUÁRIO ##############################
theme('DarkBlue')

# ------ Menu Definition ------ #
menu_def = [
    ['&Arquivo', ['&Abrir', '&Salvar', '---', 'Propriedades', '&Sair'  ]],
    ['&Editar', ['Colar', ['Especial', 'Normal',], 'Desfazer'],],
    ['A&juda', '&Sobre...']
]

layout_esquerda = [
    [Text(f'Olá {login}, seja bem vindo:')], 
    [Text()], 
    #[Image(filename='fusion.png')],
]
layout_direita = [
    [Text(f'Data de modificação dos arquivos?')],
    [Text()],
    [Text(f'Monitor de Carga...'), Push(), Text('Download'),  Push(), Text('Inicio Carga'), Push(),Text('Fim Carga')],
    [Text(f'BOV 1066 - Legado:'),  Push(), Text(f'{datas_down[0]}'), Push(), Text(f'{datas_ini[0]}'),Push(), Text(f'{datas_fim[0]}')],
    [Text(f'BOV 1067 - Legado:'),  Push(), Text(f'{datas_down[1]}'), Push(), Text(f'{datas_ini[1]}'),Push(), Text(f'{datas_fim[1]}')],
    [Text(f'BOV 1064 - Legado:'),  Push(), Text(f'{datas_down[2]}'), Push(), Text(f'{datas_ini[2]}'),Push(), Text(f'{datas_fim[2]}')],
    [Text(f'BOV 1059 - Legado:'),  Push(), Text(f'{datas_down[3]}'), Push(), Text(f'{datas_ini[3]}'),Push(), Text(f'{datas_fim[3]}')],
    [Text(f'BOV 1065 - Legado:'),  Push(), Text(f'{datas_down[4]}'), Push(), Text(f'{datas_ini[4]}'),Push(), Text(f'{datas_fim[4]}')],
    [Text(f'BOV 1058 - Legado:'),  Push(), Text(f'{datas_down[5]}'), Push(), Text(f'{datas_ini[5]}'),Push(), Text(f'{datas_fim[5]}')],
    [Text(f'BOV 6163 - NF VL:'),   Push(), Text(f'{datas_down[6]}'), Push(), Text(f'{datas_ini[6]}'),Push(), Text(f'{datas_fim[6]}')],
    [Text(f'BOV 6162 - NF Gross:'),Push(), Text(f'{datas_down[7]}'), Push(), Text(f'{datas_ini[7]}'),Push(), Text(f'{datas_fim[7]}')],
    
    [Text()],
    [Text(f'Diretório de Rede...')],
    [Text(f'Demonstrativo Gross:'), Text(f'dd/mm/aaaa hh:mm:ss')],
    [Text(f'Sharepoint...')],
    [Text(f'Anteneiros:'), Text(f'dd/mm/aaaa hh:mm:ss')],
    [Text()],
    [Button('Atualizar', key = '-ATUALIZARDATAS-')]
]
layout_centro = [
    [Text(f'Executar Processo Nova Fibra:')],
    [Button('Rodar Nova Fibra', key = '-RODARNOVAFIBRA-')],
    [Text()],
    [Text(f'Executar Processo Fibra LEGADO:')], 
    [Button('Rodar Fibra Legado', key = '-RODARFIBRALEGADO-')],
    [Text()],
    [Text(f'Igualar Tendências:')], 
    [Button('Nova Fibra', key = '-IGUALANF-'),Button('Fibra Varejo', key = '-IGUALAFIBRAVAR-'),Button('Tabelas Fibra', key = '-IGUALATBLFIBRA-'),Button('Fibra Empresarial', key = '-IGUALAFIBRAEMP-')],
    [Text()],
    [Text(f'Prepara Updates:')], 
    [Button('btn1'), Push(), Button('btn2'), Push(), Button('btn3')]
]
layout = [
    [MenuBar(menu_def), Column(layout_esquerda),VSeparator(), Column(layout_direita),VSeparator(), Column(layout_centro)]
]



window = Window(
    f'Sistema de Automação de Tendências - {hoje}',
    #size=(500,300),
    layout=layout
)
#print(window.read())

while True:
    event,values = window.read()

    match(event):
        case '-COPIAR-':
            #copia_arquivo_renomeia(),
            popup('Arquivo Copiado')
        case '-PUXARARQUIVO-':
            #puxa_conteudo_online()
            popup('Arquivo PUXADO')
        case None:
            break

window.close()

