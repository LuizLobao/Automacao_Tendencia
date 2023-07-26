import PySimpleGUI as sg

use_custom_titlebar = True if sg.running_trinket() else False

def make_window(theme=None):

    NAME_SIZE = 10


    def name(name):
        dots = NAME_SIZE-len(name)-2
        return sg.Text(name + ' ' + '_'*dots, size=(NAME_SIZE,1), justification='l',pad=(0,0), font='Courier 10')

    sg.theme(theme)

    # NOTE that we're using our own LOCAL Menu element
    if use_custom_titlebar:
        Menu = sg.MenubarCustom
    else:
        Menu = sg.Menu

    treedata = sg.TreeData()

    treedata.Insert("", '_A_', 'Tree Item 1', [1234], )
    treedata.Insert("", '_B_', 'B', [])
    treedata.Insert("_A_", '_A1_', 'Sub Item 1', ['can', 'be', 'anything'], )

    frame_demonstrativo_gross = [
        [sg.Text('Processo Demonstrativo Gross')],
        [sg.Checkbox('Copiar Arquivo',disabled=False, default=1),sg.Checkbox('Gerar Tab_Din',disabled=False, default=1), sg.Button('iniciar', disabled=False,k='btn_dem_gross')],
        [sg.ProgressBar(100, orientation='h', s=(30,20), key='progressbar_demonstrativo')],
    ]
    frame_processo_VL_VLL_Legado = [
        [sg.Text('Processo VL/VLL Legado')],
        [sg.Checkbox('Executa Procedure',disabled=False, default=1),sg.Checkbox('Monta Excel',disabled=False, default=1),sg.Checkbox('Envia Email',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(30,20), key='progressbar_vll')],
    ]
    frame_iguala_tendencias = [
        [sg.Text('Processo Igualar Tendencia = Real')],
        [sg.Checkbox('Iguala X',disabled=False, default=1),sg.Checkbox('Iguala Y',disabled=False, default=1),sg.Checkbox('Iguala Z',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(30,20), key='progressbar_iguala')],
    ]
    frame_monitorcarga = [
        [name('Verificar Disponibilidade de Arquivos'), sg.Button('atualizar', k='-MON_CARGA-')],
        [sg.ProgressBar(100, orientation='h', s=(47,20))],
        [sg.Table([['BOV 6163','22/10/2022 09:26','22/10/2022 09:36','22/10/2022 10:26'], 
                ['BOV 6162','22/10/2022 09:26','22/10/2022 09:36','22/10/2022 10:26'],
                ['BOV 9012','22/10/2022 09:26','',''],
                ['BOV 3456','22/10/2022 09:26','',''],
                ['BOV 7890','22/10/2022 09:26','',''],
                ['BOV 9876','22/10/2022 09:26','',''],
                ['Dem. Gross','X','Y','Z']], ['Arquivo','Download','Inicio Carga','Fim Carga'], num_rows=7, justification='l')]
    ]
    frame_receita_contratada = [
        [sg.Text('Processos para atualização da Receita Contratada')],
        [sg.Checkbox('Ajusta Ticket Var',disabled=False, default=1),sg.Checkbox('Ajusta Ticket Emp',disabled=False, default=1),sg.Checkbox('Roda procedure X',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(30,20), key='progressbar_receitacontratada')],
    ]
    frame_tendencia_liberada = [
        [sg.Text('Envia E-mail de Tendências Liberadas')],
        [sg.Checkbox('Envia X',disabled=False, default=1),sg.Checkbox('Envia Y',disabled=False, default=1),sg.Checkbox('Envia Z',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(30,20), key='progressbar_enviaemailtendliberada')],
    ]
    frame_pdv_outros = [
        [sg.Text('Envia E-mail de Tendências Liberadas')],
        [sg.Checkbox('Envia X',disabled=False, default=1),sg.Checkbox('Envia Y',disabled=False, default=1),sg.Checkbox('Envia Z',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(30,20), key='progressbar_enviaemailtendliberada')],
    ]
    layout_esquerda = [
                        [sg.Frame('Verifica Monitor de Carga', frame_monitorcarga, size=(550,230),expand_y=True)],
                        [sg.Frame('Demontrativo de Gross', frame_demonstrativo_gross, size=(550,100))],
                        [sg.Frame('VL/VLL Legado', frame_processo_VL_VLL_Legado, size=(550,100))],
                        [sg.Frame('Tendência = Real', frame_iguala_tendencias, size=(550,100))],
    ]
    layout_direita = [
                        [sg.Frame('Receita Contratada', frame_receita_contratada, size=(550,100))],
                        [sg.Frame('E-mail Tendências Liberadas', frame_tendencia_liberada, size=(550,100))],
                        [sg.Frame('Envia Lista de PDV Outros', frame_pdv_outros, size=(550,100))],
    ]
                      
    layout = [[Menu([['Arquivo', ['Sair']], ['Editar', ['Edit Me', ]],['Sobre']],  k='-CUST MENUBAR-',p=0)],
              [sg.T('Sistema de Automações de Rotinas - Atividades Lobão', font='_ 14', justification='c', expand_x=True)],
              [sg.T('Atividades de Rotinas Diárias', font='_ 14', justification='c', expand_x=True)],
              [sg.Col(layout_esquerda, p=0),sg.Col(layout_direita, p=0,vertical_alignment='top')]
              ]

    window = sg.Window('Sistema de Automações de Rotinas - Luiz Lobão', layout, finalize=True, right_click_menu=sg.MENU_RIGHT_CLICK_EDITME_VER_EXIT, keep_on_top=False, use_custom_titlebar=use_custom_titlebar)

    
    window['progressbar_demonstrativo'].update(30)
    window['progressbar_vll'].update(50)
    window['progressbar_iguala'].update(80)

    return window

window = make_window()

valor = 1
while True:
    event, values = window.read()
    # sg.Print(event, values)
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == '-USE CUSTOM TITLEBAR-':
        use_custom_titlebar = values['-USE CUSTOM TITLEBAR-']
        sg.set_options(use_custom_titlebar=use_custom_titlebar)
        window.close()
        window = make_window()
    if event == 'Edit Me':
        sg.execute_editor(__file__)
    if event == 'btn_dem_gross':
        valor = valor + 1 
        window['progressbar_demonstrativo'].update(valor)
    elif event == 'Version':
        sg.popup_scrolled(__file__, sg.get_versions(), keep_on_top=True, non_blocking=True)
    
window.close()