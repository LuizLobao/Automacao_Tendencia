import PySimpleGUI as sg

"""
    Demo - Element List
    All elements shown in 1 window as simply as possible.
    Copyright 2022 PySimpleGUI
"""


use_custom_titlebar = True if sg.running_trinket() else False

def make_window(theme=None):

    NAME_SIZE = 59


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
        [sg.Checkbox('Copiar Arquivo',disabled=False, default=1),sg.Checkbox('Gerar Tab_Din',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(47,20))],
    ]
    frame_processo_VL_VLL_Legado = [
        [sg.Text('Processo VL/VLL Legado')],
        [sg.Checkbox('Executa Procedure',disabled=False, default=1),sg.Checkbox('Monta Excel',disabled=False, default=1),sg.Checkbox('Envia Email',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(47,20))],
    ]
    frame_iguala_tendencias = [
        [sg.Text('Processo Igualar Tendencia = Real')],
        [sg.Checkbox('Iguala X',disabled=False, default=1),sg.Checkbox('Iguala Y',disabled=False, default=1),sg.Checkbox('Iguala Z',disabled=False, default=1), sg.Button('iniciar', disabled=True)],
        [sg.ProgressBar(100, orientation='h', s=(47,20))],
    ]
    layout_esquerda = [
                       [sg.Text('Rotinas para Tendência')],
                       [name('Verificar Disponibilidade de Arquivos'), sg.Button('atualizar', k='-MON_CARGA-')],
                       [sg.Table([['BOV 6163','22/10/2022 09:26','22/10/2022 09:36','22/10/2022 10:26'], 
                                  ['BOV 6162','22/10/2022 09:26','22/10/2022 09:36','22/10/2022 10:26'],
                                  ['BOV 9012','22/10/2022 09:26','',''],
                                  ['BOV 3456','22/10/2022 09:26','',''],
                                  ['BOV 7890','22/10/2022 09:26','',''],
                                  ['BOV 9876','22/10/2022 09:26','',''],
                                  ['Dem. Gross','X','Y','Z']], ['Arquivo','Download','Inicio Carga','Fim Carga'], num_rows=7, justification='l')],
                        [sg.Text('')],
                        [name('Processo VLL Nova Fibra (6162 e 6163)'), sg.Button('iniciar', disabled=True)],
                        [sg.Output(s=(40,3))],
                        #[name('Processo Demonstrativo Gross'), sg.Button('iniciar', disabled=True)],
                        #[sg.Output(s=(40,3))],
                        [sg.Frame('Demontrativo de Gross', frame_demonstrativo_gross)],
                        #[name('Processo VL/VLL Legado'), sg.Button('iniciar', disabled=True)],
                        [sg.Frame('VL/VLL Legado', frame_processo_VL_VLL_Legado)],
                        #[sg.Output(s=(40,3)), sg.Button('abrir excel', disabled=True)],
                        [sg.Frame('Tendência = Real', frame_iguala_tendencias)],

                        #[name('Iguala Real e Tendência'), sg.Button('iniciar', disabled=True)],
                        ]


    layout_direita = []

    layout_l = [
                [name('Text'), sg.Text('Text')],
                [name('Input'), sg.Input(s=15)],
                [name('Multiline'), sg.Multiline(s=(15,2))],
                [name('Output'), sg.Output(s=(15,2))],
                [name('Combo'), sg.Combo(sg.theme_list(), default_value=sg.theme(), s=(15,22), enable_events=True, readonly=True, k='-COMBO-')],
                [name('OptionMenu'), sg.OptionMenu(['OptionMenu',],s=(15,2))],
                [name('Checkbox'), sg.Checkbox('Checkbox')],
                [name('Radio'), sg.Radio('Radio', 1)],
                [name('Spin'), sg.Spin(['Spin',], s=(15,2))],
                [name('Button'), sg.Button('Button')],
                [name('ButtonMenu'), sg.ButtonMenu('ButtonMenu', sg.MENU_RIGHT_CLICK_EDITME_EXIT)],
                [name('Slider'), sg.Slider((0,10), orientation='h', s=(10,15))],
                [name('Listbox'), sg.Listbox(['Listbox', 'Listbox 2'], no_scrollbar=True,  s=(15,2))],
                [name('Image'), sg.Image(sg.EMOJI_BASE64_HAPPY_THUMBS_UP)],
                [name('Graph'), sg.Graph((125, 50), (0,0), (125,50), k='-GRAPH-')]  ]

    layout_r  = [[name('Canvas'), sg.Canvas(background_color=sg.theme_button_color()[1], size=(125,40))],
                [name('ProgressBar'), sg.ProgressBar(100, orientation='h', s=(10,20), k='-PBAR-')],
                [name('Table'), sg.Table([[1,2,3], [4,5,6]], ['Col 1','Col 2','Col 3'], num_rows=2)],
                [name('Tree'), sg.Tree(treedata, ['Heading',], num_rows=3)],
                [name('Horizontal Separator'), sg.HSep()],
                [name('Vertical Separator'), sg.VSep()],
                [name('Frame'), sg.Frame('Frame', [[sg.T(s=15)]])],
                [name('Column'), sg.Column([[sg.T(s=15)]])],
                [name('Tab, TabGroup'), sg.TabGroup([[sg.Tab('Tab1',[[sg.T(s=(15,2))]]), sg.Tab('Tab2', [[]])]])],
                [name('Pane'), sg.Pane([sg.Col([[sg.T('Pane 1')]]), sg.Col([[sg.T('Pane 2')]])])],
                [name('Push'), sg.Push(), sg.T('Pushed over')],
                [name('VPush'), sg.VPush()],
                [name('Sizer'), sg.Sizer(1,1)],
                [name('StatusBar'), sg.StatusBar('StatusBar')],
                [name('Sizegrip'), sg.Sizegrip()]  ]

    # Note - LOCAL Menu element is used (see about for how that's defined)
    layout = [[Menu([['Arquivo', ['Sair']], ['Editar', ['Edit Me', ]],['Sobre']],  k='-CUST MENUBAR-',p=0)],
              [sg.T('Sistema de Automações de Rotinas - Atividades Lobão', font='_ 14', justification='c', expand_x=True)],
              [sg.T('Atividades de Rotinas Diárias', font='_ 14', justification='c', expand_x=True)],
              #[sg.Checkbox('Use Custom Titlebar & Menubar', use_custom_titlebar, enable_events=True, k='-USE CUSTOM TITLEBAR-', p=0)],
              [sg.Col(layout_esquerda, p=0),sg.Col(layout_l, p=0), sg.Col(layout_r, p=0)]]

    window = sg.Window('Sistema de Automações de Rotinas - Luiz Lobão', layout, finalize=True, right_click_menu=sg.MENU_RIGHT_CLICK_EDITME_VER_EXIT, keep_on_top=True, use_custom_titlebar=use_custom_titlebar)

    window['-PBAR-'].update(30)                                                     # Show 30% complete on ProgressBar
    window['-GRAPH-'].draw_image(data=sg.EMOJI_BASE64_HAPPY_JOY, location=(0,50))   # Draw something in the Graph Element

    return window


window = make_window()

while True:
    event, values = window.read()
    # sg.Print(event, values)
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if values['-COMBO-'] != sg.theme():
        sg.theme(values['-COMBO-'])
        window.close()
        window = make_window()
    if event == '-USE CUSTOM TITLEBAR-':
        use_custom_titlebar = values['-USE CUSTOM TITLEBAR-']
        sg.set_options(use_custom_titlebar=use_custom_titlebar)
        window.close()
        window = make_window()
    if event == 'Edit Me':
        sg.execute_editor(__file__)
    elif event == 'Version':
        sg.popup_scrolled(__file__, sg.get_versions(), keep_on_top=True, non_blocking=True)
window.close()