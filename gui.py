import PySimpleGUI as sg

sg.theme('DarkAmber')

# Code is from PySimpleGUI.org, with modifications
SYMBOL_UP =    '▲'
SYMBOL_DOWN =  '▼'

def collapse(layout, key, visible=True):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this seciton visible / invisible
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    return sg.pin(sg.Column(layout, key=key, visible=visible))

def create_input_column(key,count):
    return [[sg.Input(key=(key,count))]]
    
adv_options = [[sg.Text("Szöveg skálázási faktora")],
               [sg.Input("98")]]

ie_column = [[sg.Input(key=('-INPIE-',0)), sg.Button('+1', k='-BTNIE-')]]
ig_column = [[sg.Input(key=('-INPIG-',0)), sg.Button('+1', k='-BTNIG-')]]
iu_column = [[sg.Input(key=('-INPIU-',0)), sg.Button('+1', k='-BTNIU-')]]
tu_column = [[sg.Input(key=('-INPTU-',0)), sg.Button('+1', k='-BTNTU-')]]

layout = [[sg.Text('Előző prezentáció')],
          [sg.Input(), sg.FileBrowse()],
          [sg.Text('Igevers előtti énekek')],
          [sg.Column(ie_column, k='-IECOL-')],
          [sg.Text('Énekek közti igeversek (formátum: Tit 3,4-7)')],
          [sg.Column(ig_column, k='-IGCOL-')],
          [sg.Text('Igevers utáni énekek')],
          [sg.Column(iu_column, k='-IUCOL-')],
          [sg.Text('Tanítás utáni énekek (pl Úrvacsorakor)')],
          [sg.Column(tu_column, k='-TUCOL-')],
          [sg.Text('Előző prezentációban az első hírdetéses slide sorszáma')],
          [sg.Input(key='-INPSL-')],
          ### Haladó beállítások
          [sg.T(SYMBOL_UP, enable_events=True, k='-OPEN ADV-'), sg.T('Haladó beállítások', enable_events=True, k='-OPEN ADV-TEXT')],
          [collapse(adv_options, '-ADV-',visible=False)],
          [sg.Button('Start',expand_x=True)]
]

window = sg.Window('Sarokkő pptx gererátor', layout)

adv_opened = False

max_num = 5
ie_num = 1
ig_num = 1
iu_num = 1
tu_num = 1

while True:             # Event Loop
    event, values = window.read()
    # Hide advanced menu by default
    print(event, values)
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event.startswith('-OPEN ADV-'):
        adv_opened = not adv_opened
        window['-OPEN ADV-'].update(SYMBOL_DOWN if adv_opened else SYMBOL_UP)
        window['-ADV-'].update(visible=adv_opened)

    if event == '-BTNIE-':
        if ie_num < max_num:
            ie_num += 1
            window.extend_layout(window['-IECOL-'],create_input_column(event,ie_num))

    if event == '-BTNIG-':
        if ig_num < max_num:
            ig_num += 1
            window.extend_layout(window['-IGCOL-'],create_input_column(event,ig_num))

    if event == '-BTNIU-':
        if iu_num < max_num:
            iu_num += 1
            window.extend_layout(window['-IUCOL-'],create_input_column(event,iu_num))

    if event == '-BTNTU-':
        if tu_num < max_num:
            tu_num += 1
            window.extend_layout(window['-TUCOL-'],create_input_column(event,tu_num))

window.close()