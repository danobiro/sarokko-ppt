import PySimpleGUI as sg

import gui_tools

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
    return [[sg.Input(key=(key,count),enable_events=True, size=(45,1))]]

def get_suggestions(text,options):
    """Returns up to 4 matching options from the list."""
    return [opt for opt in options if text.lower() in opt.lower()][:4]

def get_uniq_chars(event):
    if type(event) is tuple:
        return event[0].strip('-')[-2:]
    else:
        return event.strip('-')[-2:]

def update_suggestions(event):
    user_input = values[event]
    matches = get_suggestions(user_input, songs)
    
    uniq_chars = get_uniq_chars(event)

    if matches:
        window[f'-SUGG{uniq_chars}-'].update(values=matches, visible=True)
        #window.TKroot.geometry("")  # Reset window size to fit contents
        #window.refresh()  # Force window to reflow elements properly
    else:
        window[f'-SUGG{uniq_chars}-'].update(values=[], visible=False)
        #window.TKroot.geometry("")  # Reset window size to fit contents
        #window['-SUGGESTIONS-'].hide_row()  # Hide the empty space

def select_suggestion(event,key):
    uniq_chars = get_uniq_chars(event)

    selected = values[event][0]
    window[(f'-INP{uniq_chars}-',key)].update(selected)
    window[event].update(values=[], visible=False)
    #window.TKroot.geometry("")  # Reset window size to fit contents
    #window['-SUGGESTIONS-'].hide_row()  # Hide the empty space

songs = gui_tools.get_songs_list()
    
adv_options = [[sg.Text("Szöveg skálázási faktora")],
               [sg.Input("98")]]

ie_column = [[sg.Input(key=('-INPIE-',0),enable_events=True, size=(45,1)), sg.Button('+1', k='-BTNIE-')]]
ig_column = [[sg.Input(key=('-INPIG-',0),enable_events=True, size=(45,1)), sg.Button('+1', k='-BTNIG-')]]
iu_column = [[sg.Input(key=('-INPIU-',0),enable_events=True, size=(45,1)), sg.Button('+1', k='-BTNIU-')]]
tu_column = [[sg.Input(key=('-INPTU-',0),enable_events=True, size=(45,1)), sg.Button('+1', k='-BTNTU-')]]

layout = [[sg.Text('Előző prezentáció')],
          [sg.Input(), sg.FileBrowse()],
          [sg.Text('Igevers előtti énekek')],
          [sg.Column(ie_column, k='-IECOL-')],
          [sg.pin(sg.Listbox(values=[], key='-SUGGIE-', size=(45, 4), enable_events=True, no_scrollbar=True, visible=False))],
          [sg.Text('Énekek közti igeversek (formátum: Tit 3,4-7)')],
          [sg.Column(ig_column, k='-IGCOL-')],
          [sg.Text('Igevers utáni énekek')],
          [sg.Column(iu_column, k='-IUCOL-')],
          [sg.pin(sg.Listbox(values=[], key='-SUGGIU-', size=(45, 4), enable_events=True, no_scrollbar=True, visible=False))],
          [sg.Text('Tanítás utáni énekek (pl Úrvacsorakor)')],
          [sg.Column(tu_column, k='-TUCOL-')],
          [sg.pin(sg.Listbox(values=[], key='-SUGGTU-', size=(45, 4), enable_events=True, no_scrollbar=True, visible=False))],
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

key = 0
while True:             # Event Loop
    event, values = window.read()
    # Hide advanced menu by default
    print(event, values)
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if type(event) is tuple:
        key = event[1]
        #if event[0] == '-INPIE-':
        update_suggestions(event)

    else:
        #if event == '-SUGGIE-':
        if '-SUGG' in event:
            select_suggestion(event,key)

        if event.startswith('-OPEN ADV-'):
            adv_opened = not adv_opened
            window['-OPEN ADV-'].update(SYMBOL_DOWN if adv_opened else SYMBOL_UP)
            window['-ADV-'].update(visible=adv_opened)

        elif event == '-BTNIE-':
            if ie_num < max_num:
                ie_num += 1
                window.extend_layout(window['-IECOL-'],create_input_column('-INPIE-',ie_num-1))

        elif event == '-BTNIG-':
            pass
            #if ig_num < max_num:
            #    ig_num += 1
            #    window.extend_layout(window['-IGCOL-'],create_input_column('-INPIG-',ig_num-1))

        elif event == '-BTNIU-':
            if iu_num < max_num:
                iu_num += 1
                window.extend_layout(window['-IUCOL-'],create_input_column('-INPIU-',iu_num-1))

        elif event == '-BTNTU-':
            if tu_num < max_num:
                tu_num += 1
                window.extend_layout(window['-TUCOL-'],create_input_column('-INPTU-',tu_num-1))

window.close()