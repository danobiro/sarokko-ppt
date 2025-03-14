import PySimpleGUI as sg

import gui_tools
import sarokko_ppt_generator

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

def create_input_column(key,count,length=45):
    return [[sg.Input(key=(key,count),enable_events=True, size=(length,1))]]

def create_ig_column():
    return [[sg.Input(key=('-INPIGE-',i),enable_events=True, size=(10,1),visible=False) for i in range(4)]]


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

def read_song_input(vals,event_key,row_num=5):
    r_list = []

    for i in range(row_num):
        curr_song = vals[(event_key,i)]
        if curr_song:
            r_list.append(curr_song)

    return r_list
    
songs = gui_tools.get_songs_list()
    
adv_options = [[sg.Text("Szöveg skálázási faktora")],
               [sg.Input("98")]]

ie_column = [[sg.Input(key=('-INPIE-',0),enable_events=True, size=(45,1)), sg.Button('+1', k='-BTNIE-')]]

ig_list = [sg.Input(key=('-INPIG-',i),enable_events=True, size=(10,1)) for i in range(4)] 
ig_list.append(sg.Button('+1', k='-BTNIG-'))
ig_column = [[sg.Column([ig_list])]]

iu_column = [[sg.Input(key=('-INPIU-',0),enable_events=True, size=(45,1)), sg.Button('+1', k='-BTNIU-')]]
tu_column = [[sg.Input(key=('-INPTU-',0),enable_events=True, size=(45,1)), sg.Button('+1', k='-BTNTU-')]]

layout = [[sg.Text('Előző prezentáció')],
          [sg.Input(k='-INPLOC-'), sg.FileBrowse()],
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
          [sg.Input(key='-INPSLSTRT-',size=(4,1)), sg.Text('És az utolsóé:'), sg.Input(key='-INPSLEND-', size=(4,1))],
          ### Haladó beállítások
          [sg.T(SYMBOL_UP, enable_events=True, k='-OPEN ADV-'), sg.T('Haladó beállítások', enable_events=True, k='-OPEN ADV-TEXT')],
          [collapse(adv_options, '-ADV-',visible=False)],
          [sg.Button('Start',expand_x=True)]
]

window = sg.Window('Sarokkő pptx gererátor', layout)

adv_opened = False

max_num = 5
ie_num = 1
iu_num = 1
tu_num = 1
ig_num = 0

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
        if '-INPIG' not in event[0]:
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
                window.extend_layout(window['-IECOL-'],create_input_column('-INPIE-',ie_num))
                ie_num += 1

        elif event == '-BTNIG-':
            max_num_ig = 4
            if ig_num < max_num_ig:
                ig_num += 1

                #if ig_num == 2:
                    # extended fields
                    #window.extend_layout(window['-IGCOL-'],create_input_column('-INPIGE-',ig_num-2,length=10))
                window.extend_layout(window['-IGCOL-'],create_ig_column())
                window[('-INPIGE-',ig_num-1)].update(visible=True)


        elif event == '-BTNIU-':
            if iu_num < max_num:
                window.extend_layout(window['-IUCOL-'],create_input_column('-INPIU-',iu_num))
                iu_num += 1

        elif event == '-BTNTU-':
            if tu_num < max_num:
                window.extend_layout(window['-TUCOL-'],create_input_column('-INPTU-',tu_num))
                tu_num += 1

        elif event == 'Start':
            passed_tests = True

            rvals = {}

            # Get previous slides location
            rvals['prev_loc'] = values['-INPLOC-']

            # Get songs before Bible verses
            rvals['ie_songs'] = read_song_input(values,'-INPIE-',ie_num)

            # Get songs after BV
            rvals['iu_songs'] = read_song_input(values,'-INPIU-',iu_num)

            # Get songs after teaching
            rvals['tu_songs'] = read_song_input(values,'-INPTU-',tu_num)

            # Get Bible verses
            rvals['verses'] = []
            for i in range(4):
                curr_str = values[('-INPIG-',i)]
                if curr_str:
                    rvals['verses'].append(curr_str)
            for i in range(4):
                try:
                    curr_str = values[('-INPIGE-',i)]
                    if curr_str:
                        rvals['verses'].append(curr_str)
                except:
                    pass

            # Get previous slide details
            try:
                rvals['last_slide_start'] = int(values['-INPSLSTRT-'])
            except:
                error_msg = "Hiba: Az előző prezentáció kezdő sorszáma nem megfelelő"
                passed_tests = False
            try:
                rvals['last_slide_end'] = int(values['-INPSLEND-'])
            except:
                rvals['last_slide_end'] = -1

            if passed_tests:
                result = gui_tools.validate_input_data(rvals)
                passed_tests = result[0]
                error_msg = result[1]
            
            if not passed_tests:
                sg.Popup(error_msg , title='Hiba', keep_on_top=True)

            else:
                sarokko_ppt_generator.run(rvals)

window.close()