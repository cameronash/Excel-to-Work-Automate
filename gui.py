import PySimpleGUI as sg

# --- Layout definition ---
layout = [
    [sg.Text('Excel file:'), 
     sg.Input(key='-EXCEL-', enable_events=True), 
     sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
    
    [sg.Text('Word file:'), 
     sg.Input(key='-WORD-', enable_events=True), 
     sg.FileBrowse(file_types=(("Word Docs", "*.doc;*.docx;*.docm"),))],
    
    [sg.Button('Run'), sg.Button('Exit')],
    
    # Optional output area to show logs or progress
    [sg.Output(size=(80, 10))]
]

# --- Window creation ---
window = sg.Window('G-Slide Runner', layout)

# --- Event loop ---
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    elif event == 'Run':
        excel_path = values['-EXCEL-']
        word_path  = values['-WORD-']
        print(f'🔹 Selected Excel: {excel_path}')
        print(f'🔹 Selected Word : {word_path}')
        # TODO: call your main() or run_value_into_word logic here
        print('▶️  Running…')

window.close()
