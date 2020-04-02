import PySimpleGUI as sg
from mbox_processor import proccess_mbox
import platform

sg.ChangeLookAndFeel('BrownBlue')
sg.SetOptions(element_padding=(0, 0))

LAYOUT_SIZES = {
    'column': 30,
    'progress': 37,
    'extract': 15,
    'mark_all': 13,
    'unmark_all': 13,
    'close': 6
}

if platform.system() == 'Windows':
    LAYOUT_SIZES = {
        'column': 26,
        'progress': 38,
        'extract': 16,
        'mark_all': 14,
        'unmark_all': 14,
        'close': 6
    }

layout = [
    [sg.Text('Archivo MBOX')],
    [sg.InputText(disabled=True, size=(46, 1)),
     sg.FileBrowse('Seleccionar', file_types=(('MBOX', '*.mbox'),), key='mboxfile', size=(12, 1))],
    [sg.Column([
        [sg.Text('Cabecera', size=(LAYOUT_SIZES['column'], 1), justification='center', text_color=sg.BLUES[0])],
        [sg.Checkbox('Incluir Fecha', key='date')],
        [sg.Checkbox('Incluir Destinatario', key='to')],
        [sg.Checkbox('Solo Email', key='to_email_only', pad=(20, 0))],
        [sg.Checkbox('Incluir Remitente', key='from')],
        [sg.Checkbox('Solo Email', key='from_email_only', pad=(20, 0))],
        [sg.Checkbox('Incluir Con Copia', key='cc')],
        [sg.Checkbox('Incluir Asunto', key='subject')],
    ]), sg.VerticalSeparator(), sg.Column([
        [sg.Text('Contenido', size=(LAYOUT_SIZES['column'], 1), justification='center', text_color=sg.BLUES[0])],
        [sg.Checkbox('Extraec Datos de Contacto', key='contact_data')],
        [sg.Checkbox('Incluir Cuerpo', key='body')],
        [sg.Checkbox('Incluir Adjuntos', key='attachment')],
    ])],
    [sg.ProgressBar(max_value=100, size=(LAYOUT_SIZES['progress'], 20), pad=(4, 4), key='progress')],
    [sg.Button('Extraer', button_color=(sg.BLUES[0], 'green'), size=(LAYOUT_SIZES['extract'], None)),
     sg.Button('Marcar Todo', size=(LAYOUT_SIZES['mark_all'], None)),
     sg.Button('Desmarcar Todo', size=(LAYOUT_SIZES['unmark_all'], None)),
     sg.Button('Cerrar', button_color=(sg.BLUES[0], 'red'), size=(LAYOUT_SIZES['close'], None))]
]

window = sg.Window('MBOX Extractor', layout, resizable=False)


def solo_mail_state(key):
    element = window[f'{key}_email_only']
    if values[key]:
        element.Update(disabled=False)
    else:
        element.Update(disabled=True, value=False)


def set_check(value):
    for key in ['date', 'to', 'from', 'cc', 'subject', 'body', 'contact_data', 'attachment']:
        window[key].Update(value=value)


def disable_all(value):
    for key in ['date', 'to', 'from', 'cc', 'subject', 'body', 'contact_data', 'attachment', 'mboxfile', 'Extraer',
                'Marcar Todo', 'Desmarcar Todo', 'Cerrar']:
        window[key].Update(disabled=value)


window.Read(timeout=0)
set_check(True)

while True:
    event, values = window.Read(timeout=10)
    if event == 'Cerrar' or event is None:
        break
    if event == 'Marcar Todo':
        set_check(True)
    if event == 'Desmarcar Todo':
        set_check(False)
    if event == 'Extraer':
        disable_all(True)
        try:
            proccess_mbox(values['mboxfile'],
                          values['date'],
                          values['to'],
                          values['to_email_only'],
                          values['from'],
                          values['from_email_only'],
                          values['cc'],
                          values['subject'],
                          values['body'],
                          values['contact_data'],
                          values['attachment'],
                          lambda curr, total: window['progress'].UpdateBar(curr, total))
        except Exception as e:
            sg.popup_ok(f'{e}', title='Error')
        disable_all(False)

    solo_mail_state('from')
    solo_mail_state('to')

window.close()
