from parse import build_record_from_document, build_record_from_string
from export import word_export, excel_export
import PySimpleGUI as sg
import subprocess as sp

VERSION = '1.0.1'
TITLE = f'NDIS Document Parser Application v{VERSION}'
INPUT_DOCUMENT_ROW = [
    sg.Text('Input Word Document:', size=(25, 1)),
    sg.In(size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(key='-INPUT FILEBROWSE-', file_types=(('Word Documents', '*.docx'),))
]
EXCEL_DOCUMENT_ROW = [
    sg.Text('Output Excel Document (Optional):', size=(25, 1)),
    sg.In(key='-OUTPUT EXCEL TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
OUTPUT_FOLDER_ROW = [
    sg.Text('Output Folder:', size=(25, 1)),
    sg.In(key='-OUTPUT FOLDER TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FolderBrowse()
]
MULTILINE = [
    sg.Multiline(
        'Import an input document to begin...',
        text_color='grey',
        key='-DATA MULTILINE-',
        size=(150, 30),
        pad=(0, 15),
        disabled=True
    )
]
EXPORT = [
    sg.Button('Export Data', key='-EXPORT BUTTON-', size=(10, 2), disabled=True)
]
COLUMN = [
    INPUT_DOCUMENT_ROW,
    EXCEL_DOCUMENT_ROW,
    OUTPUT_FOLDER_ROW,
    MULTILINE,
    EXPORT
]
LAYOUT = [
    [
        sg.Column(COLUMN, element_justification='center')
    ]
]


class ScrolledText(sg.tk.Text):
    def __init__(self, master=None, **kw):
        horizontal_scrollbar = True
        self.frame = sg.tk.Frame(master)
        self.vbar = sg.tk.Scrollbar(self.frame)
        self.vbar.pack(side=sg.tk.RIGHT, fill=sg.tk.Y)

        if horizontal_scrollbar:
            self.hbar = sg.tk.Scrollbar(self.frame, orient='horizontal')
            self.hbar.pack(side=sg.tk.BOTTOM, fill=sg.tk.X)

        kw.update({'yscrollcommand': self.vbar.set})

        if horizontal_scrollbar:
            kw.update({'xscrollcommand': self.hbar.set})

        sg.tk.Text.__init__(self, self.frame, **kw)
        self.pack(side=sg.tk.LEFT, fill=sg.tk.BOTH, expand=True)
        self.vbar['command'] = self.yview

        if horizontal_scrollbar:
            self.hbar['command'] = self.xview

        text_meths = vars(sg.tk.Text).keys()
        methods = (
                vars(sg.tk.Pack).keys() |
                vars(sg.tk.Grid).keys() |
                vars(sg.tk.Place).keys())
        methods = methods.difference(text_meths)

        for m in methods:
            if m[0] != '_' and m != 'config' and m != 'configure':
                setattr(self, m, getattr(self.frame, m))

    def __str__(self):
        return str(self.frame)


def handle_window():
    """Create the window and handle its events

    Returns:
        None

    """
    window = sg.Window(TITLE, LAYOUT)

    output_excel_text = window['-OUTPUT EXCEL TEXT-']
    output_folder_text = window['-OUTPUT FOLDER TEXT-']
    data_multiline = window['-DATA MULTILINE-']
    export_button = window['-EXPORT BUTTON-']

    # Event Loop
    ml_enabled = False
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        # Input Path was updated
        if event == 0:
            path = values['-INPUT FILEBROWSE-']
            if not path:
                continue

            record = build_record_from_document(values['-INPUT FILEBROWSE-'])
            data_multiline.update(value=str(record))

            if not ml_enabled:
                data_multiline.Widget.configure(wrap='none')
                data_multiline.update(disabled=False, text_color='black')
                export_button.update(disabled=False)
                ml_enabled = True
        # Clicked the 'Export Data' button
        elif event == '-EXPORT BUTTON-':
            output_folder_path = output_folder_text.get()
            record = build_record_from_string(values['-DATA MULTILINE-'])
            if not output_folder_path:
                sg.Popup('Please select an output folder an try again.',
                         title='Error')
                continue

            if record is None:
                sg.Popup('Invalid formatting. Re-import the document to reset it or fix '
                         'the formatting manually, and then try again.',
                         title='Error')
                continue

            output_excel_path = output_excel_text.get()
            if output_excel_path:
                excel_export(record, optional_xml_path=output_excel_path)
            else:
                excel_export(record, export_folder=output_folder_path)

            word_export(record, output_folder_path)

            output_folder_path = output_folder_path.replace('/', '\\')
            sp.Popen(f'explorer {output_folder_path}')

    window.close()


if __name__ == '__main__':
    # Add a horizontal scrollbar to multiline elements
    sg.tk.scrolledtext.ScrolledText = ScrolledText

    handle_window()
