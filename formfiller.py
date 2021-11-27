import csv
import tkinter
from datetime import date
from os.path import basename
from os.path import join as joinpath
from os.path import splitext
from tkinter import ttk
from tkinter.filedialog import askopenfilename

from docxtpl import DocxTemplate


CSV_PROG_ID_PATH = '.csv'
CSV_PROG_ID = 'CSV'
RIGHT_CLICK_OPTION_PATH = r'Software\Classes\CSV\shell\formfiller'
RIGHT_CLICK_TITLE = 'Fill Form'

CURRENT_DATE_KEY = 'current_date'


def main():
    root = tkinter.Tk()
    root.title('Form Filler')
    # root.geometry('500x200')

    # Client details selection
    details_path_entry = build_file_selection_row(
        root, 0, 'Client details file:')

    # Template selection
    template_path_entry = build_file_selection_row(root, 1, 'Template file:')

    def submit():
        params_path = details_path_entry.get()
        template_path = template_path_entry.get()
        fill_forms(template_path, params_path)

    submit_button = tkinter.Button(root, text='Fill form', command=submit)
    submit_button.grid(column=0, row=3)

    root.mainloop()


def build_file_selection_row(root, row, label_text):
    ttk.Label(root, text=label_text).grid(row=row, column=0)
    entry = ttk.Entry(root)
    entry.grid(row=row, column=1)

    def browse_file():
        filename = askopenfilename()
        entry.delete(0, 'end')
        entry.insert(0, filename)

    button = tkinter.Button(root, text='Browse', command=browse_file)
    button.grid(row=row, column=2)

    return entry


def fill_forms(template_path, params_path):
    print('Template path: {}'.format(template_path))
    print('Params path: {}'.format(params_path))

    params = get_params(params_path)
    print('Params: {}'.format(params))

    doc = DocxTemplate(template_path)
    doc.render(params)
    doc.save(".\{}.docx".format(
        generate_file_name(params_path, template_path)))

    print('Filled form successfully!')


def get_params(params_path):
    with open(params_path, encoding='utf8') as fp:
        params = [row[:2] for row in csv.reader(fp)]
    params.append([CURRENT_DATE_KEY, date.today().strftime('%d/%m/%Y')])
    return params


def generate_file_name(params_path, template_path):
    params_name = splitext(basename(params_path))[0]
    template_name = splitext(basename(template_path))[0]
    return '{}_{}'.format(template_name, params_name)


if __name__ == '__main__':
    main()
