#!/usr/bin/env python

import csv
import tkinter.ttk as tk
from datetime import date
from os.path import basename
from os.path import join as joinpath
from os.path import splitext, dirname, basename
from tkinter import messagebox, Tk
from tkinter.filedialog import askopenfilename

from docxtpl import DocxTemplate

CURRENT_DATE_KEY = 'current_date'


def main():
    root = Tk()
    root.title('Form Filler')
    # root.geometry('500x200')

    # Client details selection
    details_path_entry = build_file_selection_row(
        root, 0, 'Client details file:')

    # Template selection
    template_path_entry = build_file_selection_row(
        root, 1, 'Template file:')

    def submit():
        params_path = details_path_entry.get()
        template_path = template_path_entry.get()
        try:
            fill_forms(template_path, params_path)
        except ValueError as e:
            messagebox.showerror('Error', e.message)
        except Exception as e:
            messagebox.showerror('Error', 'Unknown error: {}'.format(e))

    submit_button = tk.Button(root, text='Fill form', command=submit)
    submit_button.grid(column=0, row=3)

    root.mainloop()


def build_file_selection_row(root, row, label_text):
    tk.Label(root, text=label_text).grid(row=row, column=0)
    entry = tk.Entry(root)
    entry.grid(row=row, column=1)

    def browse_file():
        filename = askopenfilename()
        entry.delete(0, 'end')
        entry.insert(0, filename)

    button = tk.Button(root, text='Browse', command=browse_file)
    button.grid(row=row, column=2)

    return entry


def fill_forms(template_path, params_path):
    print('Template path: {}'.format(template_path))
    print('Params path: {}'.format(params_path))

    params = get_params(params_path)
    print('Params: {}'.format(params))

    doc = DocxTemplate(template_path)
    doc.render(params)
    doc.save(get_output_path(params_path, template_path, 'docx'))
    print('Filled form successfully!')


def get_params(params_path):
    with open(params_path, encoding='utf8') as fp:
        params = [row[:2] for row in csv.reader(fp)]
    params.append([CURRENT_DATE_KEY, date.today().strftime('%d/%m/%Y')])
    return params


def get_output_path(params_path, template_path, file_extension):
    client_folder = dirname(params_path)
    client_name = splitext(basename(params_path))[0]
    template_name = splitext(basename(template_path))[0]
    file_name = '{}_{}.{}'.format(client_name, template_name, file_extension)
    return joinpath(client_folder, file_name)


if __name__ == '__main__':
    main()
