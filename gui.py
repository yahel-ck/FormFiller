#!/usr/bin/env python
# -*- coding: utf-8 -*-
import tkinter.ttk as tk
from os.path import basename, dirname
from os.path import join as joinpath
from os.path import splitext
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename

from formfiller import fill_form, read_params

PARAMS_FILE_TYPES = (('CSV file', '*.csv'), ('JSON file', '*.json'))
TEMPLATE_FILE_TYPES = (('Word file', '*.docx'), ('PDF file', '*.pdf'))


def main():
    root = Tk()
    root.title('Form Filler')
    root.resizable(True, False)
    root.minsize(width=400, height=100)
    # root.geometry('500x200')

    # Client details selection
    details_path_entry = build_file_selection_row(
        root, 0, 'Client', PARAMS_FILE_TYPES, pady=(8, 2))

    # Template selection
    template_path_entry = build_file_selection_row(
        root, 1, 'Form', TEMPLATE_FILE_TYPES)

    def submit():
        params_path = details_path_entry.get()
        template_path = template_path_entry.get()
        print('Params: {}'.format(params_path))
        print('Template: {}'.format(template_path))
        params = read_params(params_path)
        print('Params: {}'.format(params))
        output_path = generate_output_path(
            params_path, template_path, splitext(template_path)[1][1:])

        try:
            fill_form(template_path, params_path, output_path)
            print('Form filled successfully!')
            print('Output: {}'.format(output_path))
        except Exception as e:
            if isinstance(e, ValueError) and 'not a Word file' in str(e.args):
                msg = e.args[0]
            else:
                msg = 'Unknown error: {}'.format(e)
            messagebox.showerror('Error', msg)
            raise e

    submit_button = tk.Button(root, text='Fill form', command=submit)
    submit_button.grid(column=0, columnspan=2, row=3,
                       pady=(4, 6), padx=8, sticky='w')

    root.columnconfigure(1, weight=1)
    root.mainloop()


def build_file_selection_row(root, row, label_text, filetypes, pady=0):
    label = tk.Label(root, text=label_text)
    label.grid(row=row, column=0, padx=(10, 10), pady=pady, sticky='w')

    entry = tk.Entry(root)
    entry.grid(row=row, column=1, padx=8, pady=pady, sticky='we')

    def browse_file():
        filename = askopenfilename(filetypes=filetypes)
        entry.delete(0, 'end')
        entry.insert(0, filename)

    button = tk.Button(root, text='Browse', command=browse_file)
    button.grid(row=row, column=2, padx=(0, 8), pady=pady, sticky='e')

    return entry


def generate_output_path(params, template_path, file_extension):
    folder = dirname(params)
    params_name = splitext(basename(params))[0]
    template_name = splitext(basename(template_path))[0]
    file_name = '{}_{}.{}'.format(params_name, template_name, file_extension)
    return joinpath(folder, file_name)


if __name__ == '__main__':
    main()
