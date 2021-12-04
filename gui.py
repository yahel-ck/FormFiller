#!/usr/bin/env python
# -*- coding: utf-8 -*-
import tkinter.ttk as tk
from tkinter import messagebox, Tk
from tkinter.filedialog import askopenfilename
from formfiller import fill_form


def main():
    root = Tk()
    root.title('Form Filler')
    root.resizable(True, False)
    root.minsize(width=400, height=100)
    # root.geometry('500x200')

    # Client details selection
    details_path_entry = build_file_selection_row(
        root, 0, 'Client', (('CSV files', '*.csv'),), pady=(8, 2))

    # Template selection
    template_path_entry = build_file_selection_row(
        root, 1, 'Form', (('Word files', '*.docx'),))

    def submit():
        params_path = details_path_entry.get()
        template_path = template_path_entry.get()
        try:
            fill_form(template_path, params_path)
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


if __name__ == '__main__':
    main()
