#!/usr/bin/env python
# -*- coding: utf-8 -*-
from tkinter.ttk import Label, Entry, Button, Frame
from os import makedirs
from os.path import basename, dirname, expanduser
from os.path import join as joinpath
from os.path import splitext
from tkinter import Image, Tk, messagebox
from tkinter.filedialog import askopenfilename

from formfiller import fill_form, read_params

PARAMS_FILE_TYPES = (('Excel file', '.xlsx'),)
TEMPLATE_FILE_TYPES = (('Word file', '*.docx'), ('PDF file', '*.pdf'))
ICON_PATH = 'icon.png'
APP_DIR = '.formfiller'


def generate_output_path(params, template_path):
    folder = dirname(params)
    params_name = splitext(basename(params))[0]
    template_name = basename(template_path)
    file_name = '{}_{}'.format(params_name, template_name)
    return joinpath(folder, file_name)


def load_setting(setting_name):
    if setting_name is None:
        return ''
    try:
        with open(joinpath(expanduser('~'), APP_DIR, setting_name), 'r') as f:
            return f.read()
    except FileNotFoundError:
        return ''


def save_setting(setting_name, value):
    dir_path = joinpath(expanduser('~'), APP_DIR)
    print(dir_path)
    makedirs(dir_path, exist_ok=True)
    with open(joinpath(dir_path, setting_name), 'w') as f:
        f.write(value)


class FormFillerApp(object):
    def __init__(self):
        self.root = Tk()
        self.root.title('Form Filler')
        self.root.resizable(True, False)
        self.root.minsize(width=400, height=100)
        self.root.columnconfigure(1, weight=1)

        # Set window icon
        img = Image("photo", file=ICON_PATH)
        self.root.tk.call('wm', 'iconphoto', self.root._w, img)
        # root.iconphoto(True, img)
        # root.iconbitmap(ICON_PATH)

        self.frame = Frame(self.root)
        self.frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.details_path_entry = self.build_file_selection_row(
            0, 'Client', PARAMS_FILE_TYPES)

        self.template_path_entry = self.build_file_selection_row(
            1, 'Form', TEMPLATE_FILE_TYPES)

        self.result_label = Label(self.frame, text='')
        self.result_label.grid(
            column=1, columnspan=2, row=3, pady=(6, 0), padx=10, sticky='wns')

        self.submit_button = Button(
            self.frame, text='Fill form', command=self.submit)
        self.submit_button.grid(column=0, row=3, pady=(6, 0), sticky='wns')

    def build_file_selection_row(self, row, label_text, filetypes):
        setting_name = 'entry_{}'.format(label_text.lower().replace(' ', '_'))
        label = Label(self.frame, text=label_text)
        label.grid(row=row, column=0, pady=(0, 2), sticky='w')

        entry = Entry(self.frame, width=45)
        entry.grid(row=row, column=1, padx=8, pady=(0, 2), sticky='we')
        entry.insert(0, load_setting(setting_name))

        def browse_file():
            filename = askopenfilename(filetypes=filetypes)
            self.result_label.config(text='')
            save_setting(setting_name, filename)
            entry.delete(0, 'end')
            entry.insert(0, filename)
            entry.xview("end")

        button = Button(self.frame, text='Browse', command=browse_file)
        button.grid(row=row, column=2, pady=(0, 2), sticky='e')

        return entry

    def submit(self):
        params_path = self.details_path_entry.get()
        template_path = self.template_path_entry.get()
        print('Params: {}'.format(params_path))
        print('Template: {}'.format(template_path))

        params = read_params(params_path)
        print('Params: {}'.format(params))

        output_path = generate_output_path(params_path, template_path)

        try:
            fill_form(template_path, params_path, output_path)
            self.result_label.config(
                text='Form saved to {}'.format(output_path))
            print('Form filled successfully!')
            print('Output: {}'.format(output_path))
        except Exception as e:
            messagebox.showerror('Error', str(e))
            self.result_label.config(text='')
            raise e


if __name__ == '__main__':
    FormFillerApp().root.mainloop()
