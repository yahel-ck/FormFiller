#!/usr/bin/env python
# -*- coding: utf-8 -*-
import csv
from datetime import date
from io import TextIOWrapper
from os.path import basename, dirname
from os.path import join as joinpath
from os.path import splitext

from docxtpl import DocxTemplate

CURRENT_DATE_KEY = 'current_date'


def fill_form(template_path, params_path, output_path=None):
    print('Template path: {}'.format(template_path))
    print('Params path: {}'.format(params_path))

    params = get_params(params_path)
    print('Params: {}'.format(params))

    if output_path is None:
        output_path = get_output_path(params_path, template_path, 'docx')

    doc = DocxTemplate(template_path)
    doc.render(params)
    doc.save(output_path)

    print('Filled form successfully!')
    print('Output path: {}'.format(output_path))


def get_params(params_path):
    with open(params_path, 'rb') as bf:
        with TextIOWrapper(bf, encoding='utf-8', newline='') as f:
            params = [row[:2] for row in csv.reader(f)]
    params.append([CURRENT_DATE_KEY, date.today().strftime('%d/%m/%Y')])
    return params


def get_output_path(params_path, template_path, file_extension):
    client_folder = dirname(params_path)
    client_name = splitext(basename(params_path))[0]
    template_name = splitext(basename(template_path))[0]
    file_name = '{}_{}.{}'.format(client_name, template_name, file_extension)
    return joinpath(client_folder, file_name)
