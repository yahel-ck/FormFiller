#!/usr/bin/env python
# -*- coding: utf-8 -*-
import csv
from datetime import date
from io import TextIOWrapper
from docxtpl import DocxTemplate
from pdfjinja import PDFTemplate

TODAY_KEY = 'today'


def fill_form(template_path, params, output_path):
    """
    Fills the form with the given parameters and saves it to the given output path.
    Uses `get_doc_template` to determine the appropriate template class.
    """
    params = read_params(params)
    doc = get_doc_template(template_path)
    doc.render(params)
    doc.save(output_path)


def get_doc_template(template_path):
    """
    Returns the appropriate template class for the given template path
    according to the file extention.
    """
    if template_path.endswith('.docx'):
        return DocxTemplate(template_path)
    elif template_path.endswith('.pdf'):
        return PDFTemplate(template_path)


def read_params(params_path):
    if isinstance(params_path, dict) or isinstance(params_path, list):
        # params already read
        return params_path

    with open(params_path, 'rb') as bf:
        with TextIOWrapper(bf, encoding='utf-8', newline='') as f:
            params = [row[:2] for row in csv.reader(f)]

    if not next((pair[0] == TODAY_KEY for pair in params), None):
        params.append([TODAY_KEY, date.today().strftime('%d/%m/%Y')])
    return params

