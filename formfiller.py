#!/usr/bin/env python
# -*- coding: utf-8 -*-
from datetime import date
from docxtpl import DocxTemplate
from pdfjinja import PDFTemplate
import pandas as pd

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

    df = pd.read_excel(params_path)
    params = [(row[0], row[1]) for row in df.values]

    return append_default_params(params)


def append_default_params(params):
    """
    Appends the default parameters to the given parameters if the their key isn't defined.
    """
    if TODAY_KEY not in map(lambda pair: pair[0], params):
        params.append((TODAY_KEY, date.today().strftime('%d/%m/%Y')))
    return params
