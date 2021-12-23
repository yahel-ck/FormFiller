#!/usr/bin/env python
# -*- coding: utf-8 -*-
from datetime import date
from jinja2.environment import Environment

import pandas as pd
from docxtpl import DocxTemplate
from jinja2 import Undefined

from pdfjinja import PDFTemplate

TODAY_KEY = 'today'


def fill_form(template_path, params, output_path, jinja_env=None, check_missing_vars=True):
    """
    Fills the form with the given parameters and saves it to the given output 
    path. Uses `get_doc_template` to determine the appropriate template class.
    """
    params = read_params(params)
    doc = get_doc_template(template_path)

    if jinja_env is None:
        jinja_env = Environment()

    if check_missing_vars:
        jinja_env.undefined = CollectUndefined(jinja_env.undefined)
        doc.render(params, jinja_env=jinja_env)
        jinja_env.undefined.assert_no_missing_vars()
    else:
        doc.render(params, jinja_env=jinja_env)

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
    Appends the default parameters to the given parameters if the their key 
    isn't defined.
    """
    if TODAY_KEY not in map(lambda pair: pair[0], params):
        params.append((TODAY_KEY, date.today().strftime('%d/%m/%Y')))
    return params


class CollectUndefined(object):
    def __init__(self, undefined_cls=Undefined):
        self.undefined_cls = undefined_cls
        self.missing_vars = []

    def __call__(self, *args, **kwds):
        undefined = self.undefined_cls(*args, **kwds)
        self.missing_vars.append(undefined._undefined_name)
        return undefined

    def assert_no_missing_vars(self):
        if len(self.missing_vars) > 0:
            raise MissingVariablesError(self.missing_vars)


class MissingVariablesError(Exception):
    def __init__(self, missing_vars, *args):
        super().__init__(*args)
        self.missing_vars = missing_vars

    def __str__(self):
        return 'Missing variables: {}'.format(self.missing_vars)
