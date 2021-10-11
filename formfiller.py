import sys
from docxtpl import DocxTemplate
import csv
from os.path import basename, splitext, join as joinpath
import winreg
from subprocess import check_call
from datetime import date


CSV_PROG_ID_PATH = '.csv'
CSV_PROG_ID = 'CSV'
RIGHT_CLICK_OPTION_PATH = r'Software\Classes\CSV\shell\formfiller'
RIGHT_CLICK_TITLE = 'Fill Form'

CURRENT_DATE_KEY = 'current_date'


def main():
    if len(sys.argv) < 2:
        setup_formfiller()
    else:
        fill_forms()


def setup_formfiller():
    try:
        set_reg(winreg.HKEY_CLASSES_ROOT, CSV_PROG_ID_PATH, CSV_PROG_ID)
        set_reg(winreg.HKEY_CURRENT_USER,
                RIGHT_CLICK_OPTION_PATH,
                RIGHT_CLICK_TITLE)
        set_reg(winreg.HKEY_CURRENT_USER,
                joinpath(RIGHT_CLICK_OPTION_PATH, 'command'),
                '"{}" "%1"'.format(sys.argv[0]))
        print('Installed successfully!')
    except PermissionError:
        print('Please run as administrator to install FormFiller')
    pause()


def fill_forms():
    params_path = get_params_path()
    print('Parameters: {}'.format(params_path))

    template_path = get_template_path()
    print('Template: {}'.format(template_path))

    with open(params_path, encoding='utf8') as fp:
        params = [row[:2] for row in csv.reader(fp)]
    params.append([CURRENT_DATE_KEY, date.today().strftime('%d/%m/%Y')])

    doc = DocxTemplate(template_path)
    doc.render(params)
    doc.save(".\{}.docx".format(
        generate_file_name(params_path, template_path)))

    print('Filled form successfully!')
    pause()


def get_params_path():
    # TODO Make sure it's a CSV file and alert user if not
    if len(sys.argv) > 1:
        return sys.argv[1]
    return input('Enter parameters file path: ')


def get_template_path():
    # TODO Make sure it's a DOCX file and alert user if not
    return input('Enter template file path: ')


def generate_file_name(params_path, template_path):
    params_name = splitext(basename(params_path))[0]
    template_name = splitext(basename(template_path))[0]
    return '{}_{}'.format(template_name, params_name)


def set_reg(key, sub_path, value, name=None):
    winreg.CreateKey(key, sub_path)
    regkey = winreg.OpenKey(key, sub_path, 0, winreg.KEY_WRITE)
    winreg.SetValueEx(regkey, name, 0, winreg.REG_SZ, value)
    winreg.CloseKey(regkey)


def pause():
    check_call(['pause'], shell=True)


if __name__ == '__main__':
    main()
