#!/usr/bin/env python

import shlex
import sys
from os.path import basename, isfile
from subprocess import DEVNULL
from subprocess import run as run_proc

SCRIPT_PATH = 'gui.py'
ICON_PATH = 'formfiller.ico'
APP_NAME = 'Form Filler'
# Commands
PYTHON_CMD = 'python'
PIP_CMD = PYTHON_CMD + ' -m pip --no-input'
CHECK_PACKAGE_CMD = PIP_CMD + ' show {0}'
INSTALL_PACKAGE_CMD = PIP_CMD + ' install {0}'
PYINSTALLER_CMD = 'pyinstaller --log-level {log_level} --icon={icon_path}' \
    ' --clean -n "{app_name}" -ywF "{python_file}"'


def main():
    # Make sure python dependencies are installed
    ensure_package('docxtpl')
    ensure_package('pyinstaller')
    ensure_package('pandas')
    ensure_package('openpyxl')

    # Make sure the script exits
    soft_assert(isfile(SCRIPT_PATH), "Can't find '{}' script to compile"
                .format(basename(SCRIPT_PATH)))

    compile_script(SCRIPT_PATH, APP_NAME, ICON_PATH)
    print('Executable should be found under ./dist folder')


def soft_assert(condition_result, error_message):
    """
    Checks the given condition and raises an exception if it's false.
    Unlike normal assert, this method raises an exception without printing the 
    traceback, only the given error_message.
    """
    if not condition_result:
        raise SystemExit('ERROR: {}'.format(error_message))


def run(cmd, quiet=True, **kwargs):
    cmd_args = shlex.split(cmd)
    if not quiet or is_debug:
        print('Running command: {}'.format(cmd))
        return run_proc(cmd_args, **kwargs)
    else:
        return run_proc(cmd_args, stdout=DEVNULL, stderr=DEVNULL, **kwargs)


def install_package(package_name):
    print("Installing package {}".format(package_name))
    cmd = run(INSTALL_PACKAGE_CMD.format(package_name))
    soft_assert(
        cmd.returncode == 0,
        "Failed to install package '{}', exiting script".format(package_name))


def ensure_package(package_name):
    cmd = run(CHECK_PACKAGE_CMD.format(package_name))
    if cmd.returncode != 0:
        install_package(package_name)


def compile_script(script_path, app_name, icon_path):
    print('Compilation starting...')
    log_level = 'DEBUG' if is_debug else 'WARN'

    cmd = PYINSTALLER_CMD.format(
        log_level=log_level,
        python_file=script_path,
        icon_path=icon_path,
        app_name=app_name)

    proc = run(cmd, quiet=False)
    soft_assert(proc.returncode == 0, "Compilation failed :(")
    print('Compiled successfully :)')


if __name__ == '__main__':
    is_debug = '--debug' in sys.argv
    main()
