#!/usr/bin/env python

import shlex
import sys
from os.path import basename, splitext, expanduser, isfile, isdir
from os.path import join as joinpath
from subprocess import DEVNULL
from subprocess import run as run_proc
from pycrosskit.shortcuts import Shortcut
from PyInstaller.__main__ import run as run_pyinstaller

SCRIPT_PATH = 'app.py'
ICON_PATH = 'icon.ico'
WINDOW_ICON_PATH = 'icon.png'
APP_NAME = 'Form Filler'
ENSURE_PACKAGES = ('docxtpl', 'pyinstaller', 'pandas', 'openpyxl', 'PyPdf2')

# Commands
PIP_CMD = sys.executable + ' -m pip --no-input'
CHECK_PACKAGE_CMD = PIP_CMD + ' show {0}'
INSTALL_PACKAGE_CMD = PIP_CMD + ' install {0}'
PYINSTALLER_CMD = (
    ' --noconfirm'
    ' --clean'
    ' --noconsole'
    ' --icon={icon_path}'
    ' --add-data "{window_icon_path}:."'
    ' -n "{app_name}"'
    ' "{python_file}"'
)


def main():
    # Make sure python dependencies are installed
    for pkg_name in ENSURE_PACKAGES:
        ensure_package(pkg_name)

    compile_script(SCRIPT_PATH, APP_NAME, ICON_PATH, WINDOW_ICON_PATH)
    print('Executable should be found under ./dist folder')
    create_desktop_shortcut(joinpath('./dist', APP_NAME + '.exe'))


def run(cmd, **kwargs):
    cmd_args = shlex.split(cmd)
    print('Running command: {}'.format(cmd))
    return run_proc(cmd_args, **kwargs)


def ensure_package(package_name):
    cmd = run(CHECK_PACKAGE_CMD.format(package_name))
    if cmd.returncode != 0:
        install_package(package_name)


def install_package(package_name):
    print("Installing package {}".format(package_name))
    cmd = run(INSTALL_PACKAGE_CMD.format(package_name))
    assert cmd.returncode == 0, "Package installation failed"


def compile_script(script_path, app_name, icon_path, window_icon):
    print('Compilation starting...')
    cmd = PYINSTALLER_CMD.format(
        python_file=script_path,
        icon_path=icon_path,
        window_icon_path=window_icon,
        app_name=app_name
    )
    print('Running pyinstaller with args: {}'.format(cmd))
    run_pyinstaller(shlex.split(cmd))


def create_desktop_shortcut(file_path, shortcut_name=None):
    return Shortcut(
        shortcut_name or splitext(basename(file_path))[0], 
        file_path,
        desktop=True,
        start_menu=True
    )


if __name__ == '__main__':
    main()
