#!/usr/bin/env python

import shlex
import sys
from os.path import basename, splitext
from os.path import join as joinpath
from subprocess import run as run_proc
import os

SCRIPT_PATH = 'app.py'
ICON_PATH = 'icon.ico'
WINDOW_ICON_PATH = 'icon.png'
APP_NAME = 'Form Filler'
ENSURE_PACKAGES = (
    'docxtpl', 'pyinstaller', 'pandas', 'openpyxl', 'PyPdf2'
)

# Commands
PIP_CMD = '"{}" -m pip --no-input'.format(sys.executable)
CHECK_PACKAGE_CMD = PIP_CMD + ' show {0}'
INSTALL_PACKAGE_CMD = PIP_CMD + ' install {0}'
PYINSTALLER_CMD = (
    ' --noconfirm'
    ' --clean'
    ' --noconsole'
    ' --icon={icon_path}'
    ' --add-data "{window_icon_path}{sep}."'
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
    print('Running: {}'.format(cmd_args))
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
    from PyInstaller.__main__ import run as run_pyinstaller
    cmd = PYINSTALLER_CMD.format(
        python_file=script_path,
        icon_path=icon_path,
        window_icon_path=window_icon,
        app_name=app_name,
        sep=';' if os.name == 'nt' else ':'
    )
    print('Running pyinstaller with args: {}'.format(cmd))
    run_pyinstaller(shlex.split(cmd))


def create_desktop_shortcut(file_path, shortcut_name=None):
    if shortcut_name is None:
        shortcut_name = splitext(basename(file_path))[0]
    if os.name == 'nt':
        ensure_package('pywin32')
        from win32com.shell import shell, shellcon
        desktop_path = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, 0, 0)
        shortcut_path = joinpath(desktop_path, '{}.lnk'.format(shortcut_name))
        shell.CreateShortcut(shortcut_path, file_path)
        print('Shortcut created at {}'.format(shortcut_path))
    else:
        print('Shortcut creation is not supported on this platform')


if __name__ == '__main__':
    main()
