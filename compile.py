#!/usr/bin/env python

import ctypes
import os
import shlex
import sys
from os.path import basename, expanduser, isdir
from os.path import join as joinpath
from os.path import splitext
from posixpath import abspath
from subprocess import run as run_proc

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

    create_desktop_shortcut(get_compiled_script_path(APP_NAME))


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


def get_compiled_script_path(app_name, output_dir='./dist'):
    if os.name == 'nt':
        return joinpath(output_dir, '{}.exe'.format(app_name))
    elif sys.platform == 'darwin':
        return joinpath(output_dir, '{}.app'.format(app_name))
    elif os.name == 'posix':
        return joinpath(output_dir, '{}'.format(app_name))


def create_desktop_shortcut(target_path, shortcut_name=None):
    """
    Creates a desktop shortcut to the specified file or directory.

    :param target_path: The path to the file or directory to create a shortcut 
    to. Can be relative or absolute.
    :param shortcut_name: The name of the shortcut file that will be placed in
    the desktop, if None will be generated automatically.
    """
    target_path = abspath(target_path)
    shortcut_name = shortcut_name or splitext(basename(target_path))[0]

    desktop_path = pathutil.get_desktop_path()
    if desktop_path and isdir(desktop_path):
        shortcut_path = joinpath(desktop_path, shortcut_name)
        print('Creating shortcut "{}" to "{}"'.format(shortcut_path, target_path))
        pathutil.create_shortcut(target_path, shortcut_path)
    else:
        print('Desktop folder not found, skipping shortcut creation')


class WindowsPathUtil(object):
    @classmethod
    def get_desktop_path(cls):
        return ctypes.windll.shell32.SHGetFolderPathW(0, 0, 0, 0)

    @classmethod
    def create_shortcut(cls, file_path, shortcut_path, is_dir=False,
                        extention='.lnk'):
        shortcut_path += extention
        kdll = ctypes.windll.LoadLibrary('kernel32.dll')
        kdll.CreateSymbolicLinkW(shortcut_path, file_path, int(is_dir))


class PosixPathUtil(object):
    @classmethod
    def get_desktop_path(cls):
        path = expanduser('~/Desktop')
        return path if isdir(path) else None

    @classmethod
    def create_shortcut(cls, file_path, shortcut_path, is_dir=False,
                        extention=''):
        shortcut_path += extention
        run('ln -s "{}" "{}"'.format(file_path, shortcut_path))


pathutil = WindowsPathUtil if os.name == 'nt' else PosixPathUtil


if __name__ == '__main__':
    main()
