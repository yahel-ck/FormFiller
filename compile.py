#!/usr/bin/env python

import shlex
import sys
from os import chdir, mkdir
from os.path import basename, isdir, isfile
from subprocess import DEVNULL
from subprocess import run as run_proc

# The path to store the compilation output
BUILD_DIR_PATH = './build'
# The path to the script to compile, relative to build folder
SCRIPT_PATH = '../formfiller.py'
# Commands
PYTHON_CMD = 'python'
PIP_CMD = PYTHON_CMD + ' -m pip --no-input'
CHECK_PACKAGE_CMD = PIP_CMD + ' show {0}'
INSTALL_PACKAGE_CMD = PIP_CMD + ' install {0}'
PYINSTALLER_CMD = 'pyinstaller --log-level {log_level} --icon=formfiller.ico' \
    ' -ywF {python_file}'


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


def compile_script(script_path):
    print('Compilation starting...')
    log_level = 'DEBUG' if is_debug else 'WARN'
    cmd = PYINSTALLER_CMD.format(log_level=log_level, python_file=script_path)
    proc = run(cmd, quiet=False)
    soft_assert(proc.returncode == 0, "Compilation failed :(")
    print('Compiled successfully :)')


def main():
    # Make sure python dependencies are installed
    ensure_package('docxtpl')
    ensure_package('pyinstaller')

    # Create and enter build folder
    if not isdir('./build'):
        mkdir('./build')
    chdir('./build')

    # Make sure the script exits
    soft_assert(isfile(SCRIPT_PATH), "Can't find '{}' script to compile"
                .format(basename(SCRIPT_PATH)))

    compile_script(SCRIPT_PATH)
    print('Executable should be found under build/dist folder')


if __name__ == '__main__':
    is_debug = '--debug' in sys.argv
    main()
