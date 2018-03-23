import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {'packages': [], 'excludes': []}

setup(  name = 'excelProcess',
        version = '1.0',
        description = 'openpyxl',
        options = {'build_exe': build_exe_options},
        executables = [Executable('excelProcess.py')])
