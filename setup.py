import sys

import cx_Freeze

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [cx_Freeze.Executable('PyStock - App.py', icon='LogoIco.ico', base=base)]

cx_Freeze.setup(
    name="PyStock",
    options={'build_exe': {'packages': ['PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'sqlite3', 'os', 'sys', 'openpyxl.drawing.image', 'datetime', 'time', 'tkinter.filedialog', 'openpyxl'],
                           'include_files': ['LogoIco.ico', 'View/', 'baseexcel.xlsx']}},
    executables=executables
)