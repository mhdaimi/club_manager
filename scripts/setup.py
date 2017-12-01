from distutils.core import setup
import py2exe, sys
 
sys.argv.append('py2exe')

setup(
    options = {'py2exe': {'bundle_files': 1, "includes": ["reportlab", "docx", "arabic_reshaper", "bidi", "pyPdf", "sqlite3", "Tkinter"]  }},
    zipfile = None,
    windows = [{
            "script":"address_book.py",
            "icon_resources": [(1, "logo.ico")],

            }],
)