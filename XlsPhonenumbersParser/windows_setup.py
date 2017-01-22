from distutils.core import setup
import py2exe
 
setup(
    console=['launcher.py'],
    options = {
        'py2exe': {
            'packages': [
                    'pyexcel',
                    'pyexcel_xls',
                    'pyexcel_xlsx',
                    'pyexcel_ods',
                    'tkinter'
                    ]
        }
    }
)