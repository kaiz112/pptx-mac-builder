from setuptools import setup

APP = ['PPT_text_4.2.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'iconfile': 'icon.icns',  # Optional: if you have a macOS icon
    'packages': ['pptx', 'pandas', 'openpyxl']
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
