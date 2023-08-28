from cx_Freeze import setup, Executable

executables = [Executable('GeradordeQRCode.py')] 

setup(
    name='GPSVistaPDF',
    version='1.0',
    description='',
    executables=executables
)
