import sys
from cx_Freeze import setup, Executable

# Lista de módulos necessários para o seu programa
include_modules = [
    "PIL.ImageDraw",
    "PIL.ImageFont",
    "PIL.ImageTk",
    "reportlab.lib.pagesizes",
    "tkinter",
    "pandas",
    "subprocess",
    "PIL.Image",
    "pyautogui",
    "textwrap",
    "qrcode",
    "time",
    "os",
    "configparser"
]

# Configuração do executável
exe = Executable(
    script="GeradordeQRCode.py", 
    base="Win32GUI",  # Use "Win32GUI" para ocultar a janela do console
)

# Configuração do setup
setup(
    name="GeradordeQRCode",  # Nome do seu aplicativo
    version="1.0",  # Versão
    description="Aplicativo feito por julio matias - Rio de Janeiro",
    options={"build_exe": {"includes": include_modules}},
    executables=[exe],
)
