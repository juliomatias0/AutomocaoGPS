from PIL import ImageDraw, ImageFont, ImageTk
from reportlab.lib.pagesizes import letter, A4
from tkinter import filedialog
from tkinter import *
import pandas as pd
import subprocess
import PIL.Image
import pyautogui
import textwrap
import qrcode
import time
import os

# Função para buscar o arquivo Excel contendo a planilha
def buscar():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    planilha = pd.read_excel(arquivo)
    return planilha

# Função para obter o nível escolhido pelo usuário
def obter_nivel_escolhido(var_nivel):
    nivel_escolhido = var_nivel.get()
    return nivel_escolhido

# Função para filtrar as informações da planilha de acordo com o nível escolhido
def filtrar_por_nivel(planilha, nivel):
    if nivel == "Todas":
        return planilha
    planilha_filtrada = planilha[planilha['Nível'] == nivel]
    return planilha_filtrada

# Função para gerar um QR code e inseri-lo em uma imagem base
def generate_qr_code(qrcode_data, lugar):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=7,
        border=2,
    )
    qr.add_data(qrcode_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")

    img_base = 'C:/Users/julio.matias/GeradordePDFEquipeDarioRJ/qrcode.png'
    imagem_base = PIL.Image.open(img_base).convert("RGB")

    x = 30 #42
    y = 54 #62
    imagem_base.paste(qr_img, (x, y))

    draw = ImageDraw.Draw(imagem_base)
    font = ImageFont.load_default()

    # Desenhar o texto centralizado abaixo do QR code com quebra de linha
    wrapped_lines = textwrap.wrap(lugar, width=55)  # Escolha o valor de "width" de acordo com sua largura de imagem

    # Calcular a altura total do texto considerando todas as linhas
    total_text_height = sum(draw.textsize(line, font=font)[1] for line in wrapped_lines)

    text_y = y + qr_img.size[1] + 15  # 10 pixels de espaço entre o QR code e o texto

    # Desenhar cada linha de texto em uma posição vertical separada
    for line in wrapped_lines:
        text_width, text_height = draw.textsize(line, font=font)
        text_x = (imagem_base.width - text_width) // 2
        draw.text((text_x, text_y), line, fill=(0, 0, 0), font=font)
        text_y += text_height  # Avançar para a próxima linha vertical

        text_x = (imagem_base.width - text_width) // 2

    # Ajustar a posição y para centralizar o texto verticalmente
    text_y = y + qr_img.size[1] + (total_text_height - text_height) // 2

    img_folder = 'C:/Users/julio.matias/GeradordePDFEquipeDarioRJ/QRCode_Imagens'
    img_path = os.path.join(img_folder, f"{lugar}_QRCODE.jpg")
    try:
        imagem_base.save(img_path)
    except OSError as e:
        print("Erro ao salvar a imagem:", str(e))
        return img_path

# Função para gerar o PDF com os QR codes
def generate_imagens(planilha):

    for index, row in planilha.iterrows():
        qrcode_data = row['QRCode']
        lugar = row['Descrição']
        img_path = generate_qr_code(qrcode_data, lugar)

    print("Imagens geradas com sucesso!")

def abrir_xnviewmp_com_pasta(pasta):
    try:
        caminho_xnviewmp = "C:/Users/julio.matias/GeradordePDFEquipeDarioRJ/xnViewMP/xnviewmp.exe"
        comando = [caminho_xnviewmp, pasta]
        subprocess.Popen(comando)
        print("xnviewmp aberto com sucesso.")
    except Exception as e:
        print("Erro ao abrir xnviewmp:", str(e))

def abrir_xnviewmp_com_pasta_selecao_automatica(pasta):
    try:
        caminho_xnviewmp = "C:/Users/julio.matias/GeradordePDFEquipeDarioRJ/xnViewMP/xnviewmp.exe"
        comando = [caminho_xnviewmp, pasta]
        subprocess.Popen(comando)  # Usar Popen em vez de run
        print("xnviewmp aberto com sucesso.")
        # Aguardar um breve atraso antes de continuar
        time.sleep(3)
        # Usar pyautogui para automatizar as ações
        pyautogui.hotkey("ctrl", "a")  # Selecionar todos os arquivos
        pyautogui.hotkey("ctrl", "p")  # Clicar na opção de "print"
        time.sleep(2)

        print("Ações de automação realizadas com sucesso.")
    except Exception as e:
        print("Erro ao abrir xnviewmp:", str(e))

def gerar():
    nivel_escolhido = var_nivel.get()  # Obter o nível escolhido
    print("Gerando QR codes para o nível:", nivel_escolhido)
    arquivo_excel = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])  # Solicita o caminho do arquivo
    planilha = pd.read_excel(arquivo_excel)  # Lê a planilha a partir do caminho do arquivo
    nivel_escolhido = obter_nivel_escolhido(var_nivel)
    planilha = filtrar_por_nivel(planilha, nivel_escolhido)
    texto_resposta.config(text="Gerando QR codes...")  # Mensagem de status
    janela.update()  # Atualizar a interface
    pasta_dos_qr_codes = 'C:/Users/julio.matias/GeradordePDFEquipeDarioRJ/QRCode_Imagens'

    if  var_selecao_automatica.get(): # Verificar a opção selecionada pelo usuário
        abrir_xnviewmp_com_pasta_selecao_automatica(pasta_dos_qr_codes)
    else:
        abrir_xnviewmp_com_pasta(pasta_dos_qr_codes)
    generate_imagens(planilha)

def main():
    planilha = buscar()  # Obtém o caminho do arquivo Excel escolhido pelo usuário
    nivel_escolhido = obter_nivel_escolhido(var_nivel)
    planilha_filtrada = filtrar_por_nivel(planilha, nivel_escolhido)
    generate_imagens(planilha_filtrada)
    gerar()

# Configuração da interface gráfica
janela = Tk()
janela.title("GPS Vista - Funcionalidade")
janela.geometry("600x200")
img_f= 'C:/Users/julio.matias/GeradordePDFEquipeDarioRJ/Visual.png'
figma_image = PIL.Image.open(img_f)
figma = ImageTk.PhotoImage(figma_image)

# Variável global para armazenar a opção de seleção automática
var_selecao_automatica = IntVar()
var_selecao_automatica.set(0)  # Valor padrão: seleção manua

# Criar um Label com a imagem como plano de fundo
background_label = Label(janela, image=figma)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

def nivel_selecionado(*args):
    nivel_escolhido = var_nivel.get()
    print("Nível selecionado:", nivel_escolhido)

# Variável global para armazenar o nível escolhido pelo usuário
var_nivel = StringVar()
var_nivel.set("Escolha o Nível")  # Valor padrão
var_nivel.trace("w", nivel_selecionado)  # Registrar a função de callback
opcoes_nivel = ["Todas","N1","N2","N3","N4","N5","CR"]

# Criar um Frame para conter o OptionMenu
frame_nivel = Frame(janela, highlightbackground="#ff0000", highlightthickness=1)
frame_nivel.grid(column=0, row=3, padx=20, pady=20)

# Criar o OptionMenu dentro do Frame
menu_nivel = OptionMenu(frame_nivel, var_nivel, *opcoes_nivel)
menu_nivel.pack(fill="both", expand=True)  # Preencher todo o espaço do frame

# Configurar cor do contorno quando o frame não está em foco
frame_nivel.bind("<FocusIn>", lambda event: frame_nivel.config(highlightbackground="blue"))
frame_nivel.bind("<FocusOut>", lambda event: frame_nivel.config(highlightbackground="#ff0000"))

# Configurar cor do contorno quando o mouse passa sobre o frame
frame_nivel.bind("<Enter>", lambda event: frame_nivel.config(highlightbackground="green"))
frame_nivel.bind("<Leave>", lambda event: frame_nivel.config(highlightbackground="#ff0000"))

botao = Button(janela, text="Gerar QR-Codes PDF", command=gerar)
botao.grid(column=0, row=4, padx=20, pady=20)

check_selecao_automatica = Checkbutton(janela, text="Seleção Automática", variable=var_selecao_automatica)
check_selecao_automatica.grid(column=1, row=3, padx=20, pady=10)

texto_resposta = Label(janela, bg="#708090", text="Selecione o arquivo!")
texto_resposta.grid(column=1, row=4, padx=20, pady=20)

# Início da execução da interface gráfica
janela.mainloop()

# Executar a função principal se este arquivo for o arquivo principal
if __name__ == "__main__":
    main()