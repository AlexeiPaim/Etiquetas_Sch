from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import landscape
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import qrcode
from customtkinter import CTkImage
from PIL import Image
from tkinter import *
import customtkinter
from customtkinter import CTk
import math
import os
from customtkinter import *





customtkinter.set_appearance_mode("Light") # Other: "Light", "System" (only macOS)
customtkinter.set_default_color_theme("green")  # Themes: blue (default), dark-blue, green


# Variáveis globais para armazenar os caminhos selecionados
selected_excel_path = ""
selected_directory_path = ""


pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
pdfmetrics.registerFont(TTFont('ArialBd', 'ArialBd.ttf'))

def mm_to_p(mm):
    return  mm/0.352777


def ler_excel_e_armazenar_dados(nome_arquivo, linha):


    try:
        # Lê o arquivo Excel
        df = pd.read_excel(nome_arquivo)

        tamanho= len(df) 

        # Verifica se o dataframe não está vazio
        if not df.empty:
            # Extrai as 16 primeiras colunas da primeira linha
            primeira_linha = df.iloc[linha, 0:16]

            # Armazena os valores em variáveis individuais
            var1 = primeira_linha.iloc[0]
            var2 = primeira_linha.iloc[1]
            var3 = primeira_linha.iloc[2]
            var4 = primeira_linha.iloc[3]
            var5 = primeira_linha.iloc[4]
            var6 = primeira_linha.iloc[5]
            var7 = primeira_linha.iloc[6]
            var8 = primeira_linha.iloc[7]
            var9 = primeira_linha.iloc[8]
            var10 = primeira_linha.iloc[9]
            var11 = primeira_linha.iloc[10]
            var12 = primeira_linha.iloc[11]
            var13 = primeira_linha.iloc[12]
            var14 = primeira_linha.iloc[13]
            var15 = primeira_linha.iloc[14]
            var16 = primeira_linha.iloc[15]

            # Retorna as variáveis
            return var1, var2, var3, var4, var5, var6, var7, var8, var9, var10, var11, var12, var13, var14, var15, var16
        else:
            print("O arquivo está vazio.")
            return None
    except Exception as e:
        print("Ocorreu um erro ao ler o arquivo:", e)
        return None


def create_pdf(tamanho):

    for i in range(0,tamanho):

        # Substitua 'nome_do_arquivo.xlsx' pelo nome do seu arquivo Excel
        
    
        tipo, Ur, Ud, Up, Un, ik, tk, fr, lc, link, air, pre, ir, iac, sn, unifilar = ler_excel_e_armazenar_dados(selected_excel_path,i)
        
        if i == 0:
            # Definir as dimensões do arquivo PDF (130x36 mm)
            width, height = mm_to_p(130), mm_to_p(36)


            # Caminho para salvar o arquivo PDF
            caminho = selected_directory_path
            nome_arquivo = "ola_mundo9.pdf"

            # Concatenar caminho e nome do arquivo
            caminho_completo = os.path.join(caminho, nome_arquivo)

            print(caminho_completo)
            
            # Criar um arquivo PDF com as dimensões especificadas
            c = canvas.Canvas(caminho_completo, pagesize=(width, height))

        # Definir o texto a ser escrito no arquivo PDF
        text = "SM Air"
        
        # Posicionar o texto no centro da página
    
        x = mm_to_p(2)
        y = mm_to_p(32)
        
        # Escrever o texto no arquivo PDF
        c.setFont('Arial',11)
        c.drawString(x, y, text)
        
        # Criar fontes separadas
        c.setFont('ArialBd', 11)
        # Desenhar texto com fontes apropriadas
        c.drawString(mm_to_p(14), mm_to_p(32), "SeT")
        
        #Retangulo maior 
        c.roundRect(mm_to_p(23),mm_to_p(18), mm_to_p(58), mm_to_p(17), mm_to_p(1))

        #Retangulo Abaixo do maior 
        c.roundRect(mm_to_p(23),mm_to_p(2), mm_to_p(58), mm_to_p(15), mm_to_p(1))

        #Retangulo Direita cima 
        c.roundRect(mm_to_p(82),mm_to_p(28), mm_to_p(46), mm_to_p(7), mm_to_p(1))

        c.roundRect(mm_to_p(82),mm_to_p(18), mm_to_p(46), mm_to_p(9), mm_to_p(1))

        c.roundRect(mm_to_p(82),mm_to_p(2), mm_to_p(22), mm_to_p(15), mm_to_p(1))

        c.roundRect(mm_to_p(105),mm_to_p(2), mm_to_p(23), mm_to_p(15), mm_to_p(1))

        
        # Adicionar o código QR
        qr = qrcode.make(link)
        qr_width, qr_height = mm_to_p(20), mm_to_p(20)
        qr_x, qr_y = mm_to_p(2), mm_to_p(1.5)
        c.drawInlineImage(qr, qr_x, qr_y, width=qr_width, height=qr_height)

        #montar texto 
        c.setFont('Arial',8)
        c.drawString(mm_to_p(2), mm_to_p(27), "Type")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(25), mm_to_p(31), "Ur:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(25), mm_to_p(27), "Ud:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(25), mm_to_p(23), "Up:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(25), mm_to_p(19), "Un:")


        c.setFont('Arial',9)
        c.drawString(mm_to_p(54), mm_to_p(31), "ik:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(54), mm_to_p(27), "tk:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(54), mm_to_p(23), "fr:")


        c.setFont('Arial',8)
        c.drawString(mm_to_p(24), mm_to_p(14), "Sealed pressure system acc. IEC62271-1")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(24), mm_to_p(9), "AIR:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(24), mm_to_p(4), "Pre:")

        c.setFont('Arial',7)
        c.drawString(mm_to_p(84), mm_to_p(32), "Internal arc performance:")

        c.setFont('Arial',7)
        c.drawString(mm_to_p(84), mm_to_p(20), "Made in:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(84), mm_to_p(23.5), "SN:")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(90), mm_to_p(8), "Ir:")

        # Inserir a imagem de 14x14 pixels
        image_path = "C:/EtiquetasSMAirSeT/carinha.PNG"  # Substitua pelo caminho da sua imagem
        image_width = mm_to_p(8)
        image_height = mm_to_p(7)
        image_x = mm_to_p(106.5)  # Posição X da imagem no PDF
        image_y = mm_to_p(8)  # Posição Y da imagem no PDF
        c.drawImage(image_path, image_x, image_y, width=image_width, height=image_height)
        
        c.setFont('Arial',9)
        c.drawString(mm_to_p(106.5), mm_to_p(4), "NNZ1586901")
        
        
        if unifilar == "MedidaUnifilária":
            # Inserir a imagem de 14x14 pixels
            image_path1 = "C:/EtiquetasSMAirSeT/1.PNG"  # Substitua pelo caminho da sua imagem
            image_width1 = mm_to_p(7)
            image_height1 = mm_to_p(11)
            image_x1 = mm_to_p(83)  # Posição X da imagem no PDF
            image_y1 = mm_to_p(4)  # Posição Y da imagem no PDF
            c.drawImage(image_path1, image_x1, image_y1, width=image_width1, height=image_height1)
            
        if  unifilar == "Disjuntordefioúnico":
            # Inserir a imagem de 14x14 pixels
            image_path1 = "C:/EtiquetasSMAirSeT/2.PNG"  # Substitua pelo caminho da sua imagem
            image_width1 = mm_to_p(7)
            image_height1 = mm_to_p(11)
            image_x1 = mm_to_p(83)  # Posição X da imagem no PDF
            image_y1 = mm_to_p(4)  # Posição Y da imagem no PDF
            c.drawImage(image_path1, image_x1, image_y1, width=image_width1, height=image_height1)
            
        if  unifilar == "InterruptorUnifilar":
            # Inserir a imagem de 14x14 pixels
            image_path1 = "C:/EtiquetasSMAirSeT/3.PNG"  # Substitua pelo caminho da sua imagem
            image_width1 = mm_to_p(7)
            image_height1 = mm_to_p(11)
            image_x1 = mm_to_p(83)  # Posição X da imagem no PDF
            image_y1 = mm_to_p(4)  # Posição Y da imagem no PDF
            c.drawImage(image_path1, image_x1, image_y1, width=image_width1, height=image_height1)
            
        if  unifilar == "Interruptorfusíveldefioúnico":
            # Inserir a imagem de 14x14 pixels
            image_path1 = "C:/EtiquetasSMAirSeT/4.PNG"  # Substitua pelo caminho da sua imagem
            image_width1 = mm_to_p(7)
            image_height1 = mm_to_p(11)
            image_x1 = mm_to_p(83)  # Posição X da imagem no PDF
            image_y1 = mm_to_p(4)  # Posição Y da imagem no PDF
            c.drawImage(image_path1, image_x1, image_y1, width=image_width1, height=image_height1)
            
            
            
        
        

        c.setFont('Arial',10)
        c.drawString(mm_to_p(2), mm_to_p(22), tipo)
        
        
        c.setFont('Arial',9)
        c.drawString(mm_to_p(31), mm_to_p(31), Ur)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(31), mm_to_p(27), Ud)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(31), mm_to_p(23), Up)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(31), mm_to_p(19), Un)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(60), mm_to_p(31), ik)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(60), mm_to_p(27), tk)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(60), mm_to_p(23), fr)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(54), mm_to_p(19), lc)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(31), mm_to_p(9), air)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(31), mm_to_p(4), pre)

        c.setFont('Arial',9)
        c.drawString(mm_to_p(94), mm_to_p(8), ir)

        c.setFont('Arial',7)
        c.drawString(mm_to_p(94), mm_to_p(20), "Brazil")

        c.setFont('Arial',9)
        c.drawString(mm_to_p(90), mm_to_p(23.5), sn)
        
        c.setFont('Arial',10)
        c.drawString(mm_to_p(91), mm_to_p(28.5), iac)


        x_offset = mm_to_p(22.5)
        y_offset = mm_to_p(2)

        c.translate(x_offset, y_offset)
        c.rotate(90)
        c.setFont('Arial', 6)
        c.drawString(0, 0, "IEC 6221-200 / EN 62221-103")
        c.rotate(90)
        c.translate(-x_offset, -y_offset)

        # Fechar a página atual
        c.showPage()
        
        if i >= 4:
            #print("chegou")
            # Fechar o arquivo PDF
            c.save()
    
def select_excel_file():
    global selected_excel_path
    # Abre o explorador de arquivos para selecionar um arquivo Excel
    selected_file = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    
    # Atualiza o label com o caminho do arquivo selecionado
    if selected_file:
        selected_excel_path = selected_file
        excel_file_label.configure(text="Arquivo selecionado: " + selected_file)
    else:
        excel_file_label.configure(text="Nenhum arquivo selecionado.")

def select_directory():
    global selected_directory_path
    # Abre o explorador de arquivos para selecionar um diretório
    selected_directory = filedialog.askdirectory()
    
    # Atualiza o label com o diretório selecionado
    if selected_directory:
        selected_directory_path = selected_directory
        directory_label.configure(text="Diretório selecionado: " + selected_directory)
    else:
        directory_label.configure(text="Nenhum diretório selecionado.")

# Função para encerrar o loop da tela
def close_window():
    root.quit()



if __name__ == "__main__":

    # Cria uma instância da janela principal
    root = CTk()
    root.title("Etiquetas de caracteristicas SM AirSeT")
    root.geometry("800x400")  # Usando o método geometry() do Tkinter
    root.iconbitmap('C:/EtiquetasSMAirSeT/se.PNG')
    

   
    label_widget = CTkLabel(
        master=root,
        text="SM AirSeT",
        font=("Arial", 32,"bold")  # Define o tamanho da fonte
    )
    label_widget.place(relx=0.75, rely=0.4)  # Posiciona o texto na tela

    label_widget = CTkLabel(
        master=root,
        text="Etiqueta de Caracteristícas",
        font=("Arial", 30,"bold")  # Define o tamanho da fonte
    )
    label_widget.place(relx=0.02, rely=0.05)  # Posiciona o texto na tela

    from PIL import Image
  
    my_image =CTkImage(light_image=Image.open('C:/EtiquetasSMAirSeT/se.PNG'),
        dark_image=Image.open('C:/EtiquetasSMAirSeT/se.PNG'),
        size=(100,100)) # WidthxHeight

    my_label =CTkLabel(root, text="", image=my_image)
    my_label.place(relx=0.8, rely=0.1)
   
    from customtkinter import CTkCanvas
    
    select_excel_button = CTkButton(
        root,
        text="Selecionar Arquivo Excel",
        command=select_excel_file,  # Replace with your function
        width=120,
        height=32,
        border_width=1,
        corner_radius=8,
        font=("Arial", 30),  # Increased font size
        #fg_color="green"  # Button text color
    )
    select_excel_button.place(relx=0.04, rely=0.2)

    select_directory_button = CTkButton(
        root, 
        text="Selecionar Diretório de Saída", 
        command=select_directory,
        width=120,
        height=32,
        border_width=1,
        corner_radius=8,
        font=("Arial", 30) # Increased font size
    )
    select_directory_button.place(relx=0.01, rely=0.4)

  
    close_button = CTkButton(
            root,
            text="Salvar",
            command=close_window,
            width=120,
            height=32,
            border_width=1,
            corner_radius=8,
            font=("Arial", 30)
    )
    close_button.place(relx=0.18, rely=0.7)



    excel_file_label = CTkLabel(root, text="")
    excel_file_label.place(relx=0.01, rely=0.3)

    directory_label = CTkLabel(root, text="")
    directory_label.place(relx=0.01, rely=0.5)

     # Inicia o loop principal da interface gráfica
    root.mainloop()

    # Leitura do arquivo Excel
    df = pd.read_excel(selected_excel_path)

    # Obtém o tamanho do dataframe
    tamanho = len(df)

    # Imprime o tamanho
    print(tamanho)

    # Cria o PDF
    create_pdf(tamanho)