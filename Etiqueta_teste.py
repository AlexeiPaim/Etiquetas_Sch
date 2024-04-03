from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import landscape
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import qrcode


pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))

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

    for i in range(1,tamanho):

        # Substitua 'nome_do_arquivo.xlsx' pelo nome do seu arquivo Excel
        nome_arquivo = 'alexei.xlsx'
    
        tipo, Ur, Ud, Up, Un, ik, tk, fr, lc, link, air, pre, ir, iac, sn, unifilar = ler_excel_e_armazenar_dados(nome_arquivo,i)
        print(i)
        if i == 1:
            # Definir as dimensões do arquivo PDF (130x36 mm)
            width, height = mm_to_p(130), mm_to_p(36)
            
            # Criar um arquivo PDF com as dimensões especificadas
            c = canvas.Canvas("ola_mundo.pdf", pagesize=(width, height))

        # Definir o texto a ser escrito no arquivo PDF
        text = "SmAirSeT"
        
        # Posicionar o texto no centro da página
    
        x = mm_to_p(2)
        y = mm_to_p(32)
        
        # Escrever o texto no arquivo PDF
        c.setFont('Arial',11)
        c.drawString(x, y, text)
        
    
        # Desenhar o retângulo com bordas arredondadas

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
        image_path = "C:/Users/alexe/OneDrive/Documentos/Integracao_trab3/carinha.PNG"  # Substitua pelo caminho da sua imagem
        image_width = mm_to_p(8)
        image_height = mm_to_p(7)
        image_x = mm_to_p(106.5)  # Posição X da imagem no PDF
        image_y = mm_to_p(8)  # Posição Y da imagem no PDF
        c.drawImage(image_path, image_x, image_y, width=image_width, height=image_height)

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
        c.drawString(mm_to_p(90), mm_to_p(23.5), sn )

    

        # Fechar a página atual
        c.showPage()
        
        if i >= 4:
            print("chegou")
            # Fechar o arquivo PDF
            c.save()
    
    

if __name__ == "__main__":

    nome_arquivo = 'alexei.xlsx'

    # Lê o arquivo Excel
    df = pd.read_excel(nome_arquivo)

    tamanho = len(df) 

    print(tamanho)
    
    create_pdf(tamanho)

