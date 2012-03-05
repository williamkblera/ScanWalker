#!/usr/bin/env python
# -*- coding: utf-8 -*-
import Image
import ImageDraw
import ImageFont

def GeraImgCodBarrasProd(caminhoimg, codeprod):
    """
    Recebe como parametro o caminho para um arquivo com a imagem do codigo de barras e o
    codigo do produto. Substitui a imagem do codigo de barras pela imagem com codigo debarras
    e o codigo do produto.
    """
    
    # Abre codigo de barras
    img = Image.open(caminhoimg)
    
    # Tamanho da imagem com o cod produto
    tamanho = 50,20
    white = (255,255,255)
    # Imagem com codig produto
    img2 = Image.new('RGB', tamanho, white)
    # Salvando Cod produto na imagem
    draw = ImageDraw.Draw(img2)
    # Fonte arial bold tamanho 15
    font = ImageFont.truetype("arialbd.ttf", 15)
    texto = codeprod
    pos = 5,2
    draw.text(pos, texto, font=font, fill = 'black')
    # Rotacionando imagem cod produto
    img2 = img2.rotate(-90)
    
    # Nova imagem para guardar o cod barras e o cod produto
    novaImg = Image.new("RGB", (123, 50), white)
    # Tamanho imagem cod barras
    W, H = img.size
    # Tamanho imagem cod produto
    w, h = img2.size
    
    # Coloca o cod barras na nova imagem
    novaImg.paste(img, (0, 0, 103, H))
    # Coloca o cod produto na nova imagem
    novaImg.paste(img2,(W, 0, W+w, H))
    # Salva a nova imagem
    
    novaImg.save(caminhoimg)