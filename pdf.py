#!/usr/bin/env python
# -*- coding: utf-8 -*-

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
# Importar deste jeito para não dar incompatibilidade com o Image do pil
import reportlab.platypus
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY


class GerarEtiquetasPDF:
    """
    Usada para gerar etiquetas de produtos em formato A4
    """
    
    def __init__(self, output="table.pdf"):
        """
        Recebe o caminho com o nome do arquivo final.
        """
        
        # Lista com os dados a serem impressos
        self.dados = []
        
        self.doc = SimpleDocTemplate(
                        output,
                        pagesize=A4,
                        rightMargin=1*cm,
                        leftMargin=1*cm,
                        topMargin=0.3*cm,
                        bottomMargin=0.3*cm
                        )

        
        
        self.styleSheet = getSampleStyleSheet()
        # Cria um novo estilo para ser adicionado
        self.styleSheet.add( ParagraphStyle( name = 'Center', alignment = TA_CENTER ) )
        
    def GeraDados(self, NomeProd, ImgCod):
        """
        Salva na lista de dados os produtos
        """
        if len(self.dados) != 0:
            if len(self.dados[-1]) == 3:
                self.dados.append([[NomeProd, ImgCod]])
            else:
                self.dados[-1].append([NomeProd, ImgCod])
                
        else:
            self.dados.append([[NomeProd, ImgCod]])
        
        
        
    def SetDados(self, nomeproduto, imagecod, qtd):
        """
        Recebe a descricao do produto e o caminho para a imagem com seus codigos
        """
        
        # Variavel I guarda a imagem
        I = reportlab.platypus.Image(imagecod)
        I.drawHeight = 1.5*cm
        I.drawWidth = 3*cm
        # Variavel T guarda o nome do produto
        T = Paragraph('''<font size=7>''' + nomeproduto.upper() + '''</font>''',self.styleSheet["Center"])
        
        for i in range(qtd):
            self.GeraDados(T,I)
        
       
        #data = [
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #        [[T,I], [T,I], [T,I]],
        #    ]
    def GeraEtiquetasBranco(self):
        # Abrindo imagem em branco
        I = reportlab.platypus.Image("Dados/blank.png")
        I.drawHeight = 1.5*cm
        I.drawWidth = 3*cm
        T = Paragraph(" ",self.styleSheet["Center"])
        self.GeraDados(T,I)
        
    def FinalizaPDF(self):
        
        if len(self.dados) == 1 and len(self.dados[-1]) < 3:            
            for i in range( 3 - len(self.dados[-1])):
                self.GeraEtiquetasBranco()
            
            
        if len(self.dados) <= 0:
            print "Sem nada para fazer"
        else:
            
            self.elementos = []
            t=Table(self.dados, 6.5*cm, 3.1*cm)
            t.setStyle(TableStyle([
                #('BACKGROUND', (1,1), (-2,-2), colors.green),
                #('TEXTCOLOR', (0,0), (1, -1), colors.red),
                #('INNERGRID', (0,0), (-1,-1), 0.5*cm, colors.white),
                #('INNERGRID', (0,0), (-1,-1), 0.5*cm, colors.black),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ]))
            
            self.elementos.append(t)
            
            self.doc.build(self.elementos)
        
                    
            
    