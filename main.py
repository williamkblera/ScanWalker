#!/usr/bin/env python
# -*- coding: utf-8 -*-

from Tkinter import *
from tkFileDialog import asksaveasfilename,askopenfilename
import tkMessageBox
import tkFont
import ttk  # Python27

from produtos import *
from EANBarCode import *
from desenho import *
from pdf import *
from copy import deepcopy
import sys
import csv

from time import time # Para debug e testes
import pygame


class Alerta:
    """
    Usado para fazer alerta sonoro
    """
    def __init__(self, music="beep-1.wav"):
        """
        Recebe como parametro o arquivo de som do alerta
        """
        pygame.init()
        pygame.mixer.music.load(music)
    def bepp(self):
        """
        Aciona alerta sonoro
        """
        pygame.mixer.music.play()



# Lista com produtos do banco
listaProdOg = listaProdutos()
listaProdOg.PegaListadoBanco()

class JanelaCodUp:
    """
    Vincula os cod. de barra ao produto no banco.
    """
    
    def __init__(self, raiz):
        """
        Recebe a janela raiz para criar o toplevel.
        """
        
        self.produtosLido = listaProdutos()
        fileName = askopenfilename(filetypes=[('CSV Files','*.csv')])
        if fileName:
            arq = csv.reader(open(fileName), delimiter='\t')
            for [COD_PRODUTO,DESCRICAO,CODBARRAS,SALDOESTOQUE] in arq:
                if COD_PRODUTO != "0":
                    
                    self.produtosLido.adicionaProdutos(cod_produto=COD_PRODUTO, descricao=DESCRICAO, cod_barras=CODBARRAS, saldo_estoque=SALDOESTOQUE)
                
            
            
            self.janela = Toplevel()
            self.janela.title("Vincular Cod. de Barras")
            self.janela.resizable(FALSE,FALSE)
            self.janela.protocol("WM_DELETE_WINDOW",self.FechaJanela)  # Colocando a função de fechar a janela
            self.Forca_Focus()
            
            # Pegando posição da janela raiz e da toplevel
            #possjnEtiq =  self.jnEtiquetas.geometry().split('+')
            possRaiz = raiz.geometry().split('+')       
            possjnEtiquetas = str("970x350") + "+" + str(possRaiz[1]) + "+" + str(int(possRaiz[-1])+int(possRaiz[0].split('x')[1])+50)
            
            
            # Posicionado janela abaixo da janela de menu
            self.janela.geometry(possjnEtiquetas)
            
            # Criando variaveis p/ label com o total de produtos e produtos novos
            self.labelTotProd = StringVar()
            self.labelTotNewProd = StringVar()
            self.labelTotOldProd = StringVar()
            self.labelTotGeralProd = StringVar()
            
            self.labelTotProd.set(str(self.produtosLido.ContaItens()))
            self.labelTotNewProd.set(str(self.produtosLido.ContaNovosItens()))
            self.labelTotOldProd.set(str(self.produtosLido.ContaOldItens()))
            self.labelTotGeralProd.set(str(self.produtosLido.ContaTotalProdutos()))
            
            #############################
            #   Primeira linha da grid
            #############################
            
            # Codigo produto
            Label(self.janela, text="Cod. Produto").grid(row=0, column=0)
            self.jnetcodproduto = Entry(self.janela, name="jnetcodprod")
            self.jnetcodproduto["width"] = 10
            self.jnetcodproduto.grid(row=0, column=1)
            
            # Setando o evento do teclado para a descricao, quando qualwuer tecla for presionada chamda filtraProdutos
            self.jnetcodproduto.bind("<KeyRelease>", self.filtraProdutos)
            
            # Descricão
            Label(self.janela, text="Descrição").grid(row=0, column=2)
            self.jnetdescricao = Entry(self.janela, name="jnetdescricao")
            self.jnetdescricao["width"] = 50
            self.jnetdescricao.grid(row=0, column=3)
            self.jnetdescricao.focus_force()
            
            # Setando o evento do teclado para a descricao, quando qualwuer tecla for presionada chamda filtraProdutos
            self.jnetdescricao.bind("<KeyRelease>", self.filtraProdutos)
            
            # Codigo de Barras
            Label(self.janela, text="Cod. Barras").grid(row=0, column=4)
            self.jnetcodbarra = Entry(self.janela, name="jnetcodbarras")
            self.jnetcodbarra["width"] = 20
            self.jnetcodbarra.grid(row=0, column=5)
            
            # Setando o evento do teclado para o Cod. Barras, quando qualquer tecla for presionada chamda filtraProdutos
            self.jnetcodbarra.bind("<KeyRelease>", self.filtraProdutos)
            self.jnetcodbarra.bind("<FocusIn>", self.filtraProdutos)
            
            # Quantidades
            Label(self.janela, text="Saldo").grid(row=0, column=6)
            self.jnetsaldo = Entry(self.janela, name="jnetsaldo")
            self.jnetsaldo["width"] = 10
            self.jnetsaldo.grid(row=0, column=7)
            
            
            # Setando o evento do teclado para o Cod. Barras, quando qualwuer tecla for presionada chamda filtraProdutos
            self.jnetsaldo.bind("<KeyRelease>", self.filtraProdutos)
            self.jnetsaldo.bind("<FocusIn>", self.filtraProdutos)
            
            
            
            ################################
            #   Fim Primeira linha da grid #
            ################################
            
            ##############################
            #   Segunda linha da grid    #
            ##############################
            self.jnettreeview = self.CriaTreeView(self.janela, numlinhas=1, numcolunas=0, columnspan=8, rowspan=10)
            self.jnettreeview.grid(row=1, column=0, columnspan=8, rowspan=10)
            self.jnettreeview._name = "treeview"
            self.jnettreeview.bind("<Double-Button-1>", self.SelecionaProduto)
            self.jnettreeview.bind("<Return>", self.SelecionaProduto)
            self.jnettreeview.bind("<KeyRelease>", self.eventTreeV)
            
            ###############################
            #  Fim Segunda linha da grid  #
            ###############################
            
            ##############################
            #   Terceira linha da grid   #
            ##############################
           
            # Criando labels para total de paginas
            Label(self.janela, text="Total de Itens", font=("Helvetica", 10)).grid(row=12, column=0, sticky=W)
            Label(self.janela, text="Novos Cadastros ", font=("Helvetica", 10)).grid(row=12, column=1, sticky=W)
            Label(self.janela, text="Produtos já Cadastrados ", font=("Helvetica", 10)).grid(row=12, column=2, sticky=W)
            Label(self.janela, text="Total de Produtos contado ", font=("Helvetica", 10)).grid(row=12, column=3, sticky=W)
            
            
            Label(self.janela, textvariable=self.labelTotProd, font=("Helvetica", 10, "bold")).grid(row=13, column=0, sticky=W)
            Label(self.janela, textvariable=self.labelTotNewProd, font=("Helvetica", 10, "bold")).grid(row=13, column=1, sticky=W)
            Label(self.janela, textvariable=self.labelTotOldProd, font=("Helvetica", 10, "bold")).grid(row=13, column=2, sticky=W)
            Label(self.janela, textvariable=self.labelTotGeralProd, font=("Helvetica", 10, "bold")).grid(row=13, column=3, sticky=W)
            ###############################
            #  Fim Terceira linha da grid #
            ###############################
            
            ###############################
            #  Quarta linha da grid       #
            ###############################
            self.jnctlbuttongeraretq = Button(self.janela, text="Enviar",
                                              command=self.Enviar)
            self.jnctlbuttongeraretq.grid(row=14, column=1)
            
            self.jnctlbuttoncancelar = Button(self.janela, text="Cancelar",
                                              command=self.FechaJanela)
            self.jnctlbuttoncancelar.grid(row=14, column=2)
            
            self.btnimportar = Button(self.janela, text="Abrir Novo Arquivo",
                                        command=self.Abrir)
            self.btnimportar.grid(row=14, column=3)
            
            ###############################
            # Fim Quarta linha da grid    #
            ###############################
            
             
            self.cadastraListaProdutos(self.produtosLido, self.jnettreeview)
            
    def Abrir(self):
        """
        Abrir um novo arquivo com os dados.
        """
        self.produtosLido = listaProdutos()
        fileName = askopenfilename(filetypes=[('CSV Files','*.csv')])
        if fileName:
            arq = csv.reader(open(fileName), delimiter='\t')
            for [COD_PRODUTO,DESCRICAO,CODBARRAS,SALDOESTOQUE] in arq:
                if COD_PRODUTO != "0":
                    
                    self.produtosLido.adicionaProdutos(cod_produto=COD_PRODUTO, descricao=DESCRICAO, cod_barras=CODBARRAS, saldo_estoque=SALDOESTOQUE)
                
            # Apaga toda a treeview para ser inserida uma nova
            x = self.jnettreeview.get_children() 
            for item in x: 
                self.jnettreeview.delete(item)
            
            self.cadastraListaProdutos(self.produtosLido, self.jnettreeview)
            self.labelTotProd.set(str(self.produtosLido.ContaItens()))
            self.labelTotNewProd.set(str(self.produtosLido.ContaNovosItens()))
            self.labelTotOldProd.set(str(self.produtosLido.ContaOldItens()))
            self.labelTotGeralProd.set(str(self.produtosLido.ContaTotalProdutos()))
            
        
    def AlteraCodBarrasBanco(self, codproduto, codbarras):
        """
        Altera no banco o codigo de barras do produto passado.
        """
        
        sql = "UPDATE ACADPROD SET CODBARRAS = '" + codbarras + "' WHERE COD_PRODUTO = '" + codproduto + "'"
        print sql
        #SELECT a.COD_PRODUTO, a.DESCRICAO, a.CODBARRAS FROM ACADPROD a
        self.cur.execute(sql)
        self.con.commit()
        #UPDATE ACADPROD SET CODBARRAS = '111' WHERE COD_PRODUTO = '45'
            
    def cadastraListaProdutos(self, listaProdutos, treeV):
        # Coloca dados na tabela
        #print "Tamanho", len(listaProdutos)
        if listaProdutos == []:
            print "Lista Vazia"
            self.labelTotEtiq.set("0")
            self.labelTotPag.set("0")
            #######TODO:XXX Aqui
        else:
            #x = []
           
            for item in listaProdutos.produtos:
                #a = (item.cod_produto, item.descricao, item.cod_barras, item.saldo_estoque)
                #x.append(a)
                treeV.insert('', 'end', values=(item.cod_produto, item.descricao, item.cod_barras, item.saldo_estoque))
                # Calculando o total de etiquetas
               
                
            #x.reverse()
            
            
            
            
            #for y in x:
            #    treeV.insert('', 'end', values=y)
            
        
    def CriaTreeView(self, frame, numlinhas, numcolunas, columnspan, rowspan):
        '''
        Cria um treeview para exibir os produtos, pede como parametro o frame onde sera colocado
        numlinhas = Linha incial da treeview
        numcolunas = Coluna inicial da treeview
        columspan = qtd de colunas ocupadas pela treeview
        rowspan = qtd de linhas ocupadas pela treeview
        '''
        pcolum = numcolunas+columnspan  # ultima coluna ocupada pela treeview
        plinhas = numlinhas+rowspan  # ultima linha ocupada pela treeview
        
        
        # Titulos das colunas na treeview
        colunas = (
                'Cod. Produto', 
                'Descricao                                                                    ', 
                'Cod. Barras       ', 
                'Quantidade')
        
        treeV = ttk.Treeview(frame, columns=colunas, show="headings", takefocus=1, selectmode="extended")  
        
        # Scrollbar Vertical
        vsb = ttk.Scrollbar(frame, orient="vertical", command=treeV.yview)
        # Scroollbar Horizontal
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=treeV.xview)
        
        # Configurando Scrollbars na treeview
        treeV.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        treeV.grid(column=numcolunas, row=numlinhas, sticky='nsew')
        
        # a possição da barra de rolagem recebe a possição inicial do grid + o seu tamanho mais 1
        # Tanto para a linha como para a coluna
        vsb.grid(column=pcolum, row=numlinhas, sticky='ns', rowspan=rowspan)
        hsb.grid(column=numcolunas, row=plinhas, sticky='ew', columnspan=columnspan)
        
        # Coloca titulos nas colunas, falta criar funçao de ordenação
        for col in colunas:
                treeV.heading(col, text=col.title(),
                    command=lambda c=col: self.sortby(treeV, c, 0))
                
                # Altera o tamanho da coluna para o tamanho do titulo
                treeV.column(col, width=tkFont.Font().measure(col.title()))
        
        
        return treeV
        
    def Conecta(self):
        """
        Abre a conexão com o banco de dados
        """
        self.con = kinterbasdb.connect(host="10.1.1.10",
                                   database='C:\SIACPlus\siacCX.fdb',
                                   user="sysdba",
                                   password="masterkey",
                                   charset="ISO8859_1"
                                   )
        print "Conexão aberta"
        self.cur = self.con.cursor()
        
    def Enviar(self):
        """
        Enviar dados para o banco
        """
        
        if tkMessageBox.askokcancel("Enviar para o Banco?", "Você tem certeza que deseja atualizar estes codigos de barras?"):
            # Acessando o Banco
            self.Conecta()
            
            for prod in self.produtosLido.produtos:
                if prod.cod_produto != "0":
                    self.AlteraCodBarrasBanco(str(prod.cod_produto), str(prod.cod_barras))
                    #print prod.cod_produto
            
            # Fechando a conexão com o banco
            self.FechaConexao()
            mesg = str(self.produtosLido.ContaItens()) + " Codigos de Barras Atualizados com sucesso!"
            tkMessageBox.showinfo(
                        title="Sucesso",
                        message=mesg
                        )
        
        
    def eventTreeV(self, event):
        """
        Gerencia os eventos do teclado na treeview
        """
        print "Evento do teclado: ", event.keycode
        if event.widget._name == "treeview":
            if event.keycode != 40 and event.keycode != 38:
                # Forçando o focus na descricao
                
                if event.keycode != 13:
                
                    self.jnetcodproduto.delete(0, END)
                    self.jnetdescricao.delete(0, END)
                    self.jnetdescricao.insert(0, event.char.upper())
                    
                self.jnetdescricao.focus_force()
                
    def FechaConexao(self):
        """
        Fecha a conexão com o banco de dados
        """
        # Fechando acesso ao banco                
        self.con.close()
        print "Conexão fechada"
    
    def FechaJanela(self):
        """
        Fecha a janela principal
        """
        if tkMessageBox.askokcancel("Exit?", "Você tem certeza que deseja fechar?"):
            self.janela.destroy()
            self.janela = None
        else:
            self.janela.focus_set()
            
    def filtraProdutos(self, event):
        """
        Utilizada para filtrar os produtos na treeview.
        """
        pass
            
    def Forca_Focus(self, event=None):
        """
        Força o focus na janela etiquetas
        """
        
        if event is None :
            self.janela.focus_set()
        else:
            print "Evento ", event.widget._name
            
    def SelecionaProduto(self, event):
        """
        Pega o produto que foi selecionado na treeview e envia para a JanelaProduto
        """
        pass
    

class JanelaRaiz:
    """
    Classe com as principais funções de gerar a janela raiz do sistema.
    """
    def __init__(self):
        
        # Criando a janela raiz do sistema
        self.janelaraiz = Tk()
        self.janelaraiz.geometry('600x90+150+5')
        if sys.platform == "win32":
            self.janelaraiz.iconbitmap(default='icon.ico')
        self.janelaraiz.title('ScanWalker - Controle de estoque')
        self.janelaraiz.resizable(FALSE,FALSE)
        self.janelaraiz.protocol("WM_DELETE_WINDOW",self.FechaJanela)  # Colocando a função de fechar a janela
        
        
    def FechaJanela(self):
        if tkMessageBox.askokcancel("Exit?", "Você tem certeza que deseja fechar?"):
            self.janelaraiz.destroy()
            self.janelaraiz = None
        else:
            self.Forca_Focus()
            
    def Forca_Focus(self):
        """
        Força o focus na janela etiquetas
        """
        self.janelaraiz.focus_set()
        
        
    def IconMenu(self):
        """
        Responsavel por criar os menus por icones na janela principal
        """
        
        
        imgetiquetas = PhotoImage(file="iconImpressaoEtiquetas.gif") # imagem do botão gera etiquetas
        
        btngeraretiqueta = Button(
            self.janelaraiz,
            compound=TOP,
            width=70,
            height=75,
            image=imgetiquetas,
            text="Etiquetas",
            command=self.AbreJanelaEtiquetas
            )
        btngeraretiqueta.grid(column=0, row=0, padx=2, pady=2)
        
        btngeraretiqueta.image = imgetiquetas # salvando a imagem no botão para o garbage collection???
        
        imgctlestoque = PhotoImage(file="iconControlEstoque.gif") # imagem do botão gera etiquetas
        
        btnctlestoque = Button(
            self.janelaraiz,
            compound=TOP,
            width=70,
            height=75,
            image=imgctlestoque,
            text="Estoque",
            command=self.AbreJanelaCTLEstoque            
            )
        btnctlestoque.grid(column=1, row=0, padx=2, pady=2)
        
        btnctlestoque.image = imgctlestoque # salvando a imagem no botão para o garbage collection???
        
        
        imgCodBarrasUp = PhotoImage(file="database_up.gif") # imagem do botão gera etiquetas
        btnCodBarrasUp = Button(
            self.janelaraiz,
            compound=TOP,
            width=70,
            height=75,
            image=imgCodBarrasUp,
            text="Enviar \nCod. Barras",
            command=self.AbreJanelaCodUp
            )
        btnCodBarrasUp.grid(column=2, row=0, padx=2, pady=2)
        
        btnCodBarrasUp.image = imgCodBarrasUp # salvando a imagem no botão para o garbage collection???
        
        #
        imgGeraLista = PhotoImage(file="document-print.gif") # imagem do botão gera etiquetas
        btnGeraLista = Button(
            self.janelaraiz,
            compound=TOP,
            width=70,
            height=75,
            image=imgGeraLista,
            text="Gerar Listas\n de Produtos",
            command=self.GeraLista
            )
        btnGeraLista.grid(column=3, row=0, padx=2, pady=2)
        
        btnGeraLista.image = imgGeraLista # salvando a imagem no botão para o garbage collection???
        
        
        
    def GeraLista(self):
        """
        Abre a janela para gerar as listas com os produtos.
        """
        pass
    
    def AbreJanelaCTLEstoque(self):
        try:
            self.janelactlestoque 
        except:
            
            self.janelactlestoque = JanelaCTLEstoque(self.janelaraiz)
            
            
        else:
            if self.janelactlestoque.janelactlestoque  == None:
                self.janelactlestoque = JanelaCTLEstoque(self.janelaraiz)
            else:
                self.janelactlestoque.Forca_Focus()
                
    def AbreJanelaCodUp(self):
        try:
            self.janelacdup 
        except:
            
            self.janelacdup = JanelaCodUp(self.janelaraiz)
            
            
        else:
            if self.janelacdup.janela  == None:
                self.janelacdup = JanelaCodUp(self.janelaraiz)
            else:
                self.janelacdup.Forca_Focus()
        
    def AbreJanelaEtiquetas(self):
        try:
            self.janelaetiquetas
        except:
            
            self.janelaetiquetas = JanelaEtiquetas(self.janelaraiz)
            
            
        else:
            if self.janelaetiquetas.janelaetiquetas == None:
                self.janelaetiquetas = JanelaEtiquetas(self.janelaraiz)
            else:
                self.janelaetiquetas.Forca_Focus()
                
            
        
        
    def mainloop(self):
        '''
        Roda a interface.
        '''
        # Roda o programa
        self.janelaraiz.mainloop()
        
class JanelaCTLEstoque:
    """
    Classe responsavel por gerar a janela de controle de estoque
    """
    
    def __init__(self, janelaraiz):
        # Preparando o alerta sonoro
        self.alerta = Alerta()
        
        
        # Criando lista de produtos
        #self.listadeprodutos = listaProdutos()
        
        # Copiando os dados pegos no banco de dados
        self.listadeprodutos = deepcopy(listaProdOg)
        self.listaprodutoscont = listaProdutos()
        
        self.janelaraiz = janelaraiz
        
        # Criando Janela TopLevel
        self.janelactlestoque = Toplevel()
        self.janelactlestoque.title("Controle de Estoque")
        #self.janelactlestoque.resizable(FALSE,FALSE)
        self.janelactlestoque.protocol("WM_DELETE_WINDOW",self.FechaJanela)  # Colocando a função de fechar a janela
        #self.janelaetiquetas.focus_set()
        self.Forca_Focus()
        
        # Pegando posição da janela raiz e da toplevel
        #possjnEtiq =  self.jnEtiquetas.geometry().split('+')
        possRaiz = janelaraiz.geometry().split('+')       
        possjnctlestoque = str("770x330") + "+" + str(possRaiz[1]) + "+" + str(int(possRaiz[-1])+int(possRaiz[0].split('x')[1])+70)
        
        
        # Posicionado janela abaixo da janela de menu
        self.janelactlestoque.geometry(possjnctlestoque)
        
        # Criando variaveis p/ label com o total de produtos e produtos novos
        self.labelTotProd = StringVar()
        self.labelTotNewProd = StringVar()
        self.labelTotOldProd = StringVar()
        self.labelTotGeralProd = StringVar()
        
        self.labelTotProd.set("0")
        self.labelTotNewProd.set("0")
        self.labelTotOldProd.set("0")
        self.labelTotGeralProd.set("0")
        
        #############################
        #   Primeira linha da grid
        #############################
        
        
        
        # Codigo de Barras
        Label(self.janelactlestoque, text="Cod. Barras").grid(row=0, column=0)
        self.jnctlcodbarra = Entry(self.janelactlestoque, name="jnctlcodbarra")
        self.jnctlcodbarra["width"] = 20
        self.jnctlcodbarra.grid(row=0, column=1, columnspan=2)
        self.jnctlcodbarra.focus_force()
        
        # Setando o evento do teclado para a descricao, quando qualquer tecla for presionada chama filtraProdutos
        self.jnctlcodbarra.bind("<Return>", self.filtraProdutos)
        
        
        # Quantidades
        Label(self.janelactlestoque, text="Saldo").grid(row=0, column=2)
        self.jnetsaldo = Entry(self.janelactlestoque, name="saldo")
        self.jnetsaldo["width"] = 20
        self.jnetsaldo.grid(row=0, column=3, columnspan=2)
        self.jnetsaldo.bind("<Return>", self.Forca_Focus)
        
        ################################
        #   Fim Primeira linha da grid #
        ################################
        
        ##############################
        #   Segunda linha da grid    #
        ##############################
        self.jnctltreeview = self.CriaTreeView(self.janelactlestoque, numlinhas=1, numcolunas=0, columnspan=4, rowspan=10)
        self.jnctltreeview.grid(row=1, column=0, columnspan=4, rowspan=10)
        self.jnctltreeview.bind("<Double-Button-1>", self.AlteraQTD)
        
        ###############################
        #  Fim Segunda linha da grid  #
        ###############################
        # Criando labels para total de pagianas
        Label(self.janelactlestoque, text="Total de Itens", font=("Helvetica", 10)).grid(row=12, column=0, sticky=W)
        Label(self.janelactlestoque, text="Novos Cadastros ", font=("Helvetica", 10)).grid(row=12, column=1, sticky=W)
        Label(self.janelactlestoque, text="Produtos já Cadastrados ", font=("Helvetica", 10)).grid(row=12, column=2, sticky=W)
        Label(self.janelactlestoque, text="Total de Produtos contado ", font=("Helvetica", 10)).grid(row=12, column=3, sticky=W)
        
        
        Label(self.janelactlestoque, textvariable=self.labelTotProd, font=("Helvetica", 10, "bold")).grid(row=13, column=0, sticky=W)
        Label(self.janelactlestoque, textvariable=self.labelTotNewProd, font=("Helvetica", 10, "bold")).grid(row=13, column=1, sticky=W)
        Label(self.janelactlestoque, textvariable=self.labelTotOldProd, font=("Helvetica", 10, "bold")).grid(row=13, column=2, sticky=W)
        Label(self.janelactlestoque, textvariable=self.labelTotGeralProd, font=("Helvetica", 10, "bold")).grid(row=13, column=3, sticky=W)
            
        
        
        ###############################
        #  Terceira linha da grid     #
        ###############################
        self.jnctlbuttongeraretq = Button(self.janelactlestoque, text="Salvar",
                                          command=self.salvar)
        self.jnctlbuttongeraretq.grid(row=25, column=1)
        
        self.jnctlbuttoncancelar = Button(self.janelactlestoque, text="Cancelar",
                                          command=self.FechaJanela)
        self.jnctlbuttoncancelar.grid(row=25, column=2)
        
        self.btnimportar = Button(self.janelactlestoque, text="Importar..",
                                    command=self.Importar)
        self.btnimportar.grid(row=25, column=3)
        
        ###############################
        # Fim Terceira linha da grid  #
        ###############################
        
        
       
        
    def Importar(self):
        """
        Usada para importar dados que foram cadastrados anteriormente.
        """
        print "importar"
        tipo = ""
        fileName = askopenfilename(filetypes=[('CSV Files','*.csv')])
        tipo = fileName.split(".")
        
        # Forçando para ser um arquivo .csv
        if tipo[-1] != "csv":
            fileName += ".csv"
            
        try:
            arq = csv.reader(open(fileName), delimiter='\t')
            print "Arquivo aberto"
            for [COD_PROD, DESCRICAO, COD_BARRAS, SALDO] in arq:
                
                if self.listaprodutoscont.adicionaProdutos(COD_PROD.decode("utf-8"), DESCRICAO.decode("utf-8"), COD_BARRAS.decode("utf-8"), int(SALDO)):
                    produto2 = Produto(cod_produto=COD_PROD.decode("utf-8"), descricao=DESCRICAO.decode("utf-8"), cod_barras=COD_BARRAS.decode("utf-8"), saldo_estoque=int(SALDO))
                    #self.listadeprodutos.addProduto(produto2)
                    self.listadeprodutos.AlteraProduto(produto2)
                    pass
                else:
                    print "Falha ao adicionar o produto: ",  DESCRICAO
        except:
            print "Error"
        finally:
            
            print "Produtos adicionados com sucesso"
            apagaTreeView(self.jnctltreeview)
            self.cadastraListaProdutos(self.listaprodutoscont, self.jnctltreeview)
            self.Forca_Focus(event="codbarras")
            
            self.labelTotProd.set(str(self.listaprodutoscont.ContaItens()))
            self.labelTotNewProd.set(str(self.listaprodutoscont.ContaNovosItens()))
            self.labelTotOldProd.set(str(self.listaprodutoscont.ContaOldItens()))
            self.labelTotGeralProd.set(str(self.listaprodutoscont.ContaTotalProdutos()))
            
            #arq.close()
        
        
        
    def AlteraQTD(self, event):
       
        #Pega valores do item
        item = self.jnctltreeview.item(self.jnctltreeview.selection())
        
        valores = item.values()
        print valores
        if valores[2] != '':      
        
            self.alteraqtdproduto = JanelaProduto(magicnumber=8)
            self.alteraqtdproduto.acodprod.set(valores[2][0])
            self.alteraqtdproduto.adescprod.set(valores[2][1])
            self.alteraqtdproduto.acodbarraprod.set(valores[2][2])
            self.alteraqtdproduto.aqtdprod.set(valores[2][3])
            self.janelactlestoque.wait_window(self.alteraqtdproduto.alteraJn)
            
            if self.alteraqtdproduto.cancela is False :
                novoProd = self.alteraqtdproduto.get()
                prod = Produto(cod_produto=str(novoProd[0]), descricao=str(novoProd[1]), cod_barras=str(novoProd[2]), saldo_estoque=int(novoProd[3]))
                print "Pord: ", prod
                if prod.saldo_estoque <= 0:
                    
                    self.listaprodutoscont.RemoveProduto(prod, confirmar=False)
                else:                
                    self.listaprodutoscont.AlteraProduto(prod)
                
                apagaTreeView(self.jnctltreeview)
                self.cadastraListaProdutos(self.listaprodutoscont, self.jnctltreeview)
                self.Forca_Focus(event="codbarras")
        
        #self.alteraqtdproduto.SetPossition(self.janelaraiz)
        
        
    def cadastraListaProdutos(self, listaProdutos, treeV):
        # Coloca dados na tabela
        #print "Tamanho", len(listaProdutos)
        if listaProdutos == []:
            print "Lista Vazia"
            #######TODO:XXX Aqui
        else:
            x = []
            for item in listaProdutos.produtos:
                a = (item.cod_produto, item.descricao, item.cod_barras, item.saldo_estoque)
                x.append(a)
            x.reverse()
            
            for y in x:
                treeV.insert('', 'end', values=y)
        
    def CriaTreeView(self, frame, numlinhas, numcolunas, columnspan, rowspan):
        '''
        Cria um treeview para exibir os produtos, pede como parametro o frame onde sera colocado
        numlinhas = Linha incial da treeview
        numcolunas = Coluna inicial da treeview
        columspan = qtd de colunas ocupadas pela treeview
        rowspan = qtd de linhas ocupadas pela treeview
        '''
        pcolum = numcolunas+columnspan  # ultima coluna ocupada pela treeview
        plinhas = numlinhas+rowspan  # ultima linha ocupada pela treeview
        
        
        # Titulos das colunas na treeview
        colunas = (
                'Cod. Produto', 
                'Descricao                                                                    ', 
                'Cod. Barras       ', 
                'Quantidade')
        
        treeV = ttk.Treeview(frame, columns=colunas, show="headings", takefocus=1)  
        
        # Scrollbar Vertical
        vsb = ttk.Scrollbar(frame, orient="vertical", command=treeV.yview)
        # Scroollbar Horizontal
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=treeV.xview)
        
        # Configurando Scrollbars na treeview
        treeV.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        treeV.grid(column=numcolunas, row=numlinhas, sticky='nsew')
        
        # a possição da barra de rolagem recebe a possição inicial do grid + o seu tamanho mais 1
        # Tanto para a linha como para a coluna
        #vsb.grid(column=pcolum, row=numlinhas, sticky='ns')
        vsb.grid(column=pcolum, row=numlinhas, sticky='ns', rowspan=rowspan)
        hsb.grid(column=numcolunas, row=plinhas, sticky='ew', columnspan=columnspan)
        
        # Coloca titulos nas colunas, falta criar funçao de ordenação
        for col in colunas:
                treeV.heading(col, text=col.title(),
                    command=lambda c=col: self.sortby(treeV, c, 0))
                
                # Altera o tamanho da coluna para o tamanho do titulo
                treeV.column(col, width=tkFont.Font().measure(col.title()))
        
        
        return treeV 
        
        
        
    def FechaJanela(self):
        if tkMessageBox.askokcancel("Exit?", "Você tem certeza que deseja fechar?"):
            self.janelactlestoque.destroy()
            self.janelactlestoque = None
        else:
            self.Forca_Focus()
            
    def filtraProdutos(self, event):
        """
        Utilizada para filtrar os produtos na treeview.
        """
        # Transformando em maiuscula
        text = event.widget.get().upper()
        event.widget.delete(0,END)
        event.widget.insert(0,text)
        
        if event.widget._name == "jnctlcodbarra":  # Entry de descricao da janela etiquetas
            Qtd = self.jnetsaldo.get()
            if Qtd == "":
                Qtd = 1
            codDig = self.jnctlcodbarra.get()
            # Testando numero do cod. barras
            EAN13 = []
            bar = EanBarCode()
            for i in codDig:
                EAN13.append(int(i))
            if len(EAN13) == 13:
                # Colocar para verificar o codigo
                if bar.verifyChecksum(EAN13):
                    # pega produto
                    produto = self.listadeprodutos.getInfosCodBarras(codDig)
                    if produto != None:
                        for i in range(int(Qtd)):
                            self.listaprodutoscont.addProduto(produto)
                            
                            
                        self.labelTotProd.set(str(self.listaprodutoscont.ContaItens()))
                        self.labelTotNewProd.set(str(self.listaprodutoscont.ContaNovosItens()))
                        self.labelTotOldProd.set(str(self.listaprodutoscont.ContaOldItens()))
                        self.labelTotGeralProd.set(str(self.listaprodutoscont.ContaTotalProdutos()))
                        
                        self.jnetsaldo.delete(0, END)
                        self.jnctlcodbarra.focus_force()
                        apagaTreeView(self.jnctltreeview)
                        self.cadastraListaProdutos(self.listaprodutoscont, self.jnctltreeview)
                    else:
                        # Codigo de Barras lido ainda não foi cadastrado
                        self.alerta.bepp()
                        if tkMessageBox.askokcancel("Atenção", "Produto não cadastrado, deseja cadastra-lo?"):
                            # Abrindo janela para alt
                            
                            self.alteracodpoduto = JanelaAlteraProduto(self.janelaraiz, codDig)
                            self.janelaraiz.wait_window(self.alteracodpoduto.janela)
                            self.listadeprodutos = deepcopy(listaProdOg)
                            
                            
                        else:
                            print "Não adicionar"
                        
                    
                else:
                    self.alerta.bepp()
                    tkMessageBox.showinfo(
                        title="Atenção",
                        message="O Cod. de Barras lido é invalido!"
                        )
                
            else:
                self.alerta.bepp()
                tkMessageBox.showinfo(
                        title="Atenção",
                        message="O Cod. de Barras lido é invalido!"
                        )
            #self.listadeprodutos
            #self.listaprodutoscont
            
            self.jnctlcodbarra.delete(0,END)
            self.jnctlcodbarra.focus_force()
           
            
                    
        #elif event.widget._name == "jnetcodprod": # Codigo do produto
        #    pass
            
        
    def Forca_Focus(self, event=None):
        """
        Força o focus na janela etiquetas
        """
        
        if event is None :
            self.janelactlestoque.focus_set()
        elif event == "codbarras":
            self.jnctlcodbarra.focus_force()
        elif event.widget._name == "saldo":
            self.jnctlcodbarra.focus_force()
        else:
            print "Evento ", event.widget._name
            
    def salvar(self):
        """
        Salva os produtos em um arquivo csv
        """
        tipo = ""
        fileName = asksaveasfilename(filetypes=[('CSV Files','*.csv')])
        tipo = fileName.split(".")
        # Forçando para ser um arquivo .csv
        if tipo[-1] != "csv":
            fileName += ".csv"
        
        #csv.register_dialect(
        try:
            
            #file = open(fileName, 'w')
            
            #arq = csv.write(open(fileName, "wb"))
            arq = open(fileName, 'w')
            for prod in self.listaprodutoscont.produtos:
                #arq.writerow([prod.cod_produto, prod.descricao, prod.cod_barras, prod.saldo_estoque])
                arq.write('%s\t' % prod.cod_produto)
                arq.write('%s\t' % prod.descricao.encode("utf-8"))
                arq.write('%s\t' % prod.cod_barras)
                arq.write('%s\n' % prod.saldo_estoque)
            #textoutput = self.text.get(0.0, END)
            #file.write(textoutput)
        except:
            print "Error"
        finally:
            print "Dados salvo com sucesso"
            #for prod in self.listaProdutosJanela.produtos:
            #    print prod.cod_produto, " " ,prod.descricao, " " ,prod.cod_barras, " " ,prod.saldo_estoque
            arq.close()
        
    def sortby(self, tree, col, descending):
        '''
        Organiza o conteudo da treeview baseado na coluna em que houve o click.
        '''
        # grab values to sort
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
    
        # reorder data
        data.sort(reverse=descending)
        for indx, item in enumerate(data):
            tree.move(item[1], '', indx)
    
        # switch the heading so that it will sort in the opposite direction
        tree.heading(col,
            command=lambda col=col: self.sortby(tree, col, int(not descending)))
        
    
       
class JanelaAlteraProduto:
    """
    Responsavel por gerar as telas que permitem selecionar um produto dentro da treeview.
    """
    def __init__(self, janelaraiz, codBarras):
        
        # Copiando os dados pegos no banco de dados
        self.listadeprodutos = listaProdOg
        self.novoCodBarras = codBarras
        
        self.janelaraiz = janelaraiz
        
        # Criando Janela TopLevel
        self.janela = Toplevel()
        self.janela.title("Cadastrar Codigo de Barras")
        self.janela.resizable(FALSE,FALSE)
        self.janela.protocol("WM_DELETE_WINDOW",self.FechaJanela)  # Colocando a função de fechar a janela
        #self.janelaetiquetas.focus_set()
        self.Forca_Focus()
        
        # Pegando posição da janela raiz e da toplevel
        #possjnEtiq =  self.jnEtiquetas.geometry().split('+')
        
        possRaiz = self.janelaraiz.geometry().split('+')       
        possjnEtiquetas = str("800x300") + "+" + str(possRaiz[1]) + "+" + str(int(possRaiz[-1])+int(possRaiz[0].split('x')[1])+50)
        
        
        # Posicionado janela abaixo da janela de menu
        self.janela.geometry(possjnEtiquetas)
        
        
        #############################
        #   Primeira linha da grid
        #############################
        
        # Codigo produto
        Label(self.janela, text="Cod. Produto").grid(row=0, column=0)
        self.codproduto = Entry(self.janela, name="codprod",)
        self.codproduto["width"] = 10
        self.codproduto.grid(row=0, column=1)
        
        # Setando o evento do teclado para a descricao, quando qualwuer tecla for presionada chamda filtraProdutos
        self.codproduto.bind("<KeyRelease>", self.filtraProdutos)
        
        # Descricão
        Label(self.janela, text="Descrição").grid(row=0, column=2)
        self.descricao = Entry(self.janela, name="descricao")
        self.descricao["width"] = 50
        self.descricao.grid(row=0, column=3)
        self.descricao.focus_force()
        
        # Setando o evento do teclado para a descricao, quando qualwuer tecla for presionada chamda filtraProdutos
        self.descricao.bind("<KeyRelease>", self.filtraProdutos)
        
        # Codigo de Barras
        Label(self.janela, text="Cod. Barras").grid(row=0, column=4)
        self.codbarra = Entry(self.janela, name="codbarras")
        self.codbarra["width"] = 20
        self.codbarra.grid(row=0, column=5)
        
        # Setando o evento do teclado para o Cod. Barras, quando qualquer tecla for presionada chamda filtraProdutos
        #self.codbarra.bind("<KeyRelease>", self.filtraProdutos)
        #self.codbarra.bind("<Enter>", self.filtraProdutos)
        self.codbarra.bind("<FocusIn>", self.filtraProdutos)
        
        # Quantidades
        Label(self.janela, text="Saldo").grid(row=0, column=6)
        self.saldo = Entry(self.janela, name="saldo")
        self.saldo["width"] = 10
        self.saldo.grid(row=0, column=7)
        
        
        # Setando o evento do teclado para o Cod. Barras, quando qualwuer tecla for presionada chamda filtraProdutos
        #self.saldo.bind("<KeyRelease>", self.filtraProdutos)
        #self.saldo.bind("<Enter>", self.filtraProdutos)
        self.saldo.bind("<FocusIn>", self.filtraProdutos)
        
        
        
        ################################
        #   Fim Primeira linha da grid #
        ################################
        
        ##############################
        #   Segunda linha da grid    #
        ##############################
        self.treeview = self.CriaTreeView(self.janela, numlinhas=1, numcolunas=0, columnspan=8, rowspan=10)
        self.treeview._name =  "treeview",
        self.treeview.grid(row=1, column=0, columnspan=8, rowspan=10)
        self.treeview.bind("<Double-Button-1>", self.SelecionaProduto)
        self.treeview.bind("<Return>", self.SelecionaProduto)
        # Setando o evento do teclado para a descricao, quando qualwuer tecla for presionada chamda filtraProdutos
        self.treeview.bind("<KeyRelease>", self.eventTreeV)
        # Colocando dados na treeview
        self.cadastraListaProdutos(self.listadeprodutos, self.treeview)
        
        ###############################
        #  Fim Segunda linha da grid  #
        ###############################
        
        ###############################
        #  Quarta linha da grid       #
        ###############################
        
        self.jnetbuttoncancelar = Button(self.janela, text="Cancelar", command=self.FechaJanela)
        self.jnetbuttoncancelar.grid(row=12, column=2)
        
        self.btncadastrar = Button(self.janela, text="Cadastrar", command=self.CadastraNovoProd)
        self.btncadastrar.grid(row=12, column=4)
        
        ###############################
        #  Fim Quarta linha da grid   #
        ###############################
        
    def CadastraNovoProd(self):
        """
        Abre a janela JanelaProduto para cadastrar um novo produto
        """
        novoProd = []
        while True:
            print "Abrindo "
            self.janelaNovoProd = JanelaProduto(magicnumber=6)
            self.janelaNovoProd.acodbarraprod.set(self.novoCodBarras)
            if novoProd:
                self.janelaNovoProd.adescprod.set(novoProd[1])
            self.janelaraiz.wait_window(self.janelaNovoProd.alteraJn)
            
            if self.janelaNovoProd.cancela is False :
                
                novoProd = self.janelaNovoProd.get()
                prod = Produto(0, novoProd[1].upper(), novoProd[2], 0)
                aux = True
                for i in self.listadeprodutos.produtos:
                    if i.descricao == prod.descricao:
                        tkMessageBox.showinfo(
                            title="Atenção",
                            message="Já existe um produto com este nome!"
                            )
                        aux = False
                        break
                if aux:
                    self.listadeprodutos.produtos.append(prod)
                    self.janela.destroy()
                    self.janela = None
                    break
                    
            else:
                break
                
                
        
        
        #self.janelaNovoProd.
        
        
    def cadastraListaProdutos(self, listaProdutos, treeV):
        # Coloca dados na tabela
        #print "Tamanho", len(listaProdutos)
        if listaProdutos == []:
            print "Lista Vazia"
            self.labelTotEtiq.set("0")
            self.labelTotPag.set("0")
            #######TODO:XXX Aqui
        else:
            #x = []
           
            for item in listaProdutos.produtos:
                #a = (item.cod_produto, item.descricao, item.cod_barras, item.saldo_estoque)
                #x.append(a)
                treeV.insert('', 'end', values=(item.cod_produto, item.descricao, item.cod_barras, item.saldo_estoque))
                # Calculando o total de etiquetas
               
                
            #x.reverse()
            
            
            
            
            #for y in x:
            #    treeV.insert('', 'end', values=y)
    
    def CriaTreeView(self, frame, numlinhas, numcolunas, columnspan, rowspan):
        '''
        Cria um treeview para exibir os produtos, pede como parametro o frame onde sera colocado
        numlinhas = Linha incial da treeview
        numcolunas = Coluna inicial da treeview
        columspan = qtd de colunas ocupadas pela treeview
        rowspan = qtd de linhas ocupadas pela treeview
        '''
        pcolum = numcolunas+columnspan  # ultima coluna ocupada pela treeview
        plinhas = numlinhas+rowspan  # ultima linha ocupada pela treeview
        
        
        # Titulos das colunas na treeview
        colunas = (
                'Cod. Produto', 
                'Descricao                                                                    ', 
                'Cod. Barras       ', 
                'Quantidade')
        
        treeV = ttk.Treeview(frame, columns=colunas, show="headings", takefocus=1)  
        
        # Scrollbar Vertical
        vsb = ttk.Scrollbar(frame, orient="vertical", command=treeV.yview)
        # Scroollbar Horizontal
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=treeV.xview)
        
        # Configurando Scrollbars na treeview
        treeV.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        treeV.grid(column=numcolunas, row=numlinhas, sticky='nsew')
        
        # a possição da barra de rolagem recebe a possição inicial do grid + o seu tamanho mais 1
        # Tanto para a linha como para a coluna
        vsb.grid(column=pcolum, row=numlinhas, sticky='ns', rowspan=rowspan)
        hsb.grid(column=numcolunas, row=plinhas, sticky='ew', columnspan=columnspan)
        
        # Coloca titulos nas colunas, falta criar funçao de ordenação
        for col in colunas:
                treeV.heading(col, text=col.title(),
                    command=lambda c=col: self.sortby(treeV, c, 0))
                
                # Altera o tamanho da coluna para o tamanho do titulo
                treeV.column(col, width=tkFont.Font().measure(col.title()))
        
        
        return treeV
    
    def FechaJanela(self):
        if tkMessageBox.askokcancel("Exit?", "Você tem certeza que deseja fechar?"):
            self.janela.destroy()
            self.janela = None
        else:
            self.Forca_Focus()
            
            
    def filtraProdutos(self, event):
        """
        Utilizada para filtrar os produtos na treeview.
        """
        # Transformando em maiuscula
        
        text = event.widget.get().upper()
        event.widget.delete(0,END)
        event.widget.insert(0,text)
        
        
        if event.widget._name == "descricao":  # Entry de descricao da janela etiquetas
            
            texto = self.descricao.get()
            self.codproduto.delete(0, END)
            
            # Caso seja seta para cima ou para baixo não entra 
            if (event.keycode != 116) and (event.keycode != 111):
                print "Keycode: ", event.keycode, " Name: ", event.widget._name
                
                if texto != "": # Texto vazio não faz nada
                    # Apaga toda a treeview para ser inserida uma nova
                    x = self.treeview.get_children() 
                    for item in x: 
                        self.treeview.delete(item)
                    
                    if len(texto) <= 3:
                        regex = "^"+texto+"(.*)"
                        prod = listaProdutos()
                        prod.produtos = self.listadeprodutos.FiltraProdutos(regex)
                        #for i in prod:
                        #    print i.descricao
                        self.cadastraListaProdutos(prod,self.treeview)
                        
                    else:                        
                        # Criando a expresão regular para a pesquisa
                        regex = ''
                        # Explodindo texto para fazer a pesquisa por expresão regular
                        textoSplit = texto.split()
                        for te in textoSplit:
                                regex += '(.*)'+te
                            
                        regex += '(.*)'
                        prod = listaProdutos()
                        prod.produtos = self.listadeprodutos.FiltraProdutos(regex, adj="descricao")
                        #for i in prod:
                        #    print i.descricao
                        self.cadastraListaProdutos(prod,self.treeview)
                        self.descricao.focus_force()
                        
                
                else:
                    #self.cadastraListaProdutos(self.produtosselecionados,self.jnettreeview)
                    #inicio = time()
                    self.cadastraListaProdutos(self.listadeprodutos, self.treeview)
                    #fimvazio = time()
                    #print "Fim Vazio", fimvazio - inicio
                
            elif event.keycode == 116:
                # Como a seta para baixo foi apertada seleciona o primeiro produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.treeview.get_children()                 
                self.treeview.selection("set", x[0])
                self.treeview.focus(x[0])
                self.treeview.see(x[0])
                self.treeview.focus_force()
            elif event.keycode == 111:
                # Como a seta para cima foi apertada seleciona o ultimo produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.treeview.get_children()                 
                self.treeview.selection("set", x[-1])
                self.treeview.focus(x[-1])
                self.treeview.see(x[-1])
                self.treeview.focus_force()
                    
                    
        elif event.widget._name == "codprod": # Codigo do produto
            texto =  self.codproduto.get()
            self.descricao.delete(0,END)
            
            # Caso seja seta para cima ou para baixo não entra 
            if (event.keycode != 116) and (event.keycode != 111):
                if texto != "": # Texto vazio não faz nada
                    # Apaga toda a treeview para ser inserida uma nova
                    x = self.treeview.get_children() 
                    for item in x: 
                        self.treeview.delete(item)
                        
                    regex = "^"+texto+"(.*)"
                    prod = listaProdutos()
                    prod.produtos = self.listadeprodutos.FiltraProdutos(regex, adj="codprod")
                    self.cadastraListaProdutos(prod,self.treeview)
                    
                else:
                    #self.cadastraListaProdutos(self.produtosselecionados,self.jnettreeview)
                    self.cadastraListaProdutos(self.listadeprodutos, self.treeview)
                    
            elif event.keycode == 116:
                # Como a seta para baixo foi apertada seleciona o primeiro produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.treeview.get_children()                 
                self.treeview.selection("set", x[0])
                self.treeview.focus(x[0])
                self.treeview.see(x[0])
                self.treeview.focus_force()
            elif event.keycode == 111:
                # Como a seta para cima foi apertada seleciona o ultimo produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.treeview.get_children()                 
                self.treeview.selection("set", x[-1])
                self.treeview.focus(x[-1])
                self.treeview.see(x[-1])
                self.treeview.focus_force()
        
        elif event.widget._name == "codbarras":
            self.descricao.focus_force()
        elif event.widget._name == "saldo":
            self.descricao.focus_force()
        
            
    def Forca_Focus(self):
        """
        Força o focus na janela etiquetas
        """
        
        self.janela.focus_set()
        self.janela.grab_set()            
        self.janela.focus_set()
    
    def eventTreeV(self, event):
        """
        Gerencia os eventos do teclado na treeview
        """
        if event.widget._name[0] == "treeview":
            if event.keycode != 116 and event.keycode != 111:
                # Forçando o focus na descricao
                self.descricao.delete(0, END)
                self.descricao.insert(0, event.char.upper())
                self.descricao.focus_force()
                

    def SelecionaProduto(self, event):
        """
        Pega o produto que foi selecionado na treeview e envia para a JanelaProduto
        """
        #Pega valores do item
        item = self.treeview.item(self.treeview.selection())
        
        
        valores = item.values()
        # Verifica se dados não estão vazio
        if valores[2] != '':
            if tkMessageBox.askokcancel("Alterar Codigo de Barras?", "Deseja vincular o cod. barras " \
                                        + str(self.novoCodBarras) + " \n ao produto " + str(valores[2][1].encode("utf-8"))):
                prod = Produto(valores[2][0], valores[2][1], self.novoCodBarras, valores[2][3])
                print "Prod: ", prod.cod_produto, ", ", prod.descricao
                self.listadeprodutos.AlteraProduto(prod)
                
                # Apaga toda a treeview para ser inserida uma nova
                x = self.treeview.get_children() 
                for item in x: 
                    self.treeview.delete(item)
                    
                self.cadastraListaProdutos(self.listadeprodutos, self.treeview)
                
                self.janela.destroy()
                self.janela = None
                
                
        #self.treeview.focus_set()
    
    def sortby(self, tree, col, descending):
        '''
        Organiza o conteudo da treeview baseado na coluna em que houve o click.
        '''
        # grab values to sort
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
    
        # reorder data
        data.sort(reverse=descending)
        for indx, item in enumerate(data):
            tree.move(item[1], '', indx)
    
        # switch the heading so that it will sort in the opposite direction
        tree.heading(col,
            command=lambda col=col: self.sortby(tree, col, int(not descending)))
    
    
class JanelaProduto:
    """
    Classe responsavel por gerar as janela de exibição e alteração de produto
    """
    
    def __init__(self, magicnumber=15):
        """
        Recebe como parametro o 'magic number', este numero representa qual campo esta disponivel para
        edição. Por padrão todos os campos estão disponiveis, magicnumber=15.
        1 = Cod. Produto pode ser editado
        2 = Cod. Baras pode ser editado
        4 = Descricao pode ser editado
        8 = Saldo Estoque pode ser editado
        A somantoria dos valores passa se mais de um campo pode ser editado.
        Ex:
        magicnumber = 13 Os seguintes itens podem ser editados, Saldo Estoque, Descricao, Cod. Produto        
        """
        
        self.magicnumber = magicnumber
        
        self.cancela = False  # Usado quando a foi fechado sem salvar ou o botão cancelar foi apertado
        #Pega valores do item
        #item = self.tree.item(self.tree.selection())
        
        #Pegando dados do produto selecionado
        #valores = item.values()
        valores = []
        valores = [1, 2, 3]
        
        if valores[2] != '':
            #prodSelecionado = Produto(valores[2][0], valores[2][1], valores[2][2], valores[2][3])
            self.acodprod = StringVar()
            self.adescprod = StringVar()
            self.acodbarraprod = StringVar()
            self.aqtdprod = StringVar()
            
            """self.acodprod.set(valores[2][0])
            self.adescprod.set(valores[2][1])
            self.acodbarraprod.set(valores[2][2])
            self.aqtdprod.set(valores[2][3])"""
            # Criando Janela TopLevel
            self.alteraJn = Toplevel()
            self.alteraJn.title("Alterar Qtd de Produtos")
            
            # Centralizando janela
            # Falta fazer
            
            
            Label(self.alteraJn, text="Cod. Produto").grid(row=0)
            self.alteraJn.protocol("WM_DELETE_WINDOW",self.FechaJanela)  # Colocando a função de fechar a janela
            self.eCodProd = Entry(self.alteraJn, textvariable=self.acodprod)
            self.eCodProd.grid(row=0, column=1)
            #self.eCodProd.bind("<Return>", self.salvaNovaQtd)
            
            # Forçando o focus na janela, e bloqueando nela
            self.alteraJn.grab_set()            
            self.alteraJn.focus_set()
            
            Label(self.alteraJn, text="Cod. Barras").grid(row=0, column=2)
            self.eCodBarras = Entry(self.alteraJn, textvariable=self.acodbarraprod)
            self.eCodBarras.grid(row=0, column=3)
            #eCodBarras.bind("<Return>", self.salvaNovaQtd)
            
            Label(self.alteraJn, text="Descricao").grid(row=1, column=0)
            self.eDescricao = Entry(self.alteraJn, textvariable=self.adescprod)
            self.eDescricao['width'] = 52
            self.eDescricao.grid(row=1, column=1, columnspan=3)
            #eDescricao.bind("<Return>", self.salvaNovaQtd)
            
            Label(self.alteraJn, text="Quantidade").grid(row=2, column=0)
            self.eQtd = Entry(self.alteraJn, textvariable=self.aqtdprod)
            self.eQtd.grid(row=2, column=1)
            #eQtd.bind("<Return>", self.salvaNovaQtd)
            
            
            self.eBtnSalvar = Button(self.alteraJn, text="Salvar", command=self.Salvar)
            self.eBtnSalvar.grid(row=3, column=0, columnspan=2)
            self.eBtnSalvar.focus_force()
            self.eBtnCancelar = Button(self.alteraJn, text="Cancelar",
                                  command=self.FechaJanela)
            self.eBtnCancelar.grid(row=3, column=2, columnspan=2)
            
            self.alteraJn.resizable(FALSE,FALSE)  # Bloqueando alteração de tamanho
            self.CalculaMagicNumber()
            
    def CalculaMagicNumber(self):
        """
        Leva em conta o magic number passado na inicialização para tornanr os campos editaveis ou não.
        Utiliza alteração de focus para fazer o entry não ser editavel
        """
        #self.magicnumber = 14
        resto = int(self.magicnumber)
        
        # Se o magic number for menor ou igual a 15 e maior que zero, calcula
        if (self.magicnumber <= 15) and (self.magicnumber >= 1):
            """
            if (int(resto)/8 == 1):
                print "Quantidade e editavel"
            resto = int(resto)%8
            if (int(resto)/4 == 1):
                print "Descricao e editavel"
            resto = int(resto)%4
            if (int(resto)/2 == 1):
                print "Cod. Barras e editavel"
            resto = int(resto)%2
            if (int(resto)/1 == 1):
                print "Cod. Produtos e editavel"
             """ 
            # Retira focus caso não seja editavel
            
            if (int(resto)/8 != 1):
                self.eQtd.bind("<FocusIn>", self.ForcaFocusSalvar)
            resto = int(resto)%8
            if (int(resto)/4 != 1):
                self.eDescricao.bind("<FocusIn>", self.ForcaFocusSalvar)
            resto = int(resto)%4
            if (int(resto)/2 != 1):
                self.eCodBarras.bind("<FocusIn>", self.ForcaFocusSalvar)
            resto = int(resto)%2
            if (int(resto)/1 != 1):
                self.eCodProd.bind("<FocusIn>", self.ForcaFocusSalvar)
            
    def FechaJanela(self):
        if tkMessageBox.askokcancel("Exit?", "Você tem certeza que deseja fechar?"):
            self.cancela = True
            self.alteraJn.destroy()
            self.alteraJn = None
        else:
            self.Forca_Focus()
            
    def Forca_Focus(self):
        """
        Força o focus na janela etiquetas
        """
        self.alteraJn.focus_set()
        self.alteraJn.grab_set()            
        self.alteraJn.focus_set()
        
    def ForcaFocusSalvar(self, event):
        """
        Força o focus no botão salvar
        """
        #print "chegou"
        #self.alteraJn.focus_get().focus_force()
        self.eBtnSalvar.focus_force()
        
    def get(self):
        """
        Retorna os valores alterados na janela de alteração.
        """
        return (self.acodprod.get(), self.adescprod.get(), self.acodbarraprod.get(), self.aqtdprod.get())
        
    def Salvar(self):
        self.cancela = False
        self.alteraJn.destroy()
        self.alteraJn = None
        
    def set(self, codprod, descricao, codbarras, qtd):
        """
        Coloca as informações na tela
        """
        self.acodprod.set(codprod)
        self.adescprod.set(descricao)
        self.acodbarraprod.set(codbarras)
        self.aqtdprod.set(qtd)
        
        
    def SetPossition(self, raiz, w=385, h=90):
        """
        Centraliza a janela no centro da tela
        """
        ws = raiz.winfo_screenwidth()
        hs = raiz.winfo_screenheight()
        
        # Calculando posição x, y
        x = (ws/2)-(w/2)
        y = (hs/2)-(h/2)
        
        self.alteraJn.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
            


class JanelaEtiquetas:
    """
    Classe responsavel por gerar as janelas de etiquetas
    """
    
    def __init__(self, janelaraiz):
        
        # Criando lista de produtos
        #self.listadeprodutos = listaProdutos()        
        #self.listadeprodutos.PegaListadoBanco()
        self.listadeprodutos = deepcopy(listaProdOg)
        self.janelaraiz = janelaraiz
        
        # Criando lista de produtos selecionados para criar as etiquetas
        self.produtosselecionados = listaProdutos()
        
        # Criando variaveis p/ label com o total de etiquetas e paginas
        self.labelTotEtiq = StringVar()
        self.labelTotPag = StringVar()
        
        self.labelTotEtiq.set("0")
        self.labelTotPag.set("0")
        
        
        
        
        # Criando Janela TopLevel
        self.janelaetiquetas = Toplevel()
        self.janelaetiquetas.title("Impressão de Etiquetas")
        self.janelaetiquetas.resizable(FALSE,FALSE)
        self.janelaetiquetas.protocol("WM_DELETE_WINDOW",self.FechaJanela)  # Colocando a função de fechar a janela
        #self.janelaetiquetas.focus_set()
        self.Forca_Focus()
        
        # Pegando posição da janela raiz e da toplevel
        #possjnEtiq =  self.jnEtiquetas.geometry().split('+')
        possRaiz = self.janelaraiz.geometry().split('+')       
        possjnEtiquetas = str("970x590") + "+" + str(possRaiz[1]) + "+" + str(int(possRaiz[-1])+int(possRaiz[0].split('x')[1])+50)
        
        
        # Posicionado janela abaixo da janela de menu
        self.janelaetiquetas.geometry(possjnEtiquetas)
        
        
        #############################
        #   Primeira linha da grid
        #############################
        
        # Codigo produto
        Label(self.janelaetiquetas, text="Cod. Produto").grid(row=0, column=0)
        self.jnetcodproduto = Entry(self.janelaetiquetas, name="jnetcodprod")
        self.jnetcodproduto["width"] = 10
        self.jnetcodproduto.grid(row=0, column=1)
        
        # Setando o evento do teclado para a descricao, quando qualwuer tecla for presionada chamda filtraProdutos
        self.jnetcodproduto.bind("<KeyRelease>", self.filtraProdutos)
        
        # Descricão
        Label(self.janelaetiquetas, text="Descrição").grid(row=0, column=2)
        self.jnetdescricao = Entry(self.janelaetiquetas, name="jnetdescricao")
        self.jnetdescricao["width"] = 50
        self.jnetdescricao.grid(row=0, column=3)
        self.jnetdescricao.focus_force()
        
        # Setando o evento do teclado para a descricao, quando qualwuer tecla for presionada chamda filtraProdutos
        self.jnetdescricao.bind("<KeyRelease>", self.filtraProdutos)
        
        # Codigo de Barras
        Label(self.janelaetiquetas, text="Cod. Barras").grid(row=0, column=4)
        self.jnetcodbarra = Entry(self.janelaetiquetas, name="jnetcodbarras")
        self.jnetcodbarra["width"] = 20
        self.jnetcodbarra.grid(row=0, column=5)
        
        # Setando o evento do teclado para o Cod. Barras, quando qualquer tecla for presionada chamda filtraProdutos
        self.jnetcodbarra.bind("<KeyRelease>", self.filtraProdutos)
        self.jnetcodbarra.bind("<FocusIn>", self.filtraProdutos)
        
        # Quantidades
        Label(self.janelaetiquetas, text="Saldo").grid(row=0, column=6)
        self.jnetsaldo = Entry(self.janelaetiquetas, name="jnetsaldo")
        self.jnetsaldo["width"] = 10
        self.jnetsaldo.grid(row=0, column=7)
        
        
        # Setando o evento do teclado para o Cod. Barras, quando qualwuer tecla for presionada chamda filtraProdutos
        self.jnetsaldo.bind("<KeyRelease>", self.filtraProdutos)
        self.jnetsaldo.bind("<FocusIn>", self.filtraProdutos)
        
        
        
        ################################
        #   Fim Primeira linha da grid #
        ################################
        
        ##############################
        #   Segunda linha da grid    #
        ##############################
        self.jnettreeview = self.CriaTreeView(self.janelaetiquetas, numlinhas=1, numcolunas=0, columnspan=8, rowspan=10)
        self.jnettreeview.grid(row=1, column=0, columnspan=8, rowspan=10)
        self.jnettreeview._name = "treeview"
        self.jnettreeview.bind("<Double-Button-1>", self.SelecionaProduto)
        self.jnettreeview.bind("<Return>", self.SelecionaProduto)
        self.jnettreeview.bind("<KeyRelease>", self.eventTreeV)
        
        ###############################
        #  Fim Segunda linha da grid  #
        ###############################
        
        ###############################
        #  Terceira linha da grid     #
        ###############################
        """
        # Codigo produto
        Label(self.janelaetiquetas, text="Cod. Produto").grid(row=13, column=0)
        self.jnetcodproduto2 = Entry(self.janelaetiquetas)
        self.jnetcodproduto2["width"] = 10
        self.jnetcodproduto2.grid(row=13, column=1)
        
        # Descricão
        Label(self.janelaetiquetas, text="Descrição").grid(row=13, column=2)
        self.jnetdescricao2 = Entry(self.janelaetiquetas)
        self.jnetdescricao2["width"] = 50
        self.jnetdescricao2.grid(row=13, column=3)
        
        # Codigo de Barras
        Label(self.janelaetiquetas, text="Cod. Barras").grid(row=13, column=4)
        self.jnetcodbarra2 = Entry(self.janelaetiquetas)
        self.jnetcodbarra2["width"] = 20
        self.jnetcodbarra2.grid(row=13, column=5)
        
        # Quantidades
        Label(self.janelaetiquetas, text="Saldo").grid(row=13, column=6)
        self.jnetsaldo2 = Entry(self.janelaetiquetas)
        self.jnetsaldo2["width"] = 10
        self.jnetsaldo2.grid(row=13, column=7)
        """
        
        ###############################
        # Fim Terceira linha da grid  #
        ###############################
        
        ##############################
        #   Quarta linha da grid     #
        ##############################
        self.jnettreeview2 = self.CriaTreeView(self.janelaetiquetas, numlinhas=14, numcolunas=0, columnspan=8, rowspan=10)
        self.jnettreeview2.grid(row=14, column=0, columnspan=8, rowspan=10)
        self.jnettreeview2.bind("<Double-Button-1>", self.AlteraQTD)
        self.jnettreeview2.bind("<Return>", self.AlteraQTD)
        
        ###############################
        #  Fim Quarta linha da grid   #
        ###############################
        
        ###############################
        #  Quinta linha da grid       #
        ###############################
        Label(self.janelaetiquetas, text="Gerar Etiquetas apartir da ", font=("Helvetica", 10)).grid(row=25, column=0)
        Label(self.janelaetiquetas, text="Total de Etiquetas ", font=("Helvetica", 10)).grid(row=25, column=4, sticky=E)
        Label(self.janelaetiquetas, text="Total de Folhas ", font=("Helvetica", 10)).grid(row=26, column=4, sticky=E)
        
        # Criando labels para total de pagianas
        Label(self.janelaetiquetas, textvariable=self.labelTotEtiq, font=("Helvetica", 10, "bold")).grid(row=25, column=5, sticky=W)
        Label(self.janelaetiquetas, textvariable=self.labelTotPag, font=("Helvetica", 10, "bold")).grid(row=26, column=5, sticky=W)
        
        Label(self.janelaetiquetas, text="Linha: ").grid(row=26, column=0, sticky=E)
        self.jnetlinha = Entry(self.janelaetiquetas)
        self.jnetlinha["width"] = 10
        self.jnetlinha.insert(0, "1")
        self.jnetlinha.grid(row=26, column=1, sticky=W)
        
        Label(self.janelaetiquetas, text="Coluna: ").grid(row=26, column=2, sticky=E)
        self.jnetcoluna = Entry(self.janelaetiquetas)
        self.jnetcoluna["width"] = 10
        self.jnetcoluna.insert(0, "1")
        self.jnetcoluna.grid(row=26, column=3, sticky=W)
        
        ###############################
        #  Fim Quinta linha da grid   #
        ###############################
        
        ###############################
        #  Sexta linha da grid       #
        ###############################
        self.jnetbuttongeraretq = Button(self.janelaetiquetas, text="Gerar Etiquetas", command=self.GerarEtiquetas)
        self.jnetbuttongeraretq.grid(row=27, column=2)
        
        self.jnetbuttoncancelar = Button(self.janelaetiquetas, text="Cancelar", command=self.FechaJanela)
        self.jnetbuttoncancelar.grid(row=27, column=4)
        
        ###############################
        #  Fim Sexta linha da grid   #
        ###############################
        
        # Colocando dados na treeview
        self.cadastraListaProdutos(self.listadeprodutos, self.jnettreeview)
        
    def cadastraListaProdutos(self, listaProdutos, treeV):
        # Coloca dados na tabela
        #print "Tamanho", len(listaProdutos)
        if listaProdutos == []:
            print "Lista Vazia"
            self.labelTotEtiq.set("0")
            self.labelTotPag.set("0")
            #######TODO:XXX Aqui
        else:
            #x = []
           
            for item in listaProdutos.produtos:
                #a = (item.cod_produto, item.descricao, item.cod_barras, item.saldo_estoque)
                #x.append(a)
                treeV.insert('', 'end', values=(item.cod_produto, item.descricao, item.cod_barras, item.saldo_estoque))
                # Calculando o total de etiquetas
               
                
            #x.reverse()
            
            
            
            
            #for y in x:
            #    treeV.insert('', 'end', values=y)
        
    def AlteraQTD(self, event):
        self.alteraqtdproduto = JanelaProduto(magicnumber=8)
        self.alteraqtdproduto.SetPossition(self.janelaraiz)
        
        #Pega valores do item
        item = self.jnettreeview2.item(self.jnettreeview2.selection())
        
        valores = item.values()
        # Verifica se dados não estão vazio
        if valores[2] != '':
            novoValor = (valores[2][0], valores[2][1], valores[2][2], valores[2][3])
            self.alteraqtdproduto.set(
                codprod=novoValor[0],
                descricao=novoValor[1],
                codbarras=novoValor[2],
                qtd=novoValor[3])
            #self.produtosselecionados.adicionaProdutos(valores[2][0],valores[2][1], valores[2][2], valores[2][3])
            #apagaTreeView(self.jnettreeview2)
            #self.cadastraListaProdutos(self.produtosselecionados, self.jnettreeview2)
            
            # Espera a janela ser fechada para continuar
            #self.janelaraiz.wait_window(self.alteraqtdproduto.alteraJn)
            self.janelaetiquetas.wait_window(self.alteraqtdproduto.alteraJn)
            
            if self.alteraqtdproduto.cancela is False :
                novoProd = self.alteraqtdproduto.get()
                
                prod = Produto(cod_produto=str(novoProd[0]), descricao=str(novoProd[1]), cod_barras=str(novoProd[2]), saldo_estoque=int(novoProd[3]))
                
                if prod.saldo_estoque <= 0:
                    
                    self.produtosselecionados.RemoveProduto(prod, confirmar=False)
                else:                
                    self.produtosselecionados.AlteraProduto(prod)
                
                apagaTreeView(self.jnettreeview2)
                self.cadastraListaProdutos(self.produtosselecionados, self.jnettreeview2)
                
                totalprod = self.produtosselecionados.ContaTotalProdutos()
                # Setando o total de etiquetas no label
                self.labelTotEtiq.set(totalprod)
                # Calculando e setando o total de paginas
                totalpag = 0
                if totalprod == 0:
                    self.labelTotPag.set(totalpag)              
                
                elif totalprod%27 > 0:
                    totalpag = totalprod/27+1
                    self.labelTotPag.set(totalpag)
                else:
                    totalpag = totalprod/27
                    self.labelTotPag.set((totalprod/27))
                
                
                
                
            
        
        
    def CriaTreeView(self, frame, numlinhas, numcolunas, columnspan, rowspan):
        '''
        Cria um treeview para exibir os produtos, pede como parametro o frame onde sera colocado
        numlinhas = Linha incial da treeview
        numcolunas = Coluna inicial da treeview
        columspan = qtd de colunas ocupadas pela treeview
        rowspan = qtd de linhas ocupadas pela treeview
        '''
        pcolum = numcolunas+columnspan  # ultima coluna ocupada pela treeview
        plinhas = numlinhas+rowspan  # ultima linha ocupada pela treeview
        
        
        # Titulos das colunas na treeview
        colunas = (
                'Cod. Produto', 
                'Descricao                                                                    ', 
                'Cod. Barras       ', 
                'Quantidade')
        
        treeV = ttk.Treeview(frame, columns=colunas, show="headings", takefocus=1)  
        
        # Scrollbar Vertical
        vsb = ttk.Scrollbar(frame, orient="vertical", command=treeV.yview)
        # Scroollbar Horizontal
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=treeV.xview)
        
        # Configurando Scrollbars na treeview
        treeV.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        treeV.grid(column=numcolunas, row=numlinhas, sticky='nsew')
        
        # a possição da barra de rolagem recebe a possição inicial do grid + o seu tamanho mais 1
        # Tanto para a linha como para a coluna
        vsb.grid(column=pcolum, row=numlinhas, sticky='ns', rowspan=rowspan)
        hsb.grid(column=numcolunas, row=plinhas, sticky='ew', columnspan=columnspan)
        
        # Coloca titulos nas colunas, falta criar funçao de ordenação
        for col in colunas:
                treeV.heading(col, text=col.title(),
                    command=lambda c=col: self.sortby(treeV, c, 0))
                
                # Altera o tamanho da coluna para o tamanho do titulo
                treeV.column(col, width=tkFont.Font().measure(col.title()))
        
        
        return treeV
    
    def eventTreeV(self, event):
        """
        Gerencia os eventos do teclado na treeview
        """
        print "Evento do teclado: ", event.keycode
        if event.widget._name == "treeview":
            if event.keycode != 40 and event.keycode != 38:
                # Forçando o focus na descricao
                
                if event.keycode != 13:
                
                    self.jnetcodproduto.delete(0, END)
                    self.jnetdescricao.delete(0, END)
                    self.jnetdescricao.insert(0, event.char.upper())
                    
                self.jnetdescricao.focus_force()
                
        
    def Forca_Focus(self):
        """
        Força o focus na janela etiquetas
        """
        self.janelaetiquetas.focus_set()
        
    def FechaJanela(self):
        if tkMessageBox.askokcancel("Exit?", "Você tem certeza que deseja fechar?"):
            self.janelaetiquetas.destroy()
            self.janelaetiquetas = None
            
        else:
            self.Forca_Focus()
            
    def filtraProdutos(self, event):
        """
        Utilizada para filtrar os produtos na treeview.
        """
        
        # Transformando em maiuscula
        text = event.widget.get().upper()
        event.widget.delete(0,END)
        event.widget.insert(0,text)
        if event.widget._name == "jnetdescricao":  # Entry de descricao da janela etiquetas
            
            texto = self.jnetdescricao.get()
            
            # Caso seja seta para cima ou para baixo não entra 
            if (event.keycode != 38) and (event.keycode != 40):
                
                
                if texto != "": # Texto vazio não faz nada
                    # Apaga toda a treeview para ser inserida uma nova
                    x = self.jnettreeview.get_children() 
                    for item in x: 
                        self.jnettreeview.delete(item)
                    
                    if len(texto) <= 3:
                        regex = "^"+texto+"(.*)"
                        prod = listaProdutos()
                        prod.produtos = self.listadeprodutos.FiltraProdutos(regex)
                        #for i in prod:
                        #    print i.descricao
                        self.cadastraListaProdutos(prod,self.jnettreeview)
                        
                    else:                        
                        # Criando a expresão regular para a pesquisa
                        regex = ''
                        # Explodindo texto para fazer a pesquisa por expresão regular
                        textoSplit = texto.split()
                        for te in textoSplit:
                                regex += '(.*)'+te
                            
                        regex += '(.*)'
                        prod = listaProdutos()
                        prod.produtos = self.listadeprodutos.FiltraProdutos(regex, adj="descricao")
                        #for i in prod:
                        #    print i.descricao
                        self.cadastraListaProdutos(prod,self.jnettreeview)
                        
                
                else:
                    #self.cadastraListaProdutos(self.produtosselecionados,self.jnettreeview)
                    self.cadastraListaProdutos(self.listadeprodutos, self.jnettreeview)
                    
            elif event.keycode == 40:
                # Como a seta para baixo foi apertada seleciona o primeiro produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.jnettreeview.get_children()                 
                self.jnettreeview.selection("set", x[0])
                self.jnettreeview.focus(x[0])
                self.jnettreeview.see(x[0])
                self.jnettreeview.focus_force()
            elif event.keycode == 38:
                # Como a seta para cima foi apertada seleciona o ultimo produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.jnettreeview.get_children()                 
                self.jnettreeview.selection("set", x[-1])
                self.jnettreeview.focus(x[-1])
                self.jnettreeview.see(x[-1])
                self.jnettreeview.focus_force()
                
                    
        elif event.widget._name == "jnetcodprod": # Codigo do produto
            texto = self.jnetcodproduto.get()
            
            # Caso seja seta para cima ou para baixo não entra 
            if (event.keycode != 38) and (event.keycode != 40):
                if texto != "": # Texto vazio não faz nada
                    # Apaga toda a treeview para ser inserida uma nova
                    x = self.jnettreeview.get_children() 
                    for item in x: 
                        self.jnettreeview.delete(item)
                        
                    regex = "^"+texto+"(.*)"
                    prod = listaProdutos()
                    prod.produtos = self.listadeprodutos.FiltraProdutos(regex, adj="codprod")
                    self.cadastraListaProdutos(prod,self.jnettreeview)
                    
                else:
                    #self.cadastraListaProdutos(self.produtosselecionados,self.jnettreeview)
                    self.cadastraListaProdutos(self.listadeprodutos, self.jnettreeview)
                    
            elif event.keycode == 40:
                # Como a seta para baixo foi apertada seleciona o primeiro produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.jnettreeview.get_children()                 
                self.jnettreeview.selection("set", x[0])
                self.jnettreeview.focus(x[0])
                self.jnettreeview.see(x[0])
                self.jnettreeview.focus_force()
            elif event.keycode == 38:
                # Como a seta para cima foi apertada seleciona o ultimo produto da treeview
                # torna ele visivel e forca o focus nela
                x = self.jnettreeview.get_children()                 
                self.jnettreeview.selection("set", x[-1])
                self.jnettreeview.focus(x[-1])
                self.jnettreeview.see(x[-1])
                self.jnettreeview.focus_force()
        
        elif event.widget._name == "jnetcodbarras":
            self.jnetdescricao.focus_force()
        elif event.widget._name == "jnetsaldo":
            self.jnetdescricao.focus_force()
            
    def GerarEtiquetas(self):
        """
        Chama a função para gerar o pdf com as etiquetas passadas.
        """
        
        if self.produtosselecionados.ContaTotalProdutos() > 0:
            barCodeVali = True
            # Percorendo lista de produtos para criar os codigos de barra
            for prod in self.produtosselecionados.produtos:
                bar = EanBarCode()
                # Convert code string in integer list
                EAN13 = []
                for digit in str(prod.cod_barras):
                   EAN13.append(int(digit))
                   
                # If the code has already a checksum
                if len(EAN13) == 13:
                    # Verify checksum
                    if bar.verifyChecksum(EAN13):
                        bar.getImage(EAN13,50,"png")
                    else:
                        tkMessageBox.showinfo(
                            title="Atenção",
                            message="Agum o produto " + prod.descricao + " esta com Cod. de Barras invalido!"
                            )
                        barCodeVali = False
                        break
                else:
                    tkMessageBox.showinfo(
                            title="Atenção",
                            message="Agum o produto " + prod.descricao + " esta com Cod. de Barras invalido!"
                            )
                    
                    barCodeVali = False
                    break
                  
            # Gerando imagens dos cod. barras junto com os cod. produto
            if barCodeVali == True:
                
            
                for prod in self.produtosselecionados.produtos:
                    caminhoimg = "Dados/"+str(prod.cod_barras)+".png"
                    GeraImgCodBarrasProd(caminhoimg, str(prod.cod_produto))
                  
                while 1:  
                    # Perguntando onde salvar o arquivo
                    tipo = ""
                    fileName = asksaveasfilename(filetypes=[('PDF Files','*.pdf')])
                    
                    if fileName: #  Caso o botão cancelar não tenha sido selecionado, não entra
                        tipo = fileName.split(".")
                        
                        # Verifica se nome não esta em branco
                        if tipo[0] == "":
                            continue
                        # Forçando para ser um arquivo .csv
                        if tipo[-1] != "pdf":
                            fileName += ".pdf"
                        try:
                            temp = open(fileName, "w")
                        except:
                            tkMessageBox.showinfo(title="Atenção", message="Arquivo em uso por outro programa ou sem permissão de alteração.\nEscolha Outro!")
                            #print "Arquivo aberto"
                        else:
                            #print "Pode continuar"
                            temp.close()
                            etiquetasPDF = GerarEtiquetasPDF(output=str(fileName))
                            print "Linhas: ", self.jnetlinha.get(), "Colunas: ", self.jnetcoluna.get()
                            if (int(self.jnetlinha.get()) > 0) and (int(self.jnetcoluna.get()) > 0) and (int(self.jnetcoluna.get()) <= 3):
                                
                                print "Linhas: ", self.jnetlinha.get(), "Colunas: ", self.jnetcoluna.get()
                            
                                qtdpre =  ( int(self.jnetlinha.get())*3 ) - ( 3 - int(self.jnetcoluna.get()) )
                                
                                for temp in range(qtdpre-1):
                                    etiquetasPDF.GeraEtiquetasBranco()
                                
                            
                            
                            for prod in self.produtosselecionados.produtos:
                                caminhoimg = "Dados/"+str(prod.cod_barras)+".png"
                                etiquetasPDF.SetDados(prod.descricao, caminhoimg, prod.saldo_estoque)
                                
                                
                            etiquetasPDF.FinalizaPDF()
                            
                            tkMessageBox.showinfo(title="Parabens", message="Geração de Etiquetas realizada com sucesso!")
                            apagaTreeView(self.jnettreeview2)
                            self.produtosselecionados = listaProdutos()
                            self.jnetcoluna.delete(0, END)
                            self.jnetcoluna.insert(0, "1")
                            self.jnetlinha.delete(0, END)
                            self.jnetlinha.insert(0, "1")
                            
                            
                            break
                    else:
                        break
                
        else:
            tkMessageBox.showinfo(title="Atenção", message="Selecione Algum Produto!")
                
         
    def SelecionaProduto(self, event):
        """
        Pega o produto que foi selecionado na treeview e envia para a JanelaProduto
        """
        #Pega valores do item
        item = self.jnettreeview.item(self.jnettreeview.selection())
        
        valores = item.values()
        # Verifica se dados não estão vazio
        if valores[2] != '':
            novoValor = (valores[2][0], valores[2][1], valores[2][2], valores[2][3])
            self.produtosselecionados.adicionaProdutos(valores[2][0],valores[2][1], valores[2][2], valores[2][3])
            apagaTreeView(self.jnettreeview2)
            self.cadastraListaProdutos(self.produtosselecionados, self.jnettreeview2)
            
                
            totalprod = self.produtosselecionados.ContaTotalProdutos()
            # Setando o total de etiquetas no label
            self.labelTotEtiq.set(totalprod)
            # Calculando e setando o total de paginas
            totalpag = 0
            if totalprod == 0:
                self.labelTotPag.set(totalpag)              
            
            elif totalprod%27 > 0:
                totalpag = totalprod/27+1
                self.labelTotPag.set(totalpag)
            else:
                totalpag = totalprod/27
                self.labelTotPag.set((totalprod/27))
            
            #self.jnetcodproduto.delete(0, END)
            #self.jnetdescricao.delete(0,END)
            
    
    def sortby(self, tree, col, descending):
        '''
        Organiza o conteudo da treeview baseado na coluna em que houve o click.
        '''
        # grab values to sort
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
    
        # reorder data
        print "Tipo: ", type(data)
        data.sort(reverse=descending)
        for indx, item in enumerate(data):
            tree.move(item[1], '', indx)
    
        # switch the heading so that it will sort in the opposite direction
        tree.heading(col,
            command=lambda col=col: self.sortby(tree, col, int(not descending)))
    


def main():
    '''
    Função principal do programa
    '''
    # Usado para criar a janela
    janela = JanelaRaiz()
    
    # Cria a barra de menu por icones na parte superior, abaixo do menu
    janela.IconMenu()
    
    # Roda o programa
    janela.mainloop()


if __name__ == "__main__":
    main()
