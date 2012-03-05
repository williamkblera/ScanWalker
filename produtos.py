#!/usr/bin/env python
# -*- coding: utf-8 -*-

import kinterbasdb # Para o acesso ao banco
import tkMessageBox
from re import *

##################################################################
#  Classe de produtos em uma loja com as seguintes informações,  #
#  Cod. Produto, Descrição, Cod. Barras, Saldo em Estoque        #
##################################################################

def apagaTreeView(tree):
    '''
    Apaga os dados da treeview passada
    '''
    x = tree.get_children() 
    for item in x: 
        tree.delete(item)

class Produto:
    '''
    Armazena as informações dos produtos, Codigo do Produto ( cod_produto)
    Codigo de Barras(cod_barras), Descrição/Nome do produto(descricao), 
    quantidade em estoque(saldo_estoque)
    '''
    
    def __len__(self):
        if self.cod_produto == None:
            return 0
    
    def __init__(self, cod_produto=None, descricao=0, cod_barras=0, saldo_estoque=0):
        self.cod_produto = cod_produto # Codigo do produto no sistema
        
        
        self.cod_barras = cod_barras # Codigo de barras do produto
        #self.cod_barras = CodigoBarras()
        #self.cod_barras.setCodBarras(cod_barras)
        
        
        self.descricao = descricao # Descricao do produto
        self.saldo_estoque = saldo_estoque # Quantidade de produtos em estoque
        
    def add(self):
        '''
        Soma mais um produto ao estoque
        '''
        self.saldo_estoque = self.saldo_estoque + 1
        



##################################################################
#  Classe lista de produtos com as seguintes informações,        #
#  Cod. Produto, Descrição, Cod. Barras, Saldo em Estoque        #
##################################################################

class listaProdutos:
    '''
    Classe que gerencia uma lista de produtos.
    '''
    
    def __init__(self):
        # Lista com os produtos
        self.produtos = []
        
    def addProduto(self, produto):
        #prod = self.getInfosCod(produto.cod_produto)
        prod = self.getInfos(produto.descricao)
        
        if prod is None:
            self.adicionaProdutos(produto.cod_produto, produto.descricao, produto.cod_barras, 1)
        else:
            #self.adicionaProdutoCodBarras(prod.cod_barras)
            prod.add()
        
    def adicionaProdutos(self, cod_produto, descricao, cod_barras, saldo_estoque=0):
        '''
        Adiciona um novo produto a lista, mas não soma nada.
        '''
        
        produto = Produto(cod_produto, descricao, cod_barras, saldo_estoque)
        # Verifica se lista esta vazia
        #print len(self.produtos)
        a = 0
       
        if self.produtos == []:
            # Como lista está vazia adiciona o primeiro produto
            
            self.produtos.append(produto)
            
            
        else:
            
            for prod in self.produtos:
                # Busca valor para ver se o item esta na lista
                if prod.descricao == produto.descricao:
                #prod.add()
                # Caso ache um produto igual marca que o produto ja esta na lista
                    return False # Produto já cadastrado
            # Como produto não existe, adiciona produto
            self.produtos.append(produto)
            return True # Produto adiciona com sucesso
            
    def AlteraProduto(self, produto):
        """
        Altera os dados de um produto.
        """
        print produto.saldo_estoque
        for prod in self.produtos:
            if produto.descricao == prod.descricao:
                prod.cod_produto = produto.cod_produto
                prod.cod_barras = produto.cod_barras
                prod.saldo_estoque = produto.saldo_estoque
                break
        
        
    def buscaProduto(self, produto):
        '''
        Busca o produto passado na lista de produtos
        '''
        
        for i in self.produtos:
            if (i.cod_produto == produto.cod_produto) and (i.descricao == produto.descricao):
                return True
            
        return False
    
    def ContaTotalProdutos(self):
        """
        Retorna o total de produtos na lista
        """
        cont = 0
        for i in self.produtos:
            cont = cont + int(i.saldo_estoque)
        return cont
    
    def FiltraProdutos(self, regex, adj="descricao"):
        """
        Faz ama busca por expressão regular na lista, usando o 'regex' como filtro e buscando no
        adjetivo (adj) passado. Retorna uma nova lista somente com os produtos filtrados
        adj="codprod" busca no codigo do produto
        adj="descricao" busca na descrição do produto
        adj="codbarras" busca no codigo de barras
        adj="saldo" busca no saldo
        """
        if adj == "descricao":
            prod = []
            for i in self.produtos:
                #print "Descricao: ", i.descricao
                
                if(match(regex,i.descricao,IGNORECASE|U)):
                    #prod.adicionaProdutos(prod.cod_produto, prod.descricao, prod.cod_barras, prod.saldo_estoque)
                    prod.append(i)
                    #print "acho"
                
                #prod.append(i)
            return prod
        
        elif adj == "codprod":
            prod = []
            for i in self.produtos:
                if(match(regex,str(i.cod_produto),IGNORECASE|U)):
                    prod.append(i)
                
                #prod.append(i)
            return prod
        
    def getInfos(self, descricao):
        """
        Busca produto pela descricao
        """
        for prod in self.produtos:
            if prod.descricao == descricao:
                return prod
        return None
    
    
    def getInfosCodBarras(self, cod_barras):
        for prod in self.produtos:
            if prod.cod_barras == cod_barras:
                return prod
        return None
    
    def getInfosCod(self, cod):
        for prod in self.produtos:
            if prod.cod_produto == cod:
                return prod
        return None
    
    def PegaListadoBanco(self):
        '''
        Pega os dados do banco de dados e salva na listaProdutos
        '''
        
        try:
            conn = kinterbasdb.connect(host="10.1.1.10",
                                   database='C:\SIACPlus\siacCX.fdb',
                                   user="sysdba",
                                   password="masterkey",
                                   charset="ISO8859_1"
                                   )
            print "Conexão aberta"
            cur = conn.cursor()
            cur.execute('SELECT a.COD_PRODUTO, a.DESCRICAO, a.CODBARRAS, a.SALDOESTOQUE FROM ACADPROD a')
            print "Select executado"
            #print "Lista de Produtos: ", self.listadeProdutos
            # Pegando dados do select e salvando em lista produtos, no momento salvando com estoque zero
            
            for i in cur:
                
                (cod_prod, descr, codbarras, saldoestoque) = i
                
                self.adicionaProdutos(cod_prod, descr, codbarras, saldoestoque)
            
            print "Produtos adicionados a lista"
            
            # Fechando acesso ao banco                
            conn.close()
            print "Conexão fechada"
            
        except:
            print "Erro ao acessar o banco"
        finally:
            print "Dados salvo com sucesso"
            
    def RemoveProduto(self, produto, confirmar=True):
        """
        Remove da lista o produto passado. Caso confirmar esteja em True ele pergunta antes de remover.
        """
        for i in self.produtos:
            if ( str(produto.descricao) == str(i.descricao)):
                if confirmar==True:
                    if tkMessageBox.askokcancel("Remover?", "Você tem certeza que deseja Apagar o item \n" + str(i.descricao)):
                        self.produtos.remove(i)
                else:
                    self.produtos.remove(i)
                #break
            
