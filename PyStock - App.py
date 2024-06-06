import os
import sys
import time
import sqlite3
import datetime
import re

import openpyxl.drawing.image

from tkinter.filedialog import askdirectory
from tkinter import Tk

from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QMessageBox

from View.PY.FrmAdmin import Ui_FrmAdmin
from View.PY.FrmLogin import Ui_login
from openpyxl import *

# Verifique se o arquivo do banco de dados já existe
if os.path.exists('database.db'):
    try:
        # Tente conectar ao banco de dados
        banco = sqlite3.connect('database.db')
        cursor = banco.cursor()
        print("Conexão bem-sucedida ao banco de dados.")

    except sqlite3.DatabaseError as e:
        print("Erro ao conectar ao banco de dados:", e)
        print("Criando um novo banco de dados...")
        os.remove('database.db')  # Remove o arquivo existente
        banco = sqlite3.connect('database.db')
        cursor = banco.cursor()
else:
    print("Criando um novo banco de dados...")
    banco = sqlite3.connect('database.db')
    cursor = banco.cursor()

# Crie tabelas para o banco de dados SQLite
cursor.execute("CREATE TABLE IF NOT EXISTS `clientes` (`CPF` TEXT, `Nome` TEXT, `Endereço` TEXT, `Contato` TEXT, `saldo_devedor` TEXT);")
cursor.execute("CREATE TABLE IF NOT EXISTS `fornecedores` (`Nome` TEXT, `Endereço` TEXT, `Contato` TEXT);")
cursor.execute("CREATE TABLE IF NOT EXISTS `login` (`usuario` TEXT, `senha` TEXT, `nivel` TEXT, `nome` TEXT);")
cursor.execute("CREATE TABLE IF NOT EXISTS `monitoramento_vendas` (`vendedor` TEXT, `cliente` TEXT, `qtde_vendido` TEXT, `total_venda` TEXT, `horario_venda` TEXT, `cpf_da_venda` TEXT,`a_prazo` INTEGER);")
cursor.execute("CREATE TABLE IF NOT EXISTS `monitoramento_compras` (`comprador` TEXT, `fornecedor` TEXT, `qtde_comprado` TEXT, `total_compra` TEXT, `horario_compra` TEXT);")
cursor.execute("CREATE TABLE IF NOT EXISTS `produtos` (`cód_produto` TEXT, `descrição` TEXT, `valor_unitário` TEXT, `qtde_estoque` TEXT, `fornecedor` TEXT, `valor_de_custo` TEXT);")
cursor.execute("CREATE TABLE IF NOT EXISTS `quem_vendeu_mais` (`nome` TEXT, `total_qtde` TEXT);")
cursor.execute("CREATE TABLE IF NOT EXISTS `vendas` (`cód` TEXT, `produto` TEXT, `valor_unitário` TEXT, `qtde` TEXT, `total` TEXT, `id` INTEGER);")
cursor.execute("CREATE TABLE IF NOT EXISTS `compras` (`cód` TEXT, `produto` TEXT, `valor_de_custo` TEXT, `qtde` TEXT, `total` TEXT, `id` INTEGER);")


cursor.execute("SELECT COUNT(*) FROM 'login'")
result = cursor.fetchone()[0]
if result == 0:
    cursor.execute("INSERT INTO `login` (`usuario`, `senha`, `nivel`, `nome`) VALUES ('1', '1', 'admin', 'Renato');")

# Comite as transações para garantir que elas sejam concluídas
banco.commit()

class FrmLogin(QMainWindow):

    def __init__(self):

        QMainWindow.__init__(self)

        self.ui = Ui_login()
        self.ui.setupUi(self)

        # Botão de logar no sistema
        self.ui.pushButton.clicked.connect(lambda: self.logar())

    def logar(self):

        global window, UserLogado

        # Pegando os colaboradores cadastrados no banco
        cursor.execute("SELECT * FROM login")
        logins = cursor.fetchall()

        # Pegando as informações inseridas
        usuario = self.ui.lineEdit.text()
        senha = self.ui.lineEdit_2.text()

        # Verificando cada Usuário
        for login in logins:

            if usuario != login[0]:
                self.ui.lineEdit.setStyleSheet('background-color: rgba(0, 0 , 0, 0);border: 2px solid rgba(0,0,0,0);'
                                               'border-bottom-color: rgb(255, 17, 49);color: rgb(0,0,0);padding-bottom: 8px;'
                                               'border-radius: 0px;font: 10pt "Montserrat";')

            if senha != login[1]:
                self.ui.lineEdit_2.setStyleSheet('background-color: rgba(0, 0 , 0, 0);border: 2px solid rgba(0,0,0,0);'
                                                 'border-bottom-color: rgb(255, 17, 49);color: rgb(0,0,0);padding-bottom: 8px;'
                                                 'border-radius: 0px;font: 10pt "Montserrat";')

            # Caso os dados estejam no banco, é iniciado o Frm de acordo com o nivel
            if usuario == login[0] and senha == login[1]:

                UserLogado = login[3]

                self.ui.lineEdit.setStyleSheet('background-color: rgba(0, 0 , 0, 0);border: 2px solid rgba(0,0,0,0);'
                                               'border-bottom-color: rgb(20, 71, 224);color: rgb(0,0,0);padding-bottom: 8px;'
                                               'border-radius: 0px;font: 10pt "Montserrat";')

                self.ui.lineEdit_2.setStyleSheet('background-color: rgba(0, 0 , 0, 0);border: 2px solid rgba(0,0,0,0);'
                                                 'border-bottom-color: rgb(20, 71, 224);color: rgb(0,0,0);padding-bottom: 8px;'
                                                 'border-radius: 0px;font: 10pt "Montserrat";')

                window.close()
                window = FrmAdmin()
                window.show()
                break


class FrmAdmin(QMainWindow):

    def __init__(self):
        global filtro, search_fornecedores, window

        QMainWindow.__init__(self)

        self.ui = Ui_FrmAdmin()
        self.ui.setupUi(self)

        # Nome do Usuário
        self.ui.lbl_seja_bem_vindo.setText(f'Seja Bem-Vindo(a) - {UserLogado}')
        self.ui.lbl_titulo_vendas.setText(f'Vendedor(a) - {UserLogado}')
        self.ui.lbl_seja_bem_vindo.setFixedWidth(500)

        # Configurando páginas e os botões do menu

        # Home
        self.ui.btn_home.clicked.connect(lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_home))

        # Colaboradores
        self.ui.btn_colaboradores.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_colaboradores))
        self.ui.btn_cadastrar_colaboradores.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_cadastro_colaboradores))
        self.ui.btn_alterar_colaboradores.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_alterar_colaboradores))

        self.ui.btn_cadastro_colaboradores.clicked.connect(self.CadastroColaboradores)
        self.ui.btn_finalizar_alterar_colaboradores.clicked.connect(self.AlterarColaboradores)

        self.ui.line_senha_alterar_colaboradores.setEchoMode(QLineEdit.EchoMode.Password)
        self.ui.btn_exluir_colaboradores.clicked.connect(self.ExcluirColaboradores)
        self.ui.tabela_alterar_colaboradores.doubleClicked.connect(self.setTextAlterarColaboradores)

        # Botões para ver/esconder senha inserida
        self.ui.btn_ver_senha_cadastro_colaboradores.clicked.connect(self.VerSenhaCadastroColaboradores)
        self.ui.btn_ver_senha_alterar_colaboradores.clicked.connect(self.VerSenhaAlterarColaboradores)

        # Tabela pg_colaboradores
        self.ui.tabela_colaboradores.setColumnWidth(0, 260)
        self.ui.tabela_colaboradores.setColumnWidth(1, 260)
        self.ui.tabela_colaboradores.setColumnWidth(2, 260)

        # Tabela alterar_colaboradores
        self.ui.tabela_alterar_colaboradores.setColumnWidth(0, 330)
        self.ui.tabela_alterar_colaboradores.setColumnWidth(1, 330)
        self.ui.tabela_alterar_colaboradores.setColumnWidth(2, 330)

        # Monitoramento vendas
        self.ui.btn_vendas_monitoramento.clicked.connect(lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_monitoramento_vendas))

        # Tabela Monitoramento vendas
        self.ui.tabela_monitoramento_vendas.setColumnWidth(0, 156)
        self.ui.tabela_monitoramento_vendas.setColumnWidth(1, 156)
        self.ui.tabela_monitoramento_vendas.setColumnWidth(2, 156)
        self.ui.tabela_monitoramento_vendas.setColumnWidth(3, 156)
        self.ui.tabela_monitoramento_vendas.setColumnWidth(4, 156)

        # Monitoramento compras
        self.ui.btn_compras_monitoramento.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_monitoramento_compras))

        # Tabela Monitoramento compras
        self.ui.tabela_monitoramento_compras.setColumnWidth(0, 156)
        self.ui.tabela_monitoramento_compras.setColumnWidth(1, 156)
        self.ui.tabela_monitoramento_compras.setColumnWidth(2, 156)
        self.ui.tabela_monitoramento_compras.setColumnWidth(3, 156)
        self.ui.tabela_monitoramento_compras.setColumnWidth(4, 156)

        # Clientes
        self.ui.btn_clientes.clicked.connect(lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_clientes))
        self.ui.btn_cadastrar_clientes_clientes.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_cadastrar_clientes))
        self.ui.btn_alterar_clientes_clientes.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_alterar_clientes))
        self.ui.btn_finalizar_cadastro_clientes.clicked.connect(self.CadastrarClientes)
        self.ui.btn_exclui_clientes.clicked.connect(self.ExcluirClientes)
        self.ui.tabela_alterar_clientes.doubleClicked.connect(self.setTextAlterarClientes)
        self.ui.btn_finalizar_alteracao_alterar_clientes.clicked.connect(self.AlterarClientes)

        # Tabela Clientes
        self.ui.tabela_clientes.setColumnWidth(0, 170)
        self.ui.tabela_clientes.setColumnWidth(1, 170)
        self.ui.tabela_clientes.setColumnWidth(2, 170)
        self.ui.tabela_clientes.setColumnWidth(3, 170)
        self.ui.tabela_clientes.setColumnWidth(3, 160)

        # Tabela Cadastrar Clientes
        self.ui.tabela_cadastrar_clientes.setColumnWidth(0, 247)
        self.ui.tabela_cadastrar_clientes.setColumnWidth(1, 247)
        self.ui.tabela_cadastrar_clientes.setColumnWidth(2, 247)
        self.ui.tabela_cadastrar_clientes.setColumnWidth(3, 249)

        # Tabela Alterar Clientes
        self.ui.tabela_alterar_clientes.setColumnWidth(0, 247)
        self.ui.tabela_alterar_clientes.setColumnWidth(1, 247)
        self.ui.tabela_alterar_clientes.setColumnWidth(2, 247)
        self.ui.tabela_alterar_clientes.setColumnWidth(3, 249)

        # Vendas
        self.ui.btn_vendas.clicked.connect(lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_vendas))

        # Tabela Vendas
        self.ui.tabela_vendas.setColumnWidth(0, 50)
        self.ui.tabela_vendas.setColumnWidth(1, 131)
        self.ui.tabela_vendas.setColumnWidth(2, 250)
        self.ui.tabela_vendas.setColumnWidth(3, 131)
        self.ui.tabela_vendas.setColumnWidth(4, 75)
        self.ui.tabela_vendas.setColumnWidth(5, 155)

        # Compras
        self.ui.btn_compras.clicked.connect(lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_compras))

        # Tabela compras
        self.ui.tabela_vendas_carregamento.setColumnWidth(0, 50)
        self.ui.tabela_vendas_carregamento.setColumnWidth(1, 131)
        self.ui.tabela_vendas_carregamento.setColumnWidth(2, 250)
        self.ui.tabela_vendas_carregamento.setColumnWidth(3, 131)
        self.ui.tabela_vendas_carregamento.setColumnWidth(4, 75)
        self.ui.tabela_vendas_carregamento.setColumnWidth(5, 155)

        # Fornecedores
        self.ui.btn_fornecedores.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_fornecedores))
        self.ui.btn_adicionar_forncedores.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_cadastrar_fornecedores))
        self.ui.btn_editar_fornecedores.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_alterar_fornecedores))
        self.ui.btn_cadastrar_forncedores.clicked.connect(self.CadastrarFornecedores)
        self.ui.tabela_alterar_fornecedores.doubleClicked.connect(self.setTextAlterarFornecedores)
        self.ui.btn_alterar_fornecedores.clicked.connect(self.AlterarFornecedores)
        self.ui.btn_excluir_fornecedores.clicked.connect(self.ExluirFornecedores)

        # Tabela Fornecedores
        self.ui.tabela_fornecedores.setColumnWidth(0, 257)
        self.ui.tabela_fornecedores.setColumnWidth(1, 257)
        self.ui.tabela_fornecedores.setColumnWidth(2, 257)

        # Tabela Cadastrar Fornecedores
        self.ui.tabela_cadastrar_fornecedores.setColumnWidth(0, 330)
        self.ui.tabela_cadastrar_fornecedores.setColumnWidth(1, 330)
        self.ui.tabela_cadastrar_fornecedores.setColumnWidth(2, 330)

        # Tabela Alterar Fornecedores
        self.ui.tabela_alterar_fornecedores.setColumnWidth(0, 330)
        self.ui.tabela_alterar_fornecedores.setColumnWidth(1, 330)
        self.ui.tabela_alterar_fornecedores.setColumnWidth(2, 330)

        # Produtos
        self.ui.btn_produtos.clicked.connect(lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_produtos))
        self.ui.btn_cadastrar_produto.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_cadastar_produtos))
        self.ui.btn_alterar_produto.clicked.connect(
            lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_alterar_produtos))
        self.ui.btn_finalizar_cadastro_produtos.clicked.connect(self.CadastrarProdutos)
        self.ui.btn_excluir_produto.clicked.connect(self.ExcluirProdutos)
        self.ui.tabela_alterar_produto.doubleClicked.connect(self.setTextAlterarProdutos)
        self.ui.btn_finalizar_alterar_produto.clicked.connect(self.AlterarProdutos)

        # Tabela Produtos
        self.ui.tabela_produto.setColumnWidth(0, 50)
        self.ui.tabela_produto.setColumnWidth(1, 50)
        self.ui.tabela_produto.setColumnWidth(2, 250)
        self.ui.tabela_produto.setColumnWidth(3, 105)
        self.ui.tabela_produto.setColumnWidth(4, 75)
        self.ui.tabela_produto.setColumnWidth(5, 155)
        self.ui.tabela_produto.setColumnWidth(6, 105)

        # Tabela Cadastrar Produtos
        self.ui.tabela_cadastro_produto.setColumnWidth(0, 50)
        self.ui.tabela_cadastro_produto.setColumnWidth(1, 50)
        self.ui.tabela_cadastro_produto.setColumnWidth(2, 280)
        self.ui.tabela_cadastro_produto.setColumnWidth(3, 145)
        self.ui.tabela_cadastro_produto.setColumnWidth(4, 75)
        self.ui.tabela_cadastro_produto.setColumnWidth(5, 250)
        self.ui.tabela_cadastro_produto.setColumnWidth(6, 145)


        # Tabela Alterar Produtos
        self.ui.tabela_alterar_produto.setColumnWidth(0, 50)
        self.ui.tabela_alterar_produto.setColumnWidth(1, 50)
        self.ui.tabela_alterar_produto.setColumnWidth(2, 280)
        self.ui.tabela_alterar_produto.setColumnWidth(3, 145)
        self.ui.tabela_alterar_produto.setColumnWidth(4, 75)
        self.ui.tabela_alterar_produto.setColumnWidth(5, 250)
        self.ui.tabela_alterar_produto.setColumnWidth(6, 145)

        # Configurações
        self.ui.btn_configs.clicked.connect(lambda: self.ui.Telas_do_menu.setCurrentWidget(self.ui.pg_configuracoes))

        # Voltar
        self.ui.btn_voltar.clicked.connect(self.Voltar)

        # Atualizando Tabelas
        self.AtualizaTabelasLogin()
        self.AtualizaTabelasClientes()
        self.AtualizaTabelasFornecedores()
        self.AtualizaTabelasProdutos()
        self.AtualizaTabelaVendas()
        self.AtualizaTabelaCompras()

        # iniciando Hora e Data do Sistema
        tempo = QTimer(self)
        tempo.timeout.connect(self.HoraData)
        tempo.timeout.connect(self.Sair)
        tempo.start(1000)

        # Atuliza previsão de texto das barras de pesquisa
        self.AtualizaCompleterSearchFornecedores()
        self.AtualizaCompleterSearchProdutos()
        self.AtualizaCompleterSearchColaboradores()
        self.AtualizaCompleterSearchClientes()
        self.AtualizaCompleterSearchVendas()
        # self.AtualizaCompleterSearchCompras() implementar

        # Conectando com a função fazer a pesquisa dos dados inseridos

        # Produtos
        self.ui.btn_pesquisar_produto.clicked.connect(lambda: self.SearchProdutos(pg='Produtos'))
        self.ui.line_search_Bar_produtos.returnPressed.connect(lambda: self.SearchProdutos(pg='Produtos'))

        self.ui.line_search_Bar_alterar_produto.returnPressed.connect(lambda: self.SearchProdutos(pg='Alterar'))
        self.ui.btn_pesquisar_alterar_produto.clicked.connect(lambda: self.SearchProdutos(pg='Alterar'))

        self.ui.line_search_Bar_cadastrar_produto.returnPressed.connect(lambda: self.SearchProdutos(pg='Cadastrar'))
        self.ui.btn_pesquisar_cadastrar_produto.clicked.connect(lambda: self.SearchProdutos(pg='Cadastrar'))

        # Fornecedores
        self.ui.btn_filtrar_fornecedores.clicked.connect(lambda: self.SearchFornecedores(pg='Fornecedores'))
        self.ui.line_search_Bar_fornecedores.returnPressed.connect(lambda: self.SearchFornecedores(pg='Fornecedores'))

        self.ui.btn_flitrar_alterar_fornecedores.clicked.connect(lambda: self.SearchFornecedores(pg='Alterar'))
        self.ui.line_search_Bar_altarar_fornecedor.returnPressed.connect(lambda: self.SearchFornecedores(pg='Alterar'))

        self.ui.btn_pesquisar_cadastrar_fornecedores.clicked.connect(lambda: self.SearchFornecedores(pg='Cadastrar'))
        self.ui.line_search_Bar_cadastrar_fornecedores.returnPressed.connect(
            lambda: self.SearchFornecedores(pg='Cadastrar'))

        # Colaboradores
        self.ui.btn_pesquisar_colaboradores.clicked.connect(lambda: self.SearchColaboradores(pg='Colaboradores'))
        self.ui.line_search_bar_colaboradores.returnPressed.connect(
            lambda: self.SearchColaboradores(pg='Colaboradores'))

        self.ui.btn_pesquisar_alterar_colaboradores.clicked.connect(lambda: self.SearchColaboradores(pg='Alterar'))
        self.ui.line_search_bar_buscar_alterar_colaboradores.returnPressed.connect(
            lambda: self.SearchColaboradores(pg='Alterar'))

        # Clientes
        self.ui.btn_pesquisar_clientes.clicked.connect(lambda: self.SearchClientes(pg='Clientes'))
        self.ui.line_search_Bar_clientes.returnPressed.connect(lambda: self.SearchClientes(pg='Clientes'))

        self.ui.btn_filtrar_alterar_clientes.clicked.connect(lambda: self.SearchClientes(pg='Alterar'))
        self.ui.line_search_Bar_alterar_clientes.returnPressed.connect(lambda: self.SearchClientes(pg='Alterar'))

        self.ui.btn_pesquisar_cadastro_clientes.clicked.connect(lambda: self.SearchClientes(pg='Cadastrar'))
        self.ui.line_search_Bar_cadastrar_clientes.returnPressed.connect(lambda: self.SearchClientes(pg='Cadastrar'))

        self.ui.line_cpf_cadastrar_clientes.setMaxLength(14)
        self.ui.line_contato_cadastrar_clientes.setMaxLength(16)

        self.ui.line_alterar_cpf_alterar_clientes.setMaxLength(14)
        self.ui.line_alterar_contato_alterar_clientes.setMaxLength(16)

        self.ui.line_cadastrar_contato_fornecedores.setMaxLength(16)
        self.ui.line_alterar_contato_fornecedor.setMaxLength(16)

        # Formatando número de contato dos Fornecedores
        self.ui.line_cadastrar_contato_fornecedores.textChanged.connect(
            lambda: self.FormataNumeroContato(pg='CadastrarFornecedores'))
        self.ui.line_alterar_contato_fornecedor.textChanged.connect(
            lambda: self.FormataNumeroContato(pg='AlterarFornecedores'))

        # Formatando número de contato dos Clientes
        self.ui.line_contato_cadastrar_clientes.textChanged.connect(lambda: self.FormataNumeroContato(pg='CadastrarClientes'))
        self.ui.line_alterar_contato_alterar_clientes.textChanged.connect(lambda: self.FormataNumeroContato(pg='AlterarClientes'))

        # Formatando CPF dos Clientes
        self.ui.line_cpf_cadastrar_clientes.textChanged.connect(lambda: self.FormataCPFClientes(pg='Cadastrar'))
        self.ui.line_alterar_cpf_alterar_clientes.textChanged.connect(lambda: self.FormataCPFClientes(pg='Alterar'))
        self.ui.line_cliente.textChanged.connect(lambda: self.FormataCPFClientes(pg='Vendas'))
        self.ui.line_xls_cpf_clientes.textChanged.connect(lambda: self.FormataCPFClientes(pg='clientes'))
        self.ui.line_xls_cpf_clientes.setMaxLength(14)

        # Formatando o valor de produtos
        self.ui.line_valor_cadastrar_produto_produto.textChanged.connect(lambda: self.FormataValorProduto(pg='Cadastrar'))
        self.ui.line_valorcusto_cadastrar_produto.textChanged.connect(lambda: self.FormataValorProduto(pg= 'CadastrarValorCusto'))
        self.ui.line_valor_alterar_produto.textChanged.connect(lambda: self.FormataValorProduto(pg='Alterar'))
        self.ui.line_valorcusto_alterar_produto.textChanged.connect(lambda: self.FormataValorProduto(pg= 'AlterarValorCusto'))


        # Formatando valor desconto
        self.ui.line_desconto_vendas.setMaxLength(3)

        # Pesquisando produto pelo código
        self.ui.line_codigo_produto.returnPressed.connect(self.PesquisandoProdutoPeloCodigo)
        self.ui.btn_confirmar_codigo.clicked.connect(self.PesquisandoProdutoPeloCodigo)
        self.ui.line_search_bar_vendas.returnPressed.connect(self.CodProdutoVendas)

        # # Confirmando cliente informado na pg Vendas
        # self.ui.line_cliente.returnPressed.connect(self.ConfirmarCliente)

        # Vendas
        self.ui.btn_adicionar_compra.clicked.connect(self.CadastrandoVendas)
        self.ui.btn_excluir_item.clicked.connect(self.ExcluirVenda)

        # Compras
        self.ui.btn_adicionar_carregamento.clicked.connect(self.CadastrandoCompras) #implementar
        self.ui.btn_excluir_item_carregamento.clicked.connect(self.ExcluirVenda)
        self.ui.btn_finalizar_compra.clicked.connect(self.FinalizarCompras())

        # Ajustando posição do Troco e do Total Vendas
        self.ui.lbl_total_venda.move(670, 20)
        self.ui.lbl_total_valor.move(910, 20)
        self.ui.line_troco.move(680, 80)
        self.ui.lbl_devolver_troco.move(680, 130)
        self.ui.lbl_troco.move(830, 130)

        # Conectando com a função Troco
        self.ui.line_troco.textChanged.connect(self.Troco)
        self.ui.line_troco.returnPressed.connect(self.Troco)

        # Conectando com a função de formatar a data
        self.ui.line_data_monitoramento_vendas.setMaxLength(7)
        self.ui.line_data_monitoramento_vendas.textChanged.connect(self.formatar_data)

        # Atualizando Total
        self.AtualizaTotal()

        # Conectando com função de Finalizar a venda
        self.ui.btn_finalizar_compra.clicked.connect(self.FinalizarVendas)
        self.AtualizaTabelaMonitoramentoVendas()

        # Conectando com a função de Limpar a tabela do monitamento
        self.ui.btn_limpar_tabela_monitoramento_vendas.clicked.connect(self.LimparTabelaMonitoramento)

        # Conectando com a barra de pesquisa do monitoramento
        self.ui.line_search_bar_monitoramentoto_vendas_vendas_vendas.returnPressed.connect(self.SearchMonitoramentoVendas)
        self.ui.btn_filtrar_monitoramento_vendas.clicked.connect(self.SearchMonitoramentoVendas)

        # Conectando com a função de gerar o arquivo xlsx de vendas
        self.ui.btn_gerar_excel_monitoramento_vendas.clicked.connect(self.GerarXls)

        # Conectando com a função de gerar o arquivo xlsx de pendencias
        self.ui.btn_gerarxls_pendencia_clientes.clicked.connect(self.GerarXlsPendencias)

        # Conectando com a função de limpar pendencias
        self.ui.btn_limpar_pendencia_cliente.clicked.connect(self.LimparPendencias)

        # Conectando com a função de gerar o arquivo xlsx de compras
        self.ui.btn_gerar_excel_compras.clicked.connect(self.GerarXlsCompras)

        # Função para encerrar o programa após 20m
        self.ui.btn_salvar.clicked.connect(self.Futuro)
        self.ui.btn_salvar.clicked.connect(self.Sair)

    # Pequenas Funções
    def Voltar(self):
        global window

        window.close()
        window = FrmLogin()
        window.show()

    def HoraData(self):
        tempoAtual = QTime.currentTime()
        tempoTexto = tempoAtual.toString('hh:mm:ss')
        data_atual = datetime.date.today()
        dataTexto = data_atual.strftime('%d/%m/%Y')

        # Colaboradores
        self.ui.lbl_hora_data_colaboradores.setText(f'{dataTexto} {tempoTexto}')
        self.ui.lbl_hora_data_alterar_colaboradores.setText(f'{dataTexto} {tempoTexto}')

        # Monitoramento de Vendas
        self.ui.lbl_hora_data_monitoramento_vendas.setText(f'{dataTexto} {tempoTexto}')

        # Vendas
        self.ui.lbl_hora_data.setText(f'{dataTexto} {tempoTexto}')

        # Produtos
        self.ui.lbl_hora_data_produtos.setText(f'{dataTexto} {tempoTexto}')
        self.ui.lbl_hora_data_alterar_produto.setText(f'{dataTexto} {tempoTexto}')
        self.ui.lbl_hora_data_cadastrar_produto.setText(f'{dataTexto} {tempoTexto}')

        # Fornecedores
        self.ui.lbl_hora_data_fornecedores.setText(f'{dataTexto} {tempoTexto}')
        self.ui.lbl_hora_data_alterar_fornecedores.setText(f'{dataTexto} {tempoTexto}')
        self.ui.lbl_hora_data_cadastrar_fornecedores.setText(f'{dataTexto} {tempoTexto}')

        # Clientes
        self.ui.lbl_hora_data_clientes.setText(f'{dataTexto} {tempoTexto}')
        self.ui.lbl_hora_data_cadastrar_clientes.setText(f'{dataTexto} {tempoTexto}')
        self.ui.lbl_hora_data_alterar_clientes.setText(f'{dataTexto} {tempoTexto}')

    def PesquisandoProdutoPeloCodigo(self):
        produtos = list()

        cod_inserido = self.ui.line_codigo_produto

        cursor.execute('SELECT * FROM produtos')
        banco_produtos = cursor.fetchall()

        produtos.clear()
        tabela = self.ui.tabela_produto

        for produto in banco_produtos:
            produtos.append(produto[0])

        items = tabela.findItems(cod_inserido.text(), Qt.MatchExactly)

        if items:
            item = items[0]
            tabela.setCurrentItem(item)

            cod_inserido.setStyleSheet(StyleNormal)

        else:
            cod_inserido.setStyleSheet(StyleError)

    def CodProdutoVendas(self):
        cursor.execute('SELECT * FROM produtos')
        banco_produtos = cursor.fetchall()

        produto_inserido = self.ui.line_search_bar_vendas

        for pos, produto in enumerate(banco_produtos):
            if produto[1] == produto_inserido.text():
                self.ui.line_codigo_vendas.setText(produto[0])
                break

    def ConfirmarCliente(self):
        global search_clientes

        cliente = self.ui.line_cliente

        if cliente.text() in search_clientes:
            cliente.setStyleSheet('''
                background-color: rgba(0, 0 , 0, 0);
                border: 2px solid rgba(0,0,0,0);
                border-bottom-color: rgb(15, 168, 103);
                color: rgb(0,0,0);
                padding-bottom: 8px;
                border-radius: 0px;
                font: 10pt "Montserrat";''')
        else:
            cliente.setStyleSheet(StyleError)

    def FinalizarVendas(self):
        a_prazo_check = self.ui.radio_venda_prazo
        cliente = self.ui.line_cliente.text()

        cursor.execute('SELECT * FROM vendas')
        banco_vendas = cursor.fetchall()

        cursor.execute('SELECT * FROM quem_vendeu_mais')
        banco_quem_mais_vendeu = cursor.fetchall()

        cursor.execute('SELECT cpf FROM clientes WHERE nome = ?', (cliente,))
        cpf_da_venda = cursor.fetchone()

        tempoAtual = QTime.currentTime()
        tempoTexto = tempoAtual.toString('hh:mm:ss')
        data_atual = datetime.date.today()
        dataTexto = data_atual.strftime('%d/%m/%Y')

        qtde_vendido = list()
        totalVenda = list()
        vendedor = UserLogado
        data_hora = f'{dataTexto} / {tempoTexto}'

        if banco_vendas and cpf_da_venda is not None and a_prazo_check.isChecked():
            a_prazo = 1

            for venda in banco_vendas:
                qtde_vendido.append(int(venda[3]))
                totalVenda.append(float(venda[4]))  # Convertendo para float

            comando_SQL = 'INSERT INTO monitoramento_vendas VALUES (?,?,?,?,?,?,?)'
            dados = f'{vendedor}', f'{cliente}', f'{sum(qtde_vendido)}', f'{sum(totalVenda)}', f'{data_hora}', f'{cpf_da_venda[0]}', f'{a_prazo}'
            cursor.execute(comando_SQL, dados)
            banco.commit()

            cursor.execute('SELECT saldo_devedor FROM clientes WHERE nome = ?', (cliente,))
            saldo_devedor = cursor.fetchone()

            saldo_devedor_total = float(''.join(re.findall(r'\d', saldo_devedor[0]))) / 100 + sum(totalVenda)

            saldo_devedor_convertido = 'R$' + "{:.2f}".format(saldo_devedor_total)

            cursor.execute('UPDATE clientes SET saldo_devedor = ? WHERE nome = ?', (saldo_devedor_convertido, cliente,))
            banco.commit()

            colaboradores = list()
            for colaborador in banco_quem_mais_vendeu:
                colaboradores.append(colaborador[0])

                if colaborador[0] == vendedor:
                    cursor.execute(
                        f'UPDATE quem_vendeu_mais set total_qtde = {int(colaborador[1]) + int(sum(qtde_vendido))} WHERE nome = "{vendedor}"')
                    banco.commit()

            if vendedor not in colaboradores:
                comando_SQL = 'INSERT INTO quem_vendeu_mais VALUES (?,?)'
                dados = f'{vendedor}', f'{sum(qtde_vendido)}'
                cursor.execute(comando_SQL, dados)
            banco.commit()

            for venda in banco_vendas:
                cursor.execute("UPDATE produtos SET qtde_estoque = qtde_estoque - ? WHERE cód_produto = ?",
                               (venda[3], venda[0]))
                banco.commit()
        elif banco_vendas and not a_prazo_check.isChecked():
            a_prazo = 0

            for venda in banco_vendas:
                qtde_vendido.append(int(venda[3]))
                totalVenda.append(float(venda[4]))  # Convertendo para float

            cpf_da_venda = None

            comando_SQL = 'INSERT INTO monitoramento_vendas VALUES (?,?,?,?,?,?,?)'
            dados = f'{vendedor}', f'{cliente}', f'{sum(qtde_vendido)}', f'{sum(totalVenda)}', f'{data_hora}', f'{cpf_da_venda}', f'{a_prazo}'
            cursor.execute(comando_SQL, dados)
            banco.commit()

            colaboradores = list()
            for colaborador in banco_quem_mais_vendeu:
                colaboradores.append(colaborador[0])

                if colaborador[0] == vendedor:
                    cursor.execute(
                        f'UPDATE quem_vendeu_mais set total_qtde = {int(colaborador[1]) + int(sum(qtde_vendido))} WHERE nome = "{vendedor}"')
                    banco.commit()

            if vendedor not in colaboradores:
                comando_SQL = 'INSERT INTO quem_vendeu_mais VALUES (?,?)'
                dados = f'{vendedor}', f'{sum(qtde_vendido)}'
                cursor.execute(comando_SQL, dados)
                banco.commit()

            for venda in banco_vendas:
                cursor.execute("UPDATE produtos SET qtde_estoque = qtde_estoque - ? WHERE cód_produto = ?",
                               (venda[3], venda[0]))
                banco.commit()

        elif cpf_da_venda is None and a_prazo_check.isChecked():
            self.Popup('Vendas', 'Não foi possivel achar o cpf do cliente informado')
        elif not banco_vendas:
            self.Popup('Vendas', 'Nenhum produto informado')

        cursor.execute('DELETE FROM vendas')
        banco.commit()
        self.AtualizaTabelaVendas()
        self.AtualizaTotal()
        self.AtualizaTabelasClientes()
        self.AtualizaTabelasProdutos()
        self.AtualizaTabelaMonitoramentoVendas()
        self.AtualizaCompleterSearchVendas()

        self.ui.line_codigo_vendas.clear()
        self.ui.line_cliente.clear()
        self.ui.line_quantidade_vendas.clear()
        self.ui.line_desconto_vendas.clear()
        self.ui.lbl_troco.setText('0,00')
        self.ui.line_cliente.setStyleSheet(StyleNormal)
        self.ui.line_search_bar_vendas.clear()
        self.ui.line_troco.clear()
        self.ui.line_desconto_vendas.setStyleSheet(StyleNormal)
        self.ui.line_codigo_vendas.setStyleSheet(StyleNormal)
        self.ui.line_quantidade_vendas.setStyleSheet(StyleNormal)

    def FinalizarCompras(self):
        cursor.execute('SELECT * FROM compras')
        banco_compras = cursor.fetchall()

        tempoAtual = QTime.currentTime()
        tempoTexto = tempoAtual.toString('hh:mm:ss')
        data_atual = datetime.date.today()
        dataTexto = data_atual.strftime('%d/%m/%Y')

        qtde_comprado = list()
        totalCompra = list()
        comprador = UserLogado
        data_hora = f'{dataTexto} / {tempoTexto}'

        fornecedor = ''

        if banco_compras:
            for compra in banco_compras:
                qtde_comprado.append(int(compra[3]))
                totalCompra.append(float(compra[4]))  # Convertendo para float

            comando_SQL = 'INSERT INTO monitoramento_compras VALUES (?,?,?,?,?,?,?)'
            dados = f'{comprador}', f'{fornecedor}', f'{sum(qtde_comprado)}', f'{sum(totalCompra)}', f'{data_hora}'
            cursor.execute(comando_SQL, dados)
            banco.commit()

            for venda in banco_compras:
                cursor.execute("UPDATE produtos SET qtde_estoque = qtde_estoque + ? WHERE cód_produto = ?",
                               (venda[3], venda[0]))
                banco.commit()
        elif not banco_compras:
            self.Popup('Compras', 'Nenhum produto informado')

        cursor.execute('DELETE FROM compras')
        banco.commit()
        self.AtualizaTabelaCompras()
        self.AtualizaTotalCompras()
        self.AtualizaTabelasClientes()
        self.AtualizaTabelasProdutos()
        self.AtualizaTabelaMonitoramentoVendas()
        self.AtualizaCompleterSearchVendas()

        self.ui.line_codigo_vendas_carregamento.clear()
        self.ui.line_quantidade_vendas_carregamento.clear()
        self.ui.line_search_bar_vendas_carregamento.clear()
        self.ui.line_codigo_vendas.setStyleSheet(StyleNormal)
        self.ui.line_quantidade_vendas.setStyleSheet(StyleNormal)

    def Troco(self):
        cursor.execute('SELECT * FROM vendas')
        banco_vendas = cursor.fetchall()

        troco_desejado = self.ui.line_troco.text()
        troco_desejado = ''.join(filter(str.isdigit, troco_desejado))

        if troco_desejado.isnumeric():
            vendas = [int(venda[4]) for venda in banco_vendas]
            troco = int(troco_desejado) - sum(vendas)

            troco_formatado = f"{troco / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.ui.lbl_troco.setText(troco_formatado)

            # Update line_troco with formatted value
            troco_desejado_float = int(troco_desejado) / 100
            troco_desejado_formatado = f"{troco_desejado_float:,.2f}".replace(",", "X").replace(".", ",").replace("X",
                                                                                                                  ".")
            self.ui.line_troco.setText(troco_desejado_formatado)
            self.ui.line_troco.setCursorPosition(len(troco_desejado_formatado))

    def AtualizaTotal(self):

        cursor.execute('SELECT * FROM vendas')
        banco_vendas = cursor.fetchall()

        vendas = list()

        for pos, venda in enumerate(banco_vendas):
            vendat = ''.join(re.findall(r'\d', venda[4]))
            vendas.append(int(vendat))

        total = lang.toString(sum(vendas) * 0.01, 'f', 2)
        self.ui.lbl_total_valor.setText(f'{total}')
        self.Troco()

    def AtualizaTotalCompras(self):

        cursor.execute('SELECT * FROM compras')
        banco_compras = cursor.fetchall()

        compras = list()

        for pos, compra in enumerate(banco_compras):
            comprat = ''.join(re.findall(r'\d', compra[4]))
            compras.append(int(comprat))

        total = lang.toString(sum(compras) * 0.01, 'f', 2)
        self.ui.lbl_total_valor_carregamento.setText(f'{total}')
        self.Troco()

    def Futuro(self):
        global futuroTexto
        atual = datetime.datetime.now()

        futuro = atual + datetime.timedelta(minutes=20)
        futuroTexto = futuro.time().strftime('%H:%M:%S')

    def Sair(self):
        global futuroTexto
        tempoAtual = QTime.currentTime()
        tempoTexto = tempoAtual.toString('hh:mm:ss')

        if self.ui.checkBox_finalizar_app.isChecked() == True:
            if tempoTexto == futuroTexto:
                sys.exit()

    def LimparTabelaMonitoramento(self):
        cursor.execute('DELETE FROM monitoramento_vendas')
        cursor.execute('DELETE FROM quem_vendeu_mais')
        self.AtualizaTabelaMonitoramentoVendas()
        self.AtualizaCompleterSearchVendas()

    # Popups
    def Popup(self, titulo, texto):
        msg = QMessageBox()
        msg.setWindowTitle(titulo)
        msg.setText(texto)

        icon = QIcon()
        icon.addPixmap(QPixmap("View/Imagens/Logo Ico.ico"), QIcon.Normal, QIcon.Off)
        msg.setWindowIcon(icon)
        x = msg.exec_()

    def PopupXlsDiretorio(self):
        msg = QMessageBox()
        msg.setWindowTitle("Erro - Gerar Excel")
        msg.setText('Selecione um diretório válido!')

        icon = QIcon()
        icon.addPixmap(QPixmap("View/Imagens/Logo Ico.ico"), QIcon.Normal, QIcon.Off)
        msg.setWindowIcon(icon)
        x = msg.exec_()

    def PopupXls(self):
        msg = QMessageBox()
        msg.setWindowTitle("Erro - Gerar Excel")
        msg.setText('Verifique se não há um ARQUIVO com o mesmo nome aberto!')

        icon = QIcon()
        icon.addPixmap(QPixmap("View/Imagens/Logo Ico.ico"), QIcon.Normal, QIcon.Off)
        msg.setWindowIcon(icon)
        x = msg.exec_()

    def PoupXlsBancoVazio(self):
        msg = QMessageBox()
        msg.setWindowTitle("Erro - Gerar Excel")
        msg.setText('Nenhuma venda informada!')

        icon = QIcon()
        icon.addPixmap(QPixmap("View/Imagens/Logo Ico.ico"), QIcon.Normal, QIcon.Off)
        msg.setWindowIcon(icon)
        x = msg.exec_()

    # Função para formatar a data desejada
    def formatar_data(self):
        text = self.ui.line_data_monitoramento_vendas.text()

        # Remover quaisquer caracteres que não sejam números
        digits_only = ''.join(filter(str.isdigit, text))

        # Adicionar uma barra após os primeiros dois caracteres (mês)
        formatted_text = digits_only[:2] if len(digits_only) >= 2 else digits_only
        if len(digits_only) > 2:
            formatted_text += '/' + digits_only[2:]

        # Atualizar o campo de texto
        self.ui.line_data_monitoramento_vendas.setText(formatted_text)
        self.ui.line_data_monitoramento_vendas.setCursorPosition(len(formatted_text))

    # Função que gera um arquivo xlsx para melhor monitoramento das vendas
    def GerarXls(self):
        global wb

        dataDesejada = self.ui.line_data_monitoramento_vendas.text()

        cursor.execute('SELECT * FROM monitoramento_vendas WHERE SUBSTR(horario_venda, 4, 7) = ?', (dataDesejada,))
        banco_monitoramento = cursor.fetchall()

        if len(banco_monitoramento) > 0:
            Tk().withdraw()
            diretorio = askdirectory()

            if diretorio != '':
                try:
                    wb.save(filename=r'{}\Relatório.xlsx'.format(diretorio))
                except:
                    self.PopupXls()
                else:

                    cursor.execute('SELECT * FROM quem_vendeu_mais')
                    banco_quem_vendeu_mais = cursor.fetchall()

                    total_vendido = 0
                    total_faturado = 0
                    total_clientes_cadastrados = 0
                    total_clientes_não_cadastrados = 0
                    quem_vendeu_mais = list()
                    colaborador = ''

                    planilha = wb['Relatório']

                    c = 18
                    for vendas in banco_monitoramento:
                        total_vendido += int(vendas[2])
                        total_faturado += int(vendas[3])
                        if vendas[1] == 'Não Informado':
                            total_clientes_não_cadastrados += 1
                        else:
                            total_clientes_cadastrados += 1
                        c += 1
                        conv = lang.toString(int(vendas[3]) * 0.01, "f", 2)
                        planilha[f'A{c}'] = vendas[0]
                        planilha[f'E{c}'] = vendas[1]
                        planilha[f'I{c}'] = int(vendas[2])
                        planilha[f'M{c}'] = 'RS ' + conv
                        planilha[f'R{c}'] = vendas[4]

                    for colaboradores in banco_quem_vendeu_mais:
                        quem_vendeu_mais.append(colaboradores[1])

                    for colaboradores in banco_quem_vendeu_mais:
                        if colaboradores[1] == max(quem_vendeu_mais, key=int):
                            colaborador = colaboradores[0]
                    conv = lang.toString(int(total_faturado) * 0.01, "f", 2)

                    planilha['F12'] = total_vendido
                    planilha['F13'] = 'RS ' + conv
                    planilha['F14'] = colaborador
                    planilha['F15'] = total_clientes_cadastrados
                    planilha['F16'] = total_clientes_não_cadastrados

                    wb.save(filename=r'{}\Relatório.xlsx'.format(diretorio))

                    for c in range(19, 19 + len(banco_monitoramento)):
                        planilha[f'A{c}'] = None
                        planilha[f'E{c}'] = None
                        planilha[f'I{c}'] = None
                        planilha[f'M{c}'] = None
                        planilha[f'R{c}'] = None
            else:
                self.PopupXlsDiretorio()
        else:
            self.PoupXlsBancoVazio()

    def GerarXlsPendencias(self):
        global wb

        cpf = self.ui.line_xls_cpf_clientes.text()

        cursor.execute('SELECT cpf FROM clientes WHERE cpf = ?', (cpf,))
        vefCPF = cursor.fetchall()

        if cpf != '' and vefCPF:
            cursor.execute('SELECT * FROM monitoramento_vendas WHERE cpf_da_venda = ? AND a_prazo = 1', (cpf,))
            banco_monitoramento = cursor.fetchall()

            if len(banco_monitoramento) > 0:
                Tk().withdraw()
                diretorio = askdirectory()

                if diretorio != '':
                    try:
                        wb.save(filename=r'{}\Relatório.xlsx'.format(diretorio))
                    except:
                        self.PopupXls()
                    else:

                        cursor.execute('SELECT * FROM quem_vendeu_mais')
                        banco_quem_vendeu_mais = cursor.fetchall()

                        total_vendido = 0
                        total_faturado = 0
                        total_clientes_cadastrados = 0
                        total_clientes_não_cadastrados = 0
                        quem_vendeu_mais = list()
                        colaborador = ''

                        planilha = wb['Relatório']

                        c = 18
                        for vendas in banco_monitoramento:
                            total_vendido += int(vendas[2])
                            total_faturado += int(vendas[3])
                            if vendas[1] == 'Não Informado':
                                total_clientes_não_cadastrados += 1
                            else:
                                total_clientes_cadastrados += 1
                            c += 1
                            conv = lang.toString(int(vendas[3]) * 0.01, "f", 2)
                            planilha[f'A{c}'] = vendas[0]
                            planilha[f'E{c}'] = vendas[1]
                            planilha[f'I{c}'] = int(vendas[2])
                            planilha[f'M{c}'] = 'RS ' + conv
                            planilha[f'R{c}'] = vendas[4]

                        for colaboradores in banco_quem_vendeu_mais:
                            quem_vendeu_mais.append(colaboradores[1])

                        for colaboradores in banco_quem_vendeu_mais:
                            if colaboradores[1] == max(quem_vendeu_mais, key=int):
                                colaborador = colaboradores[0]
                        conv = lang.toString(int(total_faturado) * 0.01, "f", 2)

                        planilha['F12'] = total_vendido
                        planilha['F13'] = 'RS ' + conv
                        planilha['F14'] = colaborador
                        planilha['F15'] = total_clientes_cadastrados
                        planilha['F16'] = total_clientes_não_cadastrados

                        wb.save(filename=r'{}\Relatório.xlsx'.format(diretorio))

                        for c in range(19, 19 + len(banco_monitoramento)):
                            planilha[f'A{c}'] = None
                            planilha[f'E{c}'] = None
                            planilha[f'I{c}'] = None
                            planilha[f'M{c}'] = None
                            planilha[f'R{c}'] = None
                else:
                    self.PopupXlsDiretorio()
            else:
                self.Popup("Erro - Gerar Excel", "Nenhuma pendencia para este cliente")
        else:
            self.Popup("Erro - Gerar Excel", "CPF não cadastrado")

    def LimparPendencias(self):
        cpf = self.ui.line_xls_cpf_clientes.text()
        if cpf != '':
            cursor.execute('UPDATE monitoramento_vendas SET a_prazo = 0 WHERE a_prazo = 1 AND cpf_da_venda = ?',(cpf,))
            banco.commit()

            cursor.execute('UPDATE clientes SET saldo_devedor = "R$ 0,00" WHERE CPF = ?',(cpf,))
            banco.commit()

            self.Popup("Limpar pendencias", "Todas pendencias foram quitadas")
        else:
            self.Popup("Limpar pendencias", "Nenhum CPF informado")

        self.AtualizaTabelasClientes()

    def GerarXlsCompras(self):
        global wb

        dataDesejada = self.ui.line_data_monitoramento_vendas.text()

        cursor.execute('SELECT * FROM monitoramento_vendas WHERE SUBSTR(horario_venda, 4, 7) = ?', (dataDesejada,))
        banco_monitoramento = cursor.fetchall()

        if len(banco_monitoramento) > 0:
            Tk().withdraw()
            diretorio = askdirectory()

            if diretorio != '':
                try:
                    wb.save(filename=r'{}\Relatório.xlsx'.format(diretorio))
                except:
                    self.PopupXls()
                else:

                    cursor.execute('SELECT * FROM quem_vendeu_mais')
                    banco_quem_vendeu_mais = cursor.fetchall()

                    total_vendido = 0
                    total_faturado = 0
                    total_clientes_cadastrados = 0
                    total_clientes_não_cadastrados = 0
                    quem_vendeu_mais = list()
                    colaborador = ''

                    planilha = wb['Relatório']

                    c = 18
                    for vendas in banco_monitoramento:
                        total_vendido += int(vendas[2])
                        total_faturado += int(vendas[3])
                        if vendas[1] == 'Não Informado':
                            total_clientes_não_cadastrados += 1
                        else:
                            total_clientes_cadastrados += 1
                        c += 1
                        conv = lang.toString(int(vendas[3]) * 0.01, "f", 2)
                        planilha[f'A{c}'] = vendas[0]
                        planilha[f'E{c}'] = vendas[1]
                        planilha[f'I{c}'] = int(vendas[2])
                        planilha[f'M{c}'] = 'RS ' + conv
                        planilha[f'R{c}'] = vendas[4]

                    for colaboradores in banco_quem_vendeu_mais:
                        quem_vendeu_mais.append(colaboradores[1])

                    for colaboradores in banco_quem_vendeu_mais:
                        if colaboradores[1] == max(quem_vendeu_mais, key=int):
                            colaborador = colaboradores[0]
                    conv = lang.toString(int(total_faturado) * 0.01, "f", 2)

                    planilha['F12'] = total_vendido
                    planilha['F13'] = 'RS ' + conv
                    planilha['F14'] = colaborador
                    planilha['F15'] = total_clientes_cadastrados
                    planilha['F16'] = total_clientes_não_cadastrados

                    wb.save(filename=r'{}\Relatório.xlsx'.format(diretorio))

                    for c in range(19, 19 + len(banco_monitoramento)):
                        planilha[f'A{c}'] = None
                        planilha[f'E{c}'] = None
                        planilha[f'I{c}'] = None
                        planilha[f'M{c}'] = None
                        planilha[f'R{c}'] = None
            else:
                self.PopupXlsDiretorio()
        else:
            self.PoupXlsBancoVazio()

    # Função para formartar o número de contato inserido
    def FormataNumeroContato(self, pg):
        global numero

        if pg == 'CadastrarFornecedores':
            numero = self.ui.line_cadastrar_contato_fornecedores
        if pg == 'AlterarFornecedores':
            numero = self.ui.line_alterar_contato_fornecedor
        if pg == 'CadastrarClientes':
            numero = self.ui.line_contato_cadastrar_clientes
        if pg == 'AlterarClientes':
            numero = self.ui.line_alterar_contato_alterar_clientes

        texto = numero.text()
        tamanho = len(numero.text())

        if tamanho == 1 and texto.isnumeric() == True:
            numero.setText(f'({texto}')

        if tamanho == 3 and texto[1:].isnumeric() == True:
            numero.setText(f'{texto}) ')

        if tamanho == 6 and texto[5].isnumeric() == True:
            numero.setText(f'{texto} ')

        if tamanho == 11 and texto[7].isnumeric() == True:
            numero.setText(f'{texto}-')

    # Função para formatar o cpf inserido
    def FormataCPFClientes(self, pg):

        if pg == 'clientes':
            CPF = self.ui.line_xls_cpf_clientes
        if pg == 'Cadastrar':
            CPF = self.ui.line_cpf_cadastrar_clientes
        if pg == 'Alterar':
            CPF = self.ui.line_alterar_cpf_alterar_clientes
        if pg == 'Vendas':
            CPF = self.ui.line_cliente

        TextoInserido = CPF.text()
        TamanhoDoTexto = len(CPF.text())

        if TamanhoDoTexto == 3 and TextoInserido.isnumeric() == True:
            CPF.setText(f'{TextoInserido}.')
        if TamanhoDoTexto == 7 and TextoInserido[4:].isnumeric() == True:
            CPF.setText(f'{TextoInserido}.')
        if TamanhoDoTexto == 11 and TextoInserido[8:].isnumeric() == True:
            CPF.setText(f'{TextoInserido}-')

    # Função para formatar o valor inserido
    def FormataValorProduto(self, pg):

        if pg == 'Cadastrar':
            Valor = self.ui.line_valor_cadastrar_produto_produto.text()
        if pg == 'Alterar':
            Valor = self.ui.line_valor_alterar_produto.text()
        if pg == 'Desconto':
            Valor = self.ui.line_desconto_vendas.text()
        if pg == 'CadastrarValorCusto':
            Valor = self.ui.line_valorcusto_cadastrar_produto.text()
        if pg == 'AlterarValorCusto':
            Valor = self.ui.line_valorcusto_alterar_produto.text()

        # Remove qualquer caractere que não seja numérico
        valorInserido = ''.join(filter(str.isdigit, Valor))

        if valorInserido:

            # Converte para inteiro e depois para float para tratar corretamente
            valor = int(valorInserido)
            valor_float = valor / 100

            # Formata como valor monetário
            numeroFormatado = f"{valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            if pg == 'Cadastrar':
                self.ui.line_valor_cadastrar_produto_produto.setText(numeroFormatado)
                # Move o cursor para o final
                self.ui.line_valor_cadastrar_produto_produto.setCursorPosition(len(numeroFormatado))
            if pg == 'Alterar':
                self.ui.line_valor_alterar_produto.setText(numeroFormatado)
                # Move o cursor para o final
                self.ui.line_valor_alterar_produto.setCursorPosition(len(numeroFormatado))
            if pg == 'Desconto':
                self.ui.line_desconto_vendas.setText(numeroFormatado)
                # Move o cursor para o final
                self.ui.line_desconto_vendas.setCursorPosition(len(numeroFormatado))
            if pg == 'CadastrarValorCusto':
                self.ui.line_valorcusto_cadastrar_produto.setText(numeroFormatado)

                self.ui.line_valorcusto_cadastrar_produto.setCursorPosition(len(numeroFormatado))

            if pg == 'AlterarValorCusto':
                self.ui.line_valorcusto_alterar_produto.setText(numeroFormatado)

                self.ui.line_valorcusto_alterar_produto.setCursorPosition(len(numeroFormatado))

    # Função para Fazer a pesquisa dos item inserido
    def SearchProdutos(self, pg):
        tabela = self
        produto = self

        if pg == 'Produtos':
            tabela = self.ui.tabela_produto
            produto = self.ui.line_search_Bar_produtos

        if pg == 'Alterar':
            tabela = self.ui.tabela_alterar_produto
            produto = self.ui.line_search_Bar_alterar_produto

        if pg == 'Cadastrar':
            tabela = self.ui.tabela_cadastro_produto
            produto = self.ui.line_search_Bar_cadastrar_produto

        items = tabela.findItems(produto.text(), Qt.MatchContains)

        if items:
            item = items[0]
            tabela.setCurrentItem(item)

    def SearchFornecedores(self, pg):
        tabela = self
        produto = self

        if pg == 'Fornecedores':
            tabela = self.ui.tabela_fornecedores
            produto = self.ui.line_search_Bar_fornecedores

        if pg == 'Alterar':
            tabela = self.ui.tabela_alterar_fornecedores
            produto = self.ui.line_search_Bar_altarar_fornecedor

        if pg == 'Cadastrar':
            tabela = self.ui.tabela_cadastrar_fornecedores
            produto = self.ui.line_search_Bar_cadastrar_fornecedores

        items = tabela.findItems(produto.text(), Qt.MatchContains)

        if items:
            item = items[0]
            tabela.setCurrentItem(item)

    def SearchColaboradores(self, pg):
        tabela = self
        produto = self

        if pg == 'Colaboradores':
            tabela = self.ui.tabela_colaboradores
            produto = self.ui.line_search_bar_colaboradores

        if pg == 'Alterar':
            tabela = self.ui.tabela_alterar_colaboradores
            produto = self.ui.line_search_bar_buscar_alterar_colaboradores

        items = tabela.findItems(produto.text(), Qt.MatchContains)

        if items:
            item = items[0]
            tabela.setCurrentItem(item)

    def SearchClientes(self, pg):
        tabela = self
        produto = self

        if pg == 'Clientes':
            tabela = self.ui.tabela_clientes
            produto = self.ui.line_search_Bar_clientes

        if pg == 'Alterar':
            tabela = self.ui.tabela_alterar_clientes
            produto = self.ui.line_search_Bar_alterar_clientes

        if pg == 'Cadastrar':
            tabela = self.ui.tabela_cadastrar_clientes
            produto = self.ui.line_search_Bar_cadastrar_clientes

        items = tabela.findItems(produto.text(), Qt.MatchContains)

        if items:
            item = items[0]
            tabela.setCurrentItem(item)

    def SearchMonitoramentoVendas(self):
        tabela = self.ui.tabela_monitoramento_vendas
        vendas = self.ui.line_search_bar_monitoramentoto_vendas_vendas

        items = tabela.findItems(vendas.text(), Qt.MatchContains)
        if items:
            item = items[0]
            tabela.setCurrentItem(item)

    # Funções para Atualiazar as previções das barras de pesquisa
    def AtualizaCompleterSearchFornecedores(self):
        global search_fornecedores

        cursor.execute('SELECT * FROM fornecedores')
        banco_fornecedores = cursor.fetchall()

        search_fornecedores.clear()
        search_fornecedores = []

        for fornecedor in banco_fornecedores:
            search_fornecedores.append(fornecedor[0])

        self.completer = QCompleter(search_fornecedores)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.ui.line_search_Bar_fornecedores.setCompleter(self.completer)
        self.ui.line_search_Bar_altarar_fornecedor.setCompleter(self.completer)
        self.ui.line_search_Bar_cadastrar_fornecedores.setCompleter(self.completer)
        self.ui.line_fornecedor_cadastrar_produto.setCompleter(self.completer)
        self.ui.line_fornecedor_alterar_produto.setCompleter(self.completer)

    def AtualizaCompleterSearchVendas(self):
        global search_monitoramento

        cursor.execute('SELECT * FROM monitoramento_vendas')
        banco_monitoramento = cursor.fetchall()

        search_monitoramento.clear()

        for venda in banco_monitoramento:
            if venda[0] not in search_monitoramento:
                search_monitoramento.append(venda[0])

            if venda[1] not in search_monitoramento:
                if venda[1] != 'Não Informado':
                    search_monitoramento.append(venda[1])
            search_monitoramento.append(venda[4])

        self.completer = QCompleter(search_monitoramento)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.ui.line_search_bar_monitoramentoto_vendas_vendas_vendas.setCompleter(self.completer)

    def AtualizaCompleterSearchProdutos(self):
        global search_produtos

        cursor.execute('SELECT * FROM produtos')
        banco_produtos = cursor.fetchall()

        search_produtos.clear()

        for produto in banco_produtos:
            search_produtos.append(produto[1])

        self.completer = QCompleter(search_produtos)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)

        self.ui.line_search_Bar_produtos.setCompleter(self.completer)
        self.ui.line_search_Bar_alterar_produto.setCompleter(self.completer)
        self.ui.line_search_Bar_cadastrar_produto.setCompleter(self.completer)
        self.ui.line_search_bar_vendas.setCompleter(self.completer)

    def AtualizaCompleterSearchColaboradores(self):
        global search_colaboradores

        cursor.execute('SELECT * FROM login')
        banco_colaboradores = cursor.fetchall()

        search_colaboradores.clear()

        for colaborador in banco_colaboradores:
            search_colaboradores.append(colaborador[0])

        self.completer = QCompleter(search_colaboradores)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)

        self.ui.line_search_bar_buscar_alterar_colaboradores.setCompleter(self.completer)
        self.ui.line_search_bar_colaboradores.setCompleter(self.completer)

    def AtualizaCompleterSearchClientes(self):
        global search_clientes

        cursor.execute('SELECT * FROM clientes')
        banco_clientes = cursor.fetchall()

        search_clientes.clear()

        for clientes in banco_clientes:
            search_clientes.append(clientes[0])
            search_clientes.append(clientes[1])

        self.completer = QCompleter(search_clientes)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)

        self.ui.line_search_Bar_clientes.setCompleter(self.completer)
        self.ui.line_search_Bar_alterar_clientes.setCompleter(self.completer)
        self.ui.line_search_Bar_cadastrar_clientes.setCompleter(self.completer)
        self.ui.line_cliente.setCompleter(self.completer)

    # Funções de Cadastro
    def CadastroColaboradores(self):

        login = self.ui.line_login_cadastro_colaboradores
        senha = self.ui.line_senha_cadastro_colaboradores
        nome = self.ui.line_nome_cadastro_colaboradores
        nivel = ''


        if login.text() != '' and senha.text() != '' and nome.text() != '':
            nivel = 'admin'

            cursor.execute('SELECT * FROM login')
            banco_login = cursor.fetchall()

            LoginNoBanco = False

            for loginBanco in banco_login:

                if loginBanco[0] == login.text():
                    LoginNoBanco = True
                    break

            if LoginNoBanco == False:
                comando_SQL = 'INSERT INTO login VALUES (?,?,?,?)'
                dados = f'{login.text()}', f'{senha.text()}', f'{nivel}', f'{nome.text()}'
                cursor.execute(comando_SQL, dados)
                banco.commit()

                login.clear()
                senha.clear()
                nome.clear()

                self.ui.line_login_cadastro_colaboradores.setStyleSheet(StyleNormal)

                self.Popup('Cadastro de colaborador', 'Cadastro efetuado com sucesso')

            elif LoginNoBanco == True:
                self.Popup('Cadastro de colaborador', 'Este login ja existe')
                self.ui.line_login_cadastro_colaboradores.setStyleSheet(StyleError)

        self.AtualizaTabelasLogin()
        self.AtualizaCompleterSearchColaboradores()

    def CadastrarClientes(self):

        cpf = self.ui.line_cpf_cadastrar_clientes
        nome = self.ui.line_nome_cadastrar_clientes
        endereco = self.ui.line_endereco_cadastrar_clientes
        contato = self.ui.line_contato_cadastrar_clientes
        saldo_devedor = 'R$ 0,00'

        cursor.execute('SELECT nome FROM clientes WHERE nome = ?',(nome.text(),))
        vefNome = cursor.fetchall()
        cursor.execute('SELECT cpf FROM clientes WHERE cpf = ?', (cpf.text(),))
        vefCPF = cursor.fetchall()

        if cpf.text() != '' and nome.text() != '' and endereco.text() != '' and contato.text() != '':
             if not vefNome and not vefCPF:
                comando_SQL = 'INSERT INTO clientes (CPF, Nome, Endereço, Contato, saldo_devedor) VALUES (?,?,?,?,?)'
                dados = f'{cpf.text()}', f'{nome.text()}', f'{endereco.text()}', f'{contato.text()}', f'{saldo_devedor}'
                cursor.execute(comando_SQL, dados)
                banco.commit()

                self.AtualizaTabelasClientes()
                self.AtualizaCompleterSearchClientes()

                cpf.clear()
                nome.clear()
                endereco.clear()
                contato.clear()
             else:
                 self.Popup("Cadastrar cliente", "Este cliente ja esta cadastrado")
        else:
            self.Popup("Cadastrar cliente", "Insira os dados que faltam")
    def CadastrarFornecedores(self):

        nome = self.ui.line_cadastrar_nome_fornecedores
        endereco = self.ui.line_cadastrar_endereco_fornecedores
        contato = self.ui.line_cadastrar_contato_fornecedores

        cursor.execute('SELECT * FROM fornecedores')
        banco_fornecedores = cursor.fetchall()

        FornecedorNoBanco = False

        for fornecedor in banco_fornecedores:
            if fornecedor[0] == nome.text():
                FornecedorNoBanco = True

        if nome.text() != '' and endereco != '' and contato != '':
            if FornecedorNoBanco == False:
                comando_SQL = 'INSERT INTO fornecedores VALUES (?,?,?)'
                dados = f'{nome.text()}', f'{endereco.text()}', f'{contato.text()}'
                cursor.execute(comando_SQL, dados)
                banco.commit()

                nome.clear()
                endereco.clear()
                contato.clear()

                nome.setStyleSheet(StyleNormal)

                self.AtualizaTabelasFornecedores()
                self.AtualizaCompleterSearchFornecedores()
            else:
                nome.setStyleSheet(StyleError)

    def CadastrarProdutos(self):

        global search_fornecedores

        cod_produto = self.ui.line_codigo_produto_cadastrar
        descricao = self.ui.line_descricao_cadastrar_produto
        valor_unitario = self.ui.line_valor_cadastrar_produto_produto
        qtde_estoque = self.ui.line_qtde_cadastrar_produto
        fornecedor = self.ui.line_fornecedor_cadastrar_produto
        valor_de_custo = self.ui.line_valorcusto_cadastrar_produto
        valor_de_custo_verf = ''.join(re.findall(r'\d', valor_de_custo.text()))

        # valor_unitario = ''.join(re.findall(r'\d', valor_unitario.text()))

        cursor.execute("SELECT * FROM produtos")
        banco_produtos = cursor.fetchall()

        ProdutoJaCadastrado = False
        FornecedorNoSearch = False

        if cod_produto.text() != '' and descricao.text() != '' and valor_unitario.text() != '' and qtde_estoque.text() != '' and fornecedor.text() != '' and valor_de_custo != '' and valor_de_custo_verf.isnumeric():
            for produto in banco_produtos:
                if produto[0] == cod_produto.text():
                    self.Popup("Erro - Cadastrar produto", "Código do produto ja cadastrado")

                    ProdutoJaCadastrado = True

                else:
                    cod_produto.setStyleSheet(StyleNormal)

            if fornecedor.text() in search_fornecedores:
                FornecedorNoSearch = True

                fornecedor.setStyleSheet(StyleNormal)

            else:
                self.Popup("Erro - Cadastrar produto", "Fornecedor não encontrado")

            if ProdutoJaCadastrado == False and FornecedorNoSearch == True:
                comando_SQL = 'INSERT INTO produtos VALUES (?,?,?,?,?,?)'
                dados = f'{cod_produto.text()}', f'{descricao.text()}', f'{valor_unitario.text()}', f'{qtde_estoque.text()}', f'{fornecedor.text()}', f'{valor_de_custo.text()}'
                cursor.execute(comando_SQL, dados)
                banco.commit()

                cod_produto.clear()
                descricao.clear()
                valor_unitario.clear()
                qtde_estoque.clear()
                fornecedor.clear()

                self.AtualizaTabelasProdutos()
                self.AtualizaCompleterSearchProdutos()
        else:
            self.Popup('Erro - Cadastrar produto', 'Preencha todos os campos corretamente')

    def CadastrandoVendas(self):

        global search_produtos, StyleError, StyleNormal

        produtoInserido = self.ui.line_codigo_vendas
        qtde = self.ui.line_quantidade_vendas
        desconto = self.ui.line_desconto_vendas
        NomeProduto = ''
        cliente = self.ui.line_cliente.text()

        ProdutoNoBanco = False
        QuantidadeMenorQueEstoque = False
        DescontoOk = False
        ValorUnitario = 0

        cursor.execute("SELECT * FROM produtos WHERE cód_produto = ?", (produtoInserido.text(),))
        banco_produtos = cursor.fetchall()

        desconto = ''.join(re.findall(r'\d', desconto.text()))

        if produtoInserido.text() != '' and produtoInserido.text().isnumeric() and cliente != '' and desconto.isnumeric():
            if banco_produtos:
                produto = banco_produtos[0]
                ProdutoNoBanco = True
                produtoInserido.setStyleSheet(StyleNormal)
                if qtde.text().isnumeric() == True:
                    if int(produto[3]) >= int(qtde.text()) and int(qtde.text()) > 0:
                        QuantidadeMenorQueEstoque = True
                        qtde.setStyleSheet(StyleNormal)

                        ValorUnitario = produto[2]
                        NomeProduto = produto[1]
                        TotalQtde = int(produto[3]) - int(qtde.text())

                        if desconto == '100':
                            valor = 1.0
                        else:
                            valor = float(desconto) / 100

                        ValorUnitario = ''.join(re.findall(r'\d', ValorUnitario))

                        valorTotal = int(ValorUnitario) * int(qtde.text())
                        descontoTotal = int(valorTotal) * valor

                        if desconto.isnumeric() and 100 >= descontoTotal:
                            DescontoOk = True
                        else:
                            self.Popup("Erro - Adicionar compra", "Desconto não pode ser maior que o valor da compra")
                    else:
                        self.Popup("Erro - Adicionar compra", "Quantidade insuficiente no estoque")
                else:
                    self.Popup("Erro - Adicionar compra", "Informe um numero para quantidade")
            else:
                self.Popup("Erro - Adicionar compra", "Código do produto é inexistente")

            if ProdutoNoBanco == True and QuantidadeMenorQueEstoque == True and DescontoOk == True:
                cursor.execute('SELECT MAX(id) FROM vendas')
                ultimo_id = cursor.fetchone()

                for id_antigo in ultimo_id:
                    if id_antigo == None:
                        id = 0
                    else:
                        id = int(id_antigo) + 1

                valorFinal = (float(valorTotal) - float(descontoTotal))/100
                valorFinal = format(valorFinal, '.2f')
                comando_SQL = 'INSERT INTO vendas VALUES (?,?,?,?,?,?)'
                dados = f'{produtoInserido.text()}', f'{NomeProduto}', f'{ValorUnitario}', f'{qtde.text()}', f'{valorFinal}', f'{id}'
                cursor.execute(comando_SQL, dados)
                banco.commit()

            self.AtualizaTotal()
            self.AtualizaTabelasProdutos()
            self.AtualizaTabelaVendas()
        else:
            self.Popup("Erro - Adicionar compra", "Preencha os campos solicitados corretamente")

    def CadastrandoCompras(self):

        global search_produtos, StyleError, StyleNormal

        produtoInserido = self.ui.line_codigo_vendas_carregamento
        qtde = self.ui.line_quantidade_vendas_carregamento
        NomeProduto = ''

        QuantidadeMaiorQueZero = False
        ProdutoNoBanco = False
        ValorDeCusto = 0

        cursor.execute("SELECT * FROM produtos WHERE cód_produto = ?", (produtoInserido.text(),))
        banco_produtos = cursor.fetchall()

        if produtoInserido.text() != '' and produtoInserido.text().isnumeric():
            if banco_produtos:
                produto = banco_produtos[0]
                ProdutoNoBanco = True
                produtoInserido.setStyleSheet(StyleNormal)
                if qtde.text().isnumeric() == True:
                    if int(qtde.text()) >= 0:
                        QuantidadeMaiorQueZero = True
                        qtde.setStyleSheet(StyleNormal)

                        ValorDeCusto = produto[2]
                        NomeProduto = produto[1]
                        TotalQtde = int(produto[3]) + int(qtde.text())

                        ValorDeCusto = ''.join(re.findall(r'\d', ValorDeCusto))

                        valorTotal = int(ValorDeCusto) * int(qtde.text())
                        descontoTotal = int(valorTotal)

                    else:
                        self.Popup("Erro - Adicionar compra", "Preencha o campo de quantidade corretamente")
                else:
                    self.Popup("Erro - Adicionar compra", "Informe um numero para quantidade")
            else:
                self.Popup("Erro - Adicionar compra", "Código do produto é inexistente")

            if ProdutoNoBanco == True and QuantidadeMaiorQueZero == True:
                cursor.execute('SELECT MAX(id) FROM vendas')
                ultimo_id = cursor.fetchone()

                for id_antigo in ultimo_id:
                    if id_antigo == None:
                        id = 0
                    else:
                        id = int(id_antigo) + 1

                valorFinal = (float(valorTotal))
                valorFinal = format(valorFinal, '.2f')
                comando_SQL = 'INSERT INTO compras VALUES (?,?,?,?,?,?)'
                dados = f'{produtoInserido.text()}', f'{NomeProduto}', f'{ValorDeCusto}', f'{qtde.text()}', f'{valorFinal}', f'{id}'
                cursor.execute(comando_SQL, dados)
                banco.commit()

            self.AtualizaTotalCompras()
            self.AtualizaTabelasProdutos()
            self.AtualizaTabelaCompras()
        else:
            self.Popup("Erro - Adicionar compra", "Preencha os campos solicitados corretamente")


    # Funções de Alterar
    def AlterarColaboradores(self):
        global id_tabela_alterar

        login = self.ui.line_login_alterar_colaboradores
        senha = self.ui.line_senha_alterar_colaboradores
        nome = self.ui.line_nome_alterar_colaboradores

        cursor.execute('SELECT * FROM login')
        banco_login = cursor.fetchall()

        if login.text() != '' and senha.text() != '' and nome.text() != '':

            LoginNoBanco = False

            for pos, user in enumerate(banco_login):
                if login.text() == user[0] and pos != id_tabela_alterar:
                    LoginNoBanco = True

            for pos, user in enumerate(banco_login):
                if pos == id_tabela_alterar:
                    if LoginNoBanco == False:
                        cursor.execute(
                            f'UPDATE login set usuario = "{login.text()}", senha = "{senha.text()}", nivel = "{user[2]}", nome = "{nome.text()}"'
                            f'WHERE usuario = "{user[0]}"')
                        banco.commit()

                        login.clear()
                        senha.clear()
                        nome.clear()

                        self.AtualizaTabelasLogin()
                        self.AtualizaCompleterSearchColaboradores()

                        self.ui.line_login_alterar_colaboradores.setStyleSheet(StyleNormal)
                        break
                    else:
                        self.ui.line_login_alterar_colaboradores.setStyleSheet(StyleError)

    def AlterarClientes(self):
        global id_alterar_Clientes

        cpf = self.ui.line_alterar_cpf_alterar_clientes
        nome = self.ui.line_alterar_nome_alterar_clientes
        endereco = self.ui.line_alterar_endereco_alterar_clientes
        contato = self.ui.line_alterar_contato_alterar_clientes

        cursor.execute('SELECT * FROM clientes')
        banco_clientes = cursor.fetchall()
        if cpf.text() != '' and nome.text() != '' and endereco.text() != '' and contato.text() != '':
            for pos, cliente in enumerate(banco_clientes):
                if pos == id_alterar_Clientes:
                    cursor.execute(
                        f'UPDATE clientes set CPF = "{cpf.text()}", nome = "{nome.text()}", endereço = "{endereco.text()}", contato = "{contato.text()}"'
                        f'WHERE CPF = "{cliente[0]}"')

                    cpf.clear()
                    nome.clear()
                    endereco.clear()
                    contato.clear()

                    self.AtualizaTabelasClientes()
                    self.AtualizaCompleterSearchClientes()

                    break

    def AlterarFornecedores(self):
        global id_alterar_fornecedores

        cursor.execute('SELECT * FROM fornecedores')
        banco_fornecedores = cursor.fetchall()

        nome = self.ui.line_alterar_nome_fornecedor
        endereco = self.ui.line_alterar_endereco_fornecedor
        contato = self.ui.line_alterar_contato_fornecedor

        if nome.text() != '' and endereco.text() != '' and contato.text() != '':
            for pos, fornecedores in enumerate(banco_fornecedores):
                if pos == id_alterar_fornecedores:
                    cursor.execute(
                        f'UPDATE fornecedores set nome = "{nome.text()}", endereço = "{endereco.text()}", contato = "{contato.text()}"'
                        f'WHERE nome = "{fornecedores[0]}"')

                    nome.clear()
                    endereco.clear()
                    contato.clear()

                    self.AtualizaTabelasFornecedores()
                    self.AtualizaCompleterSearchFornecedores()
                    break

    def AlterarProdutos(self):
        global id_alterar_produtos
        global search_fornecedores

        cursor.execute('SELECT * FROM produtos')
        banco_produtos = cursor.fetchall()

        cod_produto = self.ui.line_codigo_alterar_produto
        descricao = self.ui.line_decricao_alterar_produto
        valor_unitario = self.ui.line_valor_alterar_produto
        valor_de_custo = self.ui.line_valorcusto_alterar_produto
        qtde_estoque = self.ui.line_qtde_alterar_produto
        fornecedor = self.ui.line_fornecedor_alterar_produto

        FornecedorNoSearch = False
        ProdutoJaCadastrado = False
        AlterarProduto = ''

        if cod_produto.text() != '' and descricao.text() != '' and valor_unitario.text() != '' and qtde_estoque.text() != '' and fornecedor.text() != '' and valor_de_custo.text() != '':
            if fornecedor.text() in search_fornecedores:
                FornecedorNoSearch = True

                fornecedor.setStyleSheet(StyleNormal)

            else:
                fornecedor.setStyleSheet(StyleError)
                self.Popup("Erro - Alterar produtos", "Fornecedor não encontrado")

            for pos, produto in enumerate(banco_produtos):
                if cod_produto.text() == produto[0] and pos != id_alterar_produtos:
                    ProdutoJaCadastrado = True

                    self.Popup("Erro - Alterar produtos", "Produto ja cadastrado")

                else:
                    cod_produto.setStyleSheet(StyleNormal)

                if pos == id_alterar_produtos:
                    AlterarProduto = produto[0]
        else:
            self.Popup("Erro - Alterar produtos", "Preencha todos campos corretamente")

        if FornecedorNoSearch == True and ProdutoJaCadastrado == False:
            cursor.execute(
                f'UPDATE produtos set cód_produto = ?, descrição = ?, valor_unitário = ?, qtde_estoque = ?, fornecedor = ?, valor_de_custo = ? WHERE cód_produto = ?',
                (
                    cod_produto.text(),
                    descricao.text(),
                    valor_unitario.text(),
                    qtde_estoque.text(),
                    fornecedor.text(),
                    valor_de_custo.text(),
                    AlterarProduto,)
            )

            cod_produto.clear()
            descricao.clear()
            valor_unitario.clear()
            qtde_estoque.clear()
            fornecedor.clear()
            valor_de_custo.clear()

            self.AtualizaTabelasProdutos()
            self.AtualizaCompleterSearchProdutos()
            self.AtualizaTabelaVendas()

    # Funções de Excluir
    def ExcluirColaboradores(self):

        id = self.ui.tabela_colaboradores.currentRow()

        cursor.execute('SELECT * FROM login')
        banco_login = cursor.fetchall()

        deletar_user = ''

        for pos, user in enumerate(banco_login):
            if id == pos:
                deletar_user = user[0]

        cursor.execute(f'DELETE FROM login WHERE usuario = "{deletar_user}"')
        banco.commit()

        self.AtualizaTabelasLogin()
        self.AtualizaCompleterSearchColaboradores()

    def ExcluirClientes(self):
        id = self.ui.tabela_clientes.currentRow()

        cursor.execute('SELECT * FROM clientes')
        banco_clientes = cursor.fetchall()

        deletar_cliente = ''

        for pos, cliente in enumerate(banco_clientes):
            if id == pos:
                deletar_cliente = cliente[0]

        cursor.execute(f'DELETE FROM clientes WHERE CPF = "{deletar_cliente}"')
        banco.commit()

        self.AtualizaTabelasClientes()
        self.AtualizaCompleterSearchClientes()

    def ExluirFornecedores(self):
        id = self.ui.tabela_fornecedores.currentRow()

        cursor.execute('SELECT * FROM fornecedores')
        banco_fornecedor = cursor.fetchall()

        deletar_fornecedor = ''

        for pos, fornecedor in enumerate(banco_fornecedor):
            if id == pos:
                deletar_fornecedor = fornecedor[0]

        cursor.execute(f'DELETE FROM fornecedores WHERE Nome = "{deletar_fornecedor}"')
        banco.commit()

        self.AtualizaTabelasFornecedores()
        self.AtualizaCompleterSearchFornecedores()

    def ExcluirProdutos(self):

        id = self.ui.tabela_produto.currentRow()

        cursor.execute('SELECT * FROM produtos')
        banco_produtos = cursor.fetchall()

        for pos, produto in enumerate(banco_produtos):
            if pos == id:
                cursor.execute(f'DELETE FROM produtos WHERE cód_produto = "{produto[0]}"')
                banco.commit()

                self.AtualizaTabelasProdutos()
                self.AtualizaCompleterSearchProdutos()

    def ExcluirVenda(self):
        id = self.ui.tabela_vendas.currentRow()

        if id != - 1:
            cursor.execute('SELECT * FROM vendas ORDER BY id ASC')
            banco_vendas = cursor.fetchall()
            cursor.execute('SELECT * FROM produtos')
            banco_produtos = cursor.fetchall()

            id_deletado = 0

            for venda in banco_vendas:
                if venda[5] == id:
                    id_deletado = venda[5]

                    for produto in banco_produtos:
                        if venda[0] == produto[0]:
                            TotalEstoque = int(venda[3]) + int(produto[3])
                            cursor.execute(f'UPDATE produtos SET qtde_estoque = "{TotalEstoque}" WHERE cód_produto = "{produto[0]}"')
                            break

                    cursor.execute(f'DELETE FROM vendas WHERE id = {id}')
                    banco.commit()
                    break

            cursor.execute('SELECT * FROM vendas ORDER BY id ASC')
            banco_vendas = cursor.fetchall()

            for venda in banco_vendas:
                if venda[5] > id_deletado:
                    cursor.execute(f'UPDATE vendas set id = {venda[5] - 1} WHERE id = "{venda[5]}"')
                    banco.commit()

        self.AtualizaTabelaVendas()
        self.AtualizaTotal()
        self.AtualizaTabelasProdutos()

    def ExcluirCompra(self):
        id = self.ui.tabela_vendas_carregamento.currentRow()

        if id != - 1:
            cursor.execute('SELECT * FROM compras ORDER BY id ASC')
            banco_compras = cursor.fetchall()
            cursor.execute('SELECT * FROM produtos')
            banco_produtos = cursor.fetchall()

            id_deletado = 0

            for compra in banco_compras:
                if compra[5] == id:
                    id_deletado = compra[5]

                    for produto in banco_produtos:
                        if venda[0] == produto[0]:
                            TotalEstoque = int(compra[3]) + int(produto[3])
                            cursor.execute(f'UPDATE produtos SET qtde_estoque = "{TotalEstoque}" WHERE cód_produto = "{produto[0]}"')
                            break

                    cursor.execute(f'DELETE FROM vendas WHERE id = {id}')
                    banco.commit()
                    break

            cursor.execute('SELECT * FROM vendas ORDER BY id ASC')
            banco_vendas = cursor.fetchall()

            for venda in banco_vendas:
                if venda[5] > id_deletado:
                    cursor.execute(f'UPDATE vendas set id = {venda[5] - 1} WHERE id = "{venda[5]}"')
                    banco.commit()

        self.AtualizaTabelaVendas()
        self.AtualizaTotal()
        self.AtualizaTabelasProdutos()

    # Funções de setar Texto
    def setTextAlterarColaboradores(self):
        global id_tabela_alterar
        nome = self.ui.line_nome_alterar_colaboradores
        login = self.ui.line_login_alterar_colaboradores
        senha = self.ui.line_senha_alterar_colaboradores

        id_tabela_alterar = self.ui.tabela_alterar_colaboradores.currentRow()

        cursor.execute('SELECT * FROM login')
        banco_login = cursor.fetchall()

        for pos, user in enumerate(banco_login):
            if pos == id_tabela_alterar:
                nome.setText(user[3])
                login.setText(user[0])
                senha.setText(user[1])

    def setTextAlterarClientes(self):
        global id_alterar_Clientes
        cpf = self.ui.line_alterar_cpf_alterar_clientes
        nome = self.ui.line_alterar_nome_alterar_clientes
        endereco = self.ui.line_alterar_endereco_alterar_clientes
        contato = self.ui.line_alterar_contato_alterar_clientes

        id_alterar_Clientes = self.ui.tabela_alterar_clientes.currentRow()

        cursor.execute('SELECT * FROM clientes')
        banco_clientes = cursor.fetchall()

        for pos, cliente in enumerate(banco_clientes):
            if pos == id_alterar_Clientes:
                cpf.setText(cliente[0])
                nome.setText(cliente[1])
                endereco.setText(cliente[2])
                contato.setText(cliente[3])

    def setTextAlterarFornecedores(self):
        global id_alterar_fornecedores

        nome = self.ui.line_alterar_nome_fornecedor
        endereco = self.ui.line_alterar_endereco_fornecedor
        contato = self.ui.line_alterar_contato_fornecedor

        id_alterar_fornecedores = self.ui.tabela_alterar_fornecedores.currentRow()

        cursor.execute('SELECT * FROM fornecedores')
        banco_fornecedores = cursor.fetchall()

        for pos, fornecedor in enumerate(banco_fornecedores):
            if pos == id_alterar_fornecedores:
                nome.setText(fornecedor[0])
                endereco.setText(fornecedor[1])
                contato.setText(fornecedor[2])

    def setTextAlterarProdutos(self):
        global id_alterar_produtos

        cod_produto = self.ui.line_codigo_alterar_produto
        descricao = self.ui.line_decricao_alterar_produto
        valor_unitario = self.ui.line_valor_alterar_produto
        qtde_estoque = self.ui.line_qtde_alterar_produto
        fornecedor = self.ui.line_fornecedor_alterar_produto
        valor_de_custo = self.ui.line_valorcusto_alterar_produto

        id_alterar_produtos = self.ui.tabela_alterar_produto.currentRow()

        cursor.execute('SELECT * FROM produtos')
        banco_produtos = cursor.fetchall()

        for pos, produto in enumerate(banco_produtos):
            if pos == id_alterar_produtos:
                cod_produto.setText(produto[0])
                descricao.setText(produto[1])
                valor_unitario.setText(produto[2])
                qtde_estoque.setText(produto[3])
                fornecedor.setText(produto[4])
                valor_de_custo.setText(produto[5])

    # Funções para ver Senha
    def VerSenhaCadastroColaboradores(self):
        global click_cadastro_colaboradores

        click_cadastro_colaboradores += 1

        if click_cadastro_colaboradores % 2 == 0:
            self.ui.line_senha_cadastro_colaboradores.setEchoMode(QLineEdit.EchoMode.Password)
            self.ui.btn_ver_senha_cadastro_colaboradores.setStyleSheet('QPushButton {'
                                                'background-image: url(:/icones/ver senha.png);'
                                                'border: 0px;'
                                                'outline: 0;'
                                                '}'
                                                ''
                                                'QPushButton:hover {'
                                                'background-image: url(:/icones/ver senha hover.png);'
                                                '}')

        if click_cadastro_colaboradores % 2 == 1:
            self.ui.line_senha_cadastro_colaboradores.setEchoMode(QLineEdit.EchoMode.Normal)
            self.ui.btn_ver_senha_cadastro_colaboradores.setStyleSheet('QPushButton {'
                                                'background-image: url(:/icones/bloquear senha.png);'
                                                'border: 0px;'
                                                'outline: 0;'
                                                '}'
                                                ''
                                                'QPushButton:hover {'
                                                'background-image: url(:/icones/bloquear senha hover.png);''}')

    def VerSenhaAlterarColaboradores(self):
        global click_alterar_colaboradores

        click_alterar_colaboradores += 1

        if click_alterar_colaboradores % 2 == 0:
            self.ui.line_senha_alterar_colaboradores.setEchoMode(QLineEdit.EchoMode.Password)
            self.ui.btn_ver_senha_alterar.setStyleSheet('QPushButton {'
                                                        'background-image: url(:/icones/ver senha.png);'
                                                        'border: 0px;'
                                                        'outline: 0;'
                                                        '}'
                                                        ''
                                                        'QPushButton:hover {'
                                                        'background-image: url(:/icones/ver senha hover.png);'
                                                        '}')
        if click_alterar_colaboradores % 2 == 1:
            self.ui.line_senha_alterar_colaboradores.setEchoMode(QLineEdit.EchoMode.Normal)
            self.ui.btn_ver_senha_alterar.setStyleSheet('QPushButton {'
                                                        'background-image: url(:/icones/bloquear senha.png);'
                                                        'border: 0px;'
                                                        'outline: 0;'
                                                        '}'
                                                        ''
                                                        'QPushButton:hover {'
                                                        'background-image: url(:/icones/bloquear senha hover.png);''}')

    # Funções para Atualizar Tabelas
    def AtualizaTabelasLogin(self):

        cursor.execute('SELECT * FROM login')
        banco_login = cursor.fetchall()

        self.ui.tabela_colaboradores.clear()
        self.ui.tabela_alterar_colaboradores.clear()

        row = 0
        self.ui.tabela_colaboradores.setRowCount(len(banco_login))
        self.ui.tabela_alterar_colaboradores.setRowCount(len(banco_login))

        colunas = ['Nome', 'Login', 'Senha']
        self.ui.tabela_colaboradores.setHorizontalHeaderLabels(colunas)
        self.ui.tabela_alterar_colaboradores.setHorizontalHeaderLabels(colunas)

        for pos, logins in enumerate(banco_login):
            self.ui.tabela_colaboradores.setItem(row, 0, QTableWidgetItem(logins[3]))
            self.ui.tabela_colaboradores.setItem(row, 1, QTableWidgetItem(logins[0]))
            self.ui.tabela_colaboradores.setItem(row, 2, QTableWidgetItem(logins[1]))

            self.ui.tabela_alterar_colaboradores.setItem(row, 0, QTableWidgetItem(logins[3]))
            self.ui.tabela_alterar_colaboradores.setItem(row, 1, QTableWidgetItem(logins[0]))
            self.ui.tabela_alterar_colaboradores.setItem(row, 2, QTableWidgetItem(logins[1]))

            row += 1

    def AtualizaTabelasClientes(self):
        cursor.execute('SELECT * FROM clientes')
        banco_clientes = cursor.fetchall()

        self.ui.tabela_clientes.clear()
        self.ui.tabela_alterar_clientes.clear()
        self.ui.tabela_cadastrar_clientes.clear()

        row = 0

        self.ui.tabela_clientes.setRowCount(len(banco_clientes))
        self.ui.tabela_alterar_clientes.setRowCount(len(banco_clientes))
        self.ui.tabela_cadastrar_clientes.setRowCount(len(banco_clientes))

        colunas = ['CPF', 'Nome', 'Endereço', 'Contato', 'Saldo devedor']
        self.ui.tabela_clientes.setHorizontalHeaderLabels(colunas)
        self.ui.tabela_alterar_clientes.setHorizontalHeaderLabels(colunas)
        self.ui.tabela_cadastrar_clientes.setHorizontalHeaderLabels(colunas)

        for clientes in banco_clientes:
            self.ui.tabela_clientes.setItem(row, 0, QTableWidgetItem(clientes[0]))
            self.ui.tabela_clientes.setItem(row, 1, QTableWidgetItem(clientes[1]))
            self.ui.tabela_clientes.setItem(row, 2, QTableWidgetItem(clientes[2]))
            self.ui.tabela_clientes.setItem(row, 3, QTableWidgetItem(clientes[3]))
            self.ui.tabela_clientes.setItem(row, 4, QTableWidgetItem(clientes[4]))

            self.ui.tabela_alterar_clientes.setItem(row, 0, QTableWidgetItem(clientes[0]))
            self.ui.tabela_alterar_clientes.setItem(row, 1, QTableWidgetItem(clientes[1]))
            self.ui.tabela_alterar_clientes.setItem(row, 2, QTableWidgetItem(clientes[2]))
            self.ui.tabela_alterar_clientes.setItem(row, 3, QTableWidgetItem(clientes[3]))
            self.ui.tabela_alterar_clientes.setItem(row, 4, QTableWidgetItem(clientes[4]))

            self.ui.tabela_cadastrar_clientes.setItem(row, 0, QTableWidgetItem(clientes[0]))
            self.ui.tabela_cadastrar_clientes.setItem(row, 1, QTableWidgetItem(clientes[1]))
            self.ui.tabela_cadastrar_clientes.setItem(row, 2, QTableWidgetItem(clientes[2]))
            self.ui.tabela_cadastrar_clientes.setItem(row, 3, QTableWidgetItem(clientes[3]))
            self.ui.tabela_cadastrar_clientes.setItem(row, 4, QTableWidgetItem(clientes[4]))
            row += 1

    def AtualizaTabelasFornecedores(self):

        cursor.execute('SELECT * FROM fornecedores')
        banco_fornecedores = cursor.fetchall()

        self.ui.tabela_fornecedores.clear()
        self.ui.tabela_cadastrar_fornecedores.clear()
        self.ui.tabela_alterar_fornecedores.clear()

        row = 0

        self.ui.tabela_fornecedores.setRowCount(len(banco_fornecedores))
        self.ui.tabela_cadastrar_fornecedores.setRowCount(len(banco_fornecedores))
        self.ui.tabela_alterar_fornecedores.setRowCount(len(banco_fornecedores))

        colunas = ['Nome', 'Endereço', 'Contato']
        self.ui.tabela_fornecedores.setHorizontalHeaderLabels(colunas)
        self.ui.tabela_alterar_fornecedores.setHorizontalHeaderLabels(colunas)
        self.ui.tabela_cadastrar_fornecedores.setHorizontalHeaderLabels(colunas)

        for fornecedores in banco_fornecedores:
            self.ui.tabela_fornecedores.setItem(row, 0, QTableWidgetItem(fornecedores[0]))
            self.ui.tabela_fornecedores.setItem(row, 1, QTableWidgetItem(fornecedores[1]))
            self.ui.tabela_fornecedores.setItem(row, 2, QTableWidgetItem(fornecedores[2]))

            self.ui.tabela_alterar_fornecedores.setItem(row, 0, QTableWidgetItem(fornecedores[0]))
            self.ui.tabela_alterar_fornecedores.setItem(row, 1, QTableWidgetItem(fornecedores[1]))
            self.ui.tabela_alterar_fornecedores.setItem(row, 2, QTableWidgetItem(fornecedores[2]))

            self.ui.tabela_cadastrar_fornecedores.setItem(row, 0, QTableWidgetItem(fornecedores[0]))
            self.ui.tabela_cadastrar_fornecedores.setItem(row, 1, QTableWidgetItem(fornecedores[1]))
            self.ui.tabela_cadastrar_fornecedores.setItem(row, 2, QTableWidgetItem(fornecedores[2]))
            row += 1

    def AtualizaTabelasProdutos(self):
        cursor.execute('SELECT * FROM produtos')
        banco_produtos = cursor.fetchall()

        self.ui.tabela_produto.clear()
        self.ui.tabela_alterar_produto.clear()
        self.ui.tabela_cadastro_produto.clear()

        row = 0

        self.ui.tabela_produto.setRowCount(len(banco_produtos))
        self.ui.tabela_alterar_produto.setRowCount(len(banco_produtos))
        self.ui.tabela_cadastro_produto.setRowCount(len(banco_produtos))

        colunas = ['Item', 'Cód', 'Produto', 'Valor Unitário', 'Qtde', 'Fornecedor', 'Valor de custo']
        self.ui.tabela_produto.setHorizontalHeaderLabels(colunas)
        colunas = ['Item', 'Cód', 'Produto', 'Valor Unitário', 'Qtde', 'Fornecedor', 'Valor de custo']
        self.ui.tabela_alterar_produto.setHorizontalHeaderLabels(colunas)
        colunas = ['Item', 'Cód', 'Produto', 'Valor Unitário', 'Qtde', 'Fornecedor', 'Valor de custo']
        self.ui.tabela_cadastro_produto.setHorizontalHeaderLabels(colunas)



        for pos, produto in enumerate(banco_produtos):

            valor_unitario = produto[2]
            valor_de_custo = produto[5]

            self.ui.tabela_produto.setItem(row, 0, QTableWidgetItem(f'{pos + 1}'))
            self.ui.tabela_produto.setItem(row, 1, QTableWidgetItem(produto[0]))
            self.ui.tabela_produto.setItem(row, 2, QTableWidgetItem(produto[1]))
            self.ui.tabela_produto.setItem(row, 3, QTableWidgetItem('R$ ' + valor_unitario))
            self.ui.tabela_produto.setItem(row, 4, QTableWidgetItem(produto[3]))
            self.ui.tabela_produto.setItem(row, 5, QTableWidgetItem(produto[4]))
            self.ui.tabela_produto.setItem(row, 6, QTableWidgetItem('R$ ' + valor_de_custo))

            self.ui.tabela_alterar_produto.setItem(row, 0, QTableWidgetItem(f'{pos + 1}'))
            self.ui.tabela_alterar_produto.setItem(row, 1, QTableWidgetItem(produto[0]))
            self.ui.tabela_alterar_produto.setItem(row, 2, QTableWidgetItem(produto[1]))
            self.ui.tabela_alterar_produto.setItem(row, 3, QTableWidgetItem('R$ ' + valor_unitario))
            self.ui.tabela_alterar_produto.setItem(row, 4, QTableWidgetItem(produto[3]))
            self.ui.tabela_alterar_produto.setItem(row, 5, QTableWidgetItem(produto[4]))
            self.ui.tabela_alterar_produto.setItem(row, 6, QTableWidgetItem('R$ ' + valor_de_custo))

            self.ui.tabela_cadastro_produto.setItem(row, 0, QTableWidgetItem(f'{pos + 1}'))
            self.ui.tabela_cadastro_produto.setItem(row, 1, QTableWidgetItem(produto[0]))
            self.ui.tabela_cadastro_produto.setItem(row, 2, QTableWidgetItem(produto[1]))
            self.ui.tabela_cadastro_produto.setItem(row, 3, QTableWidgetItem('R$ ' + valor_unitario))
            self.ui.tabela_cadastro_produto.setItem(row, 4, QTableWidgetItem(produto[3]))
            self.ui.tabela_cadastro_produto.setItem(row, 5, QTableWidgetItem(produto[4]))
            self.ui.tabela_cadastro_produto.setItem(row, 6, QTableWidgetItem('R$ ' + valor_de_custo))

            row += 1

    def AtualizaTabelaMonitoramentoVendas(self):
        cursor.execute('SELECT * FROM monitoramento_vendas')
        banco_monitoramento = cursor.fetchall()

        self.ui.tabela_monitoramento_vendas.clear()

        row = 0

        self.ui.tabela_monitoramento_vendas.setRowCount(len(banco_monitoramento))

        colunas = ['Vendedor', 'Cliente', 'Qtde Vendido', 'Total Venda', 'Data/horário']
        self.ui.tabela_monitoramento_vendas.setHorizontalHeaderLabels(colunas)

        for venda in banco_monitoramento:
            # total_venda = lang.toString(int(venda[3]) * 0.01, 'f', 2)

            self.ui.tabela_monitoramento_vendas.setItem(row, 0, QTableWidgetItem(venda[0]))
            self.ui.tabela_monitoramento_vendas.setItem(row, 1, QTableWidgetItem(venda[1]))
            self.ui.tabela_monitoramento_vendas.setItem(row, 2, QTableWidgetItem(venda[2]))
            self.ui.tabela_monitoramento_vendas.setItem(row, 3, QTableWidgetItem('R$ ' + venda[3]))
            self.ui.tabela_monitoramento_vendas.setItem(row, 4, QTableWidgetItem(venda[4]))
            row += 1

    def AtualizaTabelaMonitoramentoCompras(self):
        cursor.execute('SELECT * FROM monitoramento_compras')
        banco_monitoramento = cursor.fetchall()

        self.ui.tabela_monitoramento_compras.clear()

        row = 0

        self.ui.tabela_monitoramento_compras.setRowCount(len(banco_monitoramento))

        colunas = ['Comprador', 'Fornecedor', 'Qtde Comprado', 'Total Compra', 'Data/horário']
        self.ui.tabela_monitoramento_compras.setHorizontalHeaderLabels(colunas)

        for compra in banco_monitoramento:

            self.ui.tabela_monitoramento_compras.setItem(row, 0, QTableWidgetItem(compra[0]))
            self.ui.tabela_monitoramento_compras.setItem(row, 1, QTableWidgetItem(compra[1]))
            self.ui.tabela_monitoramento_compras.setItem(row, 2, QTableWidgetItem(compra[2]))
            self.ui.tabela_monitoramento_compras.setItem(row, 3, QTableWidgetItem('R$ ' + compra[3]))
            self.ui.tabela_monitoramento_compras.setItem(row, 4, QTableWidgetItem(compra[4]))
            row += 1

    def AtualizaTabelaVendas(self):
        cursor.execute('SELECT * FROM vendas ORDER BY id ASC')
        banco_vendas = cursor.fetchall()

        row = 0

        self.ui.tabela_vendas.setRowCount(len(banco_vendas))
        self.ui.tabela_vendas.clear()

        colunas = ['Item', 'Cód', 'Produto', 'Valor Unitário', 'Qtde', 'Total']
        self.ui.tabela_vendas.setHorizontalHeaderLabels(colunas)

        for venda in banco_vendas:

            valor_unitario = ''.join(re.findall(r'\d', venda[2]))

            valor_unitario = lang.toString(int(valor_unitario) * 0.01, 'f', 2)

            # total = lang.toString(int(venda[4]) * 0.01, 'f', 2)
            total = venda[4]

            self.ui.tabela_vendas.setItem(row, 0, QTableWidgetItem(f'{venda[5]}'))
            self.ui.tabela_vendas.setItem(row, 1, QTableWidgetItem(venda[0]))
            self.ui.tabela_vendas.setItem(row, 2, QTableWidgetItem(venda[1]))
            self.ui.tabela_vendas.setItem(row, 3, QTableWidgetItem('R$ ' + valor_unitario))
            self.ui.tabela_vendas.setItem(row, 4, QTableWidgetItem(venda[3]))
            self.ui.tabela_vendas.setItem(row, 5, QTableWidgetItem('R$ ' + total))

            row += 1

    def AtualizaTabelaCompras(self):
        cursor.execute('SELECT * FROM compras ORDER BY id ASC')
        banco_compras = cursor.fetchall()

        row = 0

        self.ui.tabela_vendas_carregamento.setRowCount(len(banco_compras))
        self.ui.tabela_vendas_carregamento.clear()

        colunas = ['Item', 'Cód', 'Produto', 'Valor De Custo', 'Qtde', 'Total']
        self.ui.tabela_vendas_carregamento.setHorizontalHeaderLabels(colunas)

        for compra in banco_compras:

            valor_de_custo = ''.join(re.findall(r'\d', compra[2]))

            valor_de_custo = lang.toString(int(valor_de_custo) * 0.01, 'f', 2)

            total = compra[4]

            self.ui.tabela_vendas_carregamento.setItem(row, 0, QTableWidgetItem(f'{compra[5]}'))
            self.ui.tabela_vendas_carregamento.setItem(row, 1, QTableWidgetItem(compra[0]))
            self.ui.tabela_vendas_carregamento.setItem(row, 2, QTableWidgetItem(compra[1]))
            self.ui.tabela_vendas_carregamento.setItem(row, 3, QTableWidgetItem('R$ ' + valor_de_custo))
            self.ui.tabela_vendas_carregamento.setItem(row, 4, QTableWidgetItem(compra[3]))
            self.ui.tabela_vendas_carregamento.setItem(row, 5, QTableWidgetItem('R$ ' + total))

            row += 1

if __name__ == '__main__':
    # Carregando Planilha
    wb = load_workbook('baseexcel.xlsx')

    # Variáveis Globais
    futuroTexto = ''

    # Clique dos botões de ver senha
    click_cadastro_colaboradores = 0
    click_alterar_colaboradores = 0

    # Id para alterar os itens das tabelas
    id_tabela_alterar = None
    id_alterar_Clientes = None
    id_alterar_fornecedores = None
    id_alterar_produtos = None

    # Lista com os itens para fazer as previções das barras
    search_fornecedores = list()
    search_produtos = list()
    search_colaboradores = list()
    search_clientes = list()
    search_monitoramento = list()
    search_vendas = list()

    # Lista das Vendas
    vendas = list()

    # Conversção para moeda real (R$)
    loc = QLocale.system().name()
    lang = QLocale(loc)

    # Var para pegar o nome de quem logou
    UserLogado = None

    # Estilo de Erro e Padrão
    StyleError = '''
               background-color: rgba(0, 0 , 0, 0);
               border: 2px solid rgba(0,0,0,0);
               border-bottom-color: rgb(255, 17, 49);;
               color: rgb(0,0,0);
               padding-bottom: 8px;
               border-radius: 0px;
               font: 10pt "Montserrat";'''

    StyleNormal = '''
                   background-color: rgba(0, 0 , 0, 0);
                   border: 2px solid rgba(0,0,0,0);
                   border-bottom-color: rgb(159,63,250);;
                   color: rgb(0,0,0);
                   padding-bottom: 8px;
                   border-radius: 0px;
                   font: 10pt "Montserrat";'''

    # Configurando Aplicação
    app = QApplication(sys.argv)
    window = FrmLogin()
    window.show()
    sys.exit(app.exec_())

    # Fechando a conexão com o banco de dados quando o programa terminar
    banco.commit()
    cursor.close()
    banco.close()