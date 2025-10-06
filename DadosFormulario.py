from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QGroupBox, QTableWidget, QHBoxLayout, QPushButton, QLabel, QLineEdit,
                              QDateEdit, QCalendarWidget, QTableWidgetItem,QSplitter, QHeaderView, QFileDialog, QMessageBox, QDialog, QTextEdit, QDialogButtonBox, QGridLayout)

from PyQt5.QtCore import Qt, QDateTime, QDate, QTextCodec, QByteArray
from PyQt5 import QtWebEngineWidgets
import pandas as pd
import numpy as np
import sys
import os
from datetime import datetime
import plotly.express as px
import matplotlib.pyplot as plt
from PyQt5.QtGui import  QTextDocument
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter, QPrintPreviewDialog
from PyQt5.QtWidgets import QComboBox, QSizePolicy
from PyQt5.QtGui import QColor, QBrush, QPixmap
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtWebEngineWidgets import QWebEngineView
import plotly.graph_objects as go
import tempfile
import traceback


### SETUP

col_nome = 0
col_vencimento_doc = 1
col_doc = 2
col_vencimento = 3
tempo_vencimento = 4
col_filial = 5
col_tipo = 6
col_valor = 7
col_observacao = 8
col_soma_acumulada = 9


n_columns = 10

###############################################


class DadosFormulario(QWidget):
    try:
        def __init__(self):
            super().__init__()

            # Grupo de widgets para a tabela de dados
            self.groupBox = QGroupBox("Tabela de Dados")
            self.table = QTableWidget()
            self.table.setColumnCount(n_columns)  # Adicionando uma coluna extra para a soma acumulada
            self.table.setHorizontalHeaderLabels(['NOME','DT DOC', 'DOC','VENCIMENTO','TEMP VENC','FILIAL', 'TIPO', 'VALOR','OBS','SOMA'])
            self.table.horizontalHeader()
            self.table.setColumnWidth(0,400)

            # Inicializar o documento
            self.documento = QTextDocument()

            # Layout para os widgets do lado direito da tabela
            self.layoutRight = QVBoxLayout()

            # Adicionando a tabela e o layout dos widgets do lado direito ao layout principal

           # Inicializar os filtros
            self.combo_filtro = QComboBox(self)
            self.combo_filtro_filial = QComboBox(self)
            self.combo_filtro_tipo = QComboBox(self)
            self.combo_filtro_doc = QComboBox(self)
            self.data_inicial = QDateEdit(self)
            self.data_inicial.setCalendarPopup(True)
            self.data_inicial.setDate(QDate.currentDate())
            self.data_final = QDateEdit(self)
            self.data_final.setCalendarPopup(True)
            self.data_final.setDate(QDate.currentDate().addDays(30))

             
            # Widgets para adicionar saldo inicial
            self.txt_saldo_inicial = QLineEdit()
            self.txt_saldo_inicial.setPlaceholderText('999,99')
            self.botao_saldo_inicial = QPushButton('Inserir')
            saldo_layout = QHBoxLayout()
            saldo_layout.addWidget(QLabel('Saldo Inicial'))
            saldo_layout.addWidget(self.txt_saldo_inicial)
            saldo_layout.addWidget(self.botao_saldo_inicial)
            self.layoutRight.addLayout(saldo_layout)
            self.botao_saldo_inicial.clicked.connect(self.calcular_soma_acumulada)




            # Layout para os campos de entrada de dados
            self.input_layout = QGridLayout()


            # Adicionar filtros ao layout direito
            filtro_layout = QVBoxLayout()

            # Linha para "Filtrar por Nome"
            linha_nome = QHBoxLayout()
            linha_nome.addWidget(QLabel('Filtrar por Nome'), stretch= 1)
            linha_nome.addWidget(self.combo_filtro, stretch= 3)
            filtro_layout.addLayout(linha_nome)

            # Linha para "Filtrar por DOC"
            linha_doc = QHBoxLayout()
            linha_doc.addWidget(QLabel('Filtrar por DOC'), stretch= 1)
            linha_doc.addWidget(self.combo_filtro_doc, stretch= 3)
            filtro_layout.addLayout(linha_doc)

            # Linha para "Filtrar por Filial" e "Filtrar por Tipo"
            linha_filial_tipo = QHBoxLayout()
            linha_filial_tipo.addWidget(QLabel('Filtrar por Filial'))
            linha_filial_tipo.addWidget(self.combo_filtro_filial)
            linha_filial_tipo.addWidget(QLabel('Filtrar por Tipo'))
            linha_filial_tipo.addWidget(self.combo_filtro_tipo)
            filtro_layout.addLayout(linha_filial_tipo)

            # Linha para "De" e "Até" (filtro de data)
            linha_data = QHBoxLayout()
            linha_data.addWidget(QLabel('De:'))
            linha_data.addWidget(self.data_inicial)
            linha_data.addWidget(QLabel('Até:'))
            linha_data.addWidget(self.data_final)
            filtro_layout.addLayout(linha_data)

            # Botão para aplicar o filtro
            self.botao_aplicar_filtro = QPushButton("Aplicar Filtro")
            self.botao_aplicar_filtro.clicked.connect(self.aplicar_filtro)
            filtro_layout.addWidget(self.botao_aplicar_filtro)

            # Adicionar o layout de filtros ao layout direito
            self.layoutRight.addLayout(filtro_layout)


            # Calendário
            self.calendario = QCalendarWidget()
            self.calendario.setGridVisible(True)  # Mostrar as grades do calendário
            self.calendario.selectionChanged.connect(self.filtrar_por_data) 
            self.layoutRight.addWidget(self.calendario)
            
            # Layout para os botões
            button_layout = QHBoxLayout()


            # Botões para ações na tabela
            self.botao_atrasado = QPushButton('Atrasados')
            self.botao_30_atrasado = QPushButton('-30 Dias')
            self.botao_hoje = QPushButton('Hoje')
            self.botao_30 = QPushButton('30 Dias')
            self.botao_15 = QPushButton('15 Dias')
            self.botao_60 = QPushButton('60 Dias')
            self.botao_90 = QPushButton('90 Dias')

            # Adicionando os botões ao layout
            button_layout.addWidget(self.botao_atrasado)
            button_layout.addWidget(self.botao_30_atrasado)
            button_layout.addWidget(self.botao_hoje)
            button_layout.addWidget(self.botao_15)
            button_layout.addWidget(self.botao_30)
            button_layout.addWidget(self.botao_60)
            button_layout.addWidget(self.botao_90)
           
            self.layoutRight.addLayout(button_layout)

            # Adicione o QTextEdit no layout direito
            self.text_observacao = QTextEdit(self)
            self.text_observacao.setReadOnly(False)  # Permitir edição
            self.text_observacao.setPlaceholderText("Observação da linha selecionada aparecerá aqui.")
            self.layoutRight.addWidget(QLabel("Observação:"))
            self.layoutRight.addWidget(self.text_observacao)

            # Botão para salvar a observação editada
            self.botao_salvar_observacao = QPushButton("Salvar Observação")
            self.botao_salvar_observacao.clicked.connect(self.salvar_observacao)
            self.layoutRight.addWidget(self.botao_salvar_observacao)

            # Conecte o evento de clique na tabela
            self.table.cellClicked.connect(self.exibir_observacao)

            # Conecte o evento de mudança de célula para exibir a observação
            self.table.currentCellChanged.connect(self.exibir_observacao)


            # Gráfico
            self.grafico = QtWebEngineWidgets.QWebEngineView(self)

            # Botão para atualizar o gráfico
            self.botao_atualizar_grafico = QPushButton('Atualizar Gráfico')
            self.botao_atualizar_grafico.setShortcut('F5')
            self.botao_atualizar_grafico.clicked.connect(self.atualizar_grafico)
            self.layoutRight.addWidget(self.botao_atualizar_grafico)

            # Conectar botões às funções correspondentes
            self.botao_atrasado.clicked.connect(self.atrasado)
            self.botao_30_atrasado.clicked.connect(self.ddl30A)
            self.botao_hoje.clicked.connect(self.hoje)
            self.botao_15.clicked.connect(self.ddl15)
            self.botao_30.clicked.connect(self.ddl30)
            self.botao_60.clicked.connect(self.ddl60)
            self.botao_90.clicked.connect(self.ddl90)

            self.table.cellActivated.connect(self.atualizar_grafico)

            # Substitua a configuração do layout principal e dos widgets
            self.layout = QVBoxLayout()  # Usar QGridLayout

           # Criar um splitter horizontal para dividir a área entre a tabela e os filtros
            vertical_splitter = QSplitter(Qt.Vertical)

            # Criar um widget para encapsular o layoutRight
            right_widget = QWidget()
            right_widget.setLayout(self.layoutRight)

            # Adicionar a tabela e os filtros ao splitter
            vertical_splitter.addWidget(self.grafico)
            vertical_splitter.addWidget(self.table)

            # Criar um splitter horizontal para dividir a área entre o vertical_splitter e os filtros
            horizontal_splitter = QSplitter(Qt.Horizontal)
            horizontal_splitter.addWidget(vertical_splitter)
            horizontal_splitter.addWidget(right_widget)

            # Definir proporções iniciais para os widgets no splitter
            vertical_splitter.setStretchFactor(0, 4)  # Gráfico ocupa 1 parte
            vertical_splitter.setStretchFactor(1, 5)  # Tabela ocupa 2 partes
            horizontal_splitter.setStretchFactor(0, 5)  # Gráfico e tabela ocupam 3 partes
            horizontal_splitter.setStretchFactor(1, 1)  # Filtros ocupam 1 parte

            # Adicionar o horizontal_splitter ao layout principal
            self.layout.addWidget(horizontal_splitter)

            self.setLayout(self.layout)

            # Gerar o gráfico inicial
            self.Grafico_chart()

            self.items = 0
            
            # Adicione uma lista para armazenar as observações
        
            self.observacoes = []
            self.row = 0
        def salvar_observacao(self):
            """Salva a observação editada no QTextEdit na lista de observações."""
            if 0 <= self.row < self.table.rowCount():
                nova_observacao = self.text_observacao.toPlainText()
                self.table.item(self.row, col_observacao).setText(nova_observacao)
                    
                QMessageBox.information(self, "Observação Salva", "A observação foi salva com sucesso.")
            else:
                QMessageBox.warning(self, "Erro", "Nenhuma linha válida foi selecionada para salvar a observação.")
    
        def exibir_observacao(self):
            """Exibe a observação da linha selecionada no QTextEdit."""

            self.row = self.table.currentIndex().row()  # Obter a linha selecionada
            # Verificar se a linha é válida
            if 0 <= self.row < self.table.rowCount():  # Verificar se a linha está dentro do intervalo
                self.text_observacao.setText(self.table.item(self.row, col_observacao).text())
            else:
                self.text_observacao.clear()


        def gerar_calendario(self):
            dialog = QDialog(self)
            dialog.setWindowTitle("Calendário de Fluxo de Caixa")
            dialog.resize(800, 600)

            layout = QVBoxLayout()

            # Criar o widget do calendário
            calendario = QCalendarWidget()
            calendario.setGridVisible(True)

            # Tabela para exibir a curva ABC
            self.tabela_abc = QTableWidget()
            self.tabela_abc.setColumnCount(4)
            self.tabela_abc.setHorizontalHeaderLabels(['Nome', 'Valor', '% Acumulada', 'Categoria'])
            self.tabela_abc.horizontalHeader().setStretchLastSection(True)

            # Função para atualizar a curva ABC ao selecionar uma data
            def atualizar_curva_abc(date):
                data_str = date.toString("yyyy-MM-dd")
                self.gerar_curva_abc_dia(data_str)

            # Conectar a função ao sinal de seleção de data
            calendario.selectionChanged.connect(lambda: atualizar_curva_abc(calendario.selectedDate()))

            # Adicionar widgets ao layout
            layout.addWidget(calendario)
            layout.addWidget(self.tabela_abc)

            dialog.setLayout(layout)
            dialog.exec_()

        def gerar_curva_abc_dia(self, data_str):
            dados = {'Nome': [], 'Valor': []}

            # Coletar dados da tabela para a data selecionada
            for row in range(self.table.rowCount()):
                item_data = self.table.item(row, col_vencimento)
                item_nome = self.table.item(row, col_nome)
                item_valor = self.table.item(row, col_valor)

                if item_data and item_nome and item_valor and item_data.text() == data_str:
                    try:
                        valor = float(item_valor.text().replace(',', '.'))
                        if valor < 0:  # Considerar apenas saídas
                            dados['Nome'].append(item_nome.text())
                            dados['Valor'].append(valor)
                    except ValueError:
                        continue

            # Criar DataFrame a partir dos dados
            df = pd.DataFrame(dados)

            # Calcular a curva ABC
            if not df.empty:
                df['Valor'] = df['Valor'].abs()
                df_agrupado = df.groupby('Nome')['Valor'].sum().sort_values(ascending=False).reset_index()
                total_geral = df_agrupado['Valor'].sum()
                df_agrupado['%'] = (df_agrupado['Valor'] / total_geral) * 100
                df_agrupado['% Acumulada'] = df_agrupado['%'].cumsum()
                df_agrupado['Categoria'] = df_agrupado['% Acumulada'].apply(
                    lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C')
                )

                # Preencher a tabela com os dados do DataFrame agrupado
                self.tabela_abc.setRowCount(len(df_agrupado))
                for row, row_data in df_agrupado.iterrows():
                    nome_item = QTableWidgetItem(row_data['Nome'])
                    valor_item = QTableWidgetItem(f"R$ {row_data['Valor']:.2f}")
                    acumulada_item = QTableWidgetItem(f"{row_data['% Acumulada']:.2f}%")
                    categoria_item = QTableWidgetItem(row_data['Categoria'])

                    self.tabela_abc.setItem(row, 0, nome_item)
                    self.tabela_abc.setItem(row, 1, valor_item)
                    self.tabela_abc.setItem(row, 2, acumulada_item)
                    self.tabela_abc.setItem(row, 3, categoria_item)
            else:
                self.tabela_abc.setRowCount(0)
            


        def atrasado(self):
            """Filtra a tabela para exibir apenas os itens de hoje."""
            today = QDate.currentDate()
            start_date = today.addDays(-2000)
            self.filtrar_data(start_date, today.addDays(-1))
            self.atualizar_lista_filtro()
        
        def ddl30A(self):
            """Filtra a tabela para exibir os itens dos últimos 30 dias."""
            today = QDate.currentDate()
            start_date = today.addDays(-30)
            self.filtrar_data(start_date, today.addDays(-1))
            self.atualizar_lista_filtro()

        def hoje(self):
            """Filtra a tabela para exibir apenas os itens de hoje."""
            today = QDate.currentDate()
            self.filtrar_data(today, today)
            self.atualizar_lista_filtro()

        def ddl30(self):
            """Filtra a tabela para exibir os itens dos últimos 30 dias."""
            today = QDate.currentDate()
            start_date = today.addDays(30)
            self.filtrar_data(today, start_date)
            self.atualizar_lista_filtro()

        def ddl60(self):
            """Filtra a tabela para exibir os itens dos últimos 60 dias."""
            today = QDate.currentDate()
            start_date = today.addDays(60)
            self.filtrar_data(today, start_date)
            self.atualizar_lista_filtro()
        
        def ddl90(self):
            """Filtra a tabela para exibir os itens dos últimos 90 dias."""
            today = QDate.currentDate()
            start_date = today.addDays(90)
            self.filtrar_data(today, start_date)
            self.atualizar_lista_filtro()


        def filtrar_por_doc(self, doc):
            """Filtra a tabela com base na coluna DOC."""
            col_doc = 1  # Substitua pelo índice correto da coluna DOC
            for row in range(self.table.rowCount()):
                item = self.table.item(row, col_doc)
                if item and doc.lower() in item.text().lower():
                    self.table.setRowHidden(row, False)  # Mostrar a linha
                else:
                    self.table.setRowHidden(row, True)  # Ocultar a linha


        def filtrar_por_data(self):
            data_selecionada = self.calendario.selectedDate()
            self.filtrar_data(data_selecionada, data_selecionada)

        def ddl15(self):
            """Filtra a tabela para exibir os itens dos últimos 30 dias."""
            today = QDate.currentDate()
            start_date = today.addDays(15)
            self.filtrar_data(today, start_date)
            self.atualizar_lista_filtro()


        def filtrar_data(self, start_date, end_date):
            """Filtra a tabela com base nas datas de início e fim."""
            for row in range(self.table.rowCount()):
                item_data = self.table.item(row, col_vencimento)
                if item_data:
                    data = QDate.fromString(item_data.text(), "yyyy-MM-dd")
                    self.table.setRowHidden(row, not (start_date <= data <= end_date))
            
            self.data_inicial.setDate(start_date)
            self.data_final.setDate(end_date)
            self.Grafico_chart()
            self.calcular_soma_acumulada()

        def aplicar_filtro(self):
            nome_filtro = self.combo_filtro.currentText()
            filial_filtro = self.combo_filtro_filial.currentText()
            tipo_filtro = self.combo_filtro_tipo.currentText()
            doc_filtro = self.combo_filtro_doc.currentText()
            data_inicial_val = self.data_inicial.date().toPyDate()
            data_final_val = self.data_final.date().toPyDate()

            for row in range(self.table.rowCount()):
                item_nome = self.table.item(row, col_nome)
                item_filial = self.table.item(row, col_filial)
                item_doc = self.table.item(row, col_doc)
                item_tipo = self.table.item(row, col_tipo)
                item_data = self.table.item(row, col_vencimento)

                mostrar_linha = True

                # Filtro por nome
                if nome_filtro != "Todos" and item_nome and item_nome.text() != nome_filtro:
                    mostrar_linha = False

                # Filtro por filial
                if filial_filtro != "Todos" and item_filial and item_filial.text() != filial_filtro:
                    mostrar_linha = False

                # Filtro por tipo
                if tipo_filtro != "Todos" and item_tipo and item_tipo.text() != tipo_filtro:
                    mostrar_linha = False

                # Filtro por data 
                if item_data:
                    try:
                        data_vencimento = datetime.strptime(item_data.text(), "%Y-%m-%d").date()
                        if not (data_inicial_val <= data_vencimento <= data_final_val):
                            mostrar_linha = False
                    except ValueError:
                        mostrar_linha = False  # Caso a data seja inválida

                # Filtro por DOC (considerar apenas a raiz do documento)
                if doc_filtro != "Todos" and item_doc:
                    doc_raiz = item_doc.text().split('/')[0].strip()  # Extrair a raiz do DOC
                    if doc_filtro.lower() != doc_raiz.lower():
                        mostrar_linha = False

                # Aplica ou remove o filtro
                self.table.setRowHidden(row, not mostrar_linha)
            self.atualizar_lista_filtro()



        def colorir_linhas(self):
            for row in range(self.table.rowCount()):
                tipo_item = self.table.item(row, col_tipo)  # Obtém o item da coluna 'TIPO'
                data_item = self.table.item(row,col_vencimento)
                data_item = datetime.strptime(data_item.text(), "%Y-%m-%d").date()

                # Verifica se o valor da coluna 'TIPO' é "R" - AGENDA DE PAGAMENTO REXROTH
                if tipo_item and tipo_item.text() == "R":
                    # Se for "M", aplica cor de fundo amarela para toda a linha
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        if item is None:
                            item = QTableWidgetItem()  # Cria um item se estiver vazio
                            self.table.setItem(row, col, item)
                        item.setBackground(QBrush(QColor(255, 255, 0)))  # Amarelo

                elif tipo_item and tipo_item.text() == "A": # ANTECIPAÇÃO ARCELOR
            
                    # Se for "A", aplica cor de fundo amarela para toda a linha
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        if item is None:
                            item = QTableWidgetItem()  # Cria um item se estiver vazio
                            self.table.setItem(row, col, item)
                        item.setBackground(QBrush(QColor(64,224,208)))

                elif tipo_item and tipo_item.text() == "E": # PLANEJAMENTO DE PAGAMENTOS EXTRAS
            
                    # Se for "E", aplica cor de fundo amarela para toda a linha
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        if item is None:
                            item = QTableWidgetItem()  # Cria um item se estiver vazio
                            self.table.setItem(row, col, item)
                        item.setBackground(QBrush(QColor(244,164,96))) 
                else:
                    # Caso contrário, mantém a cor original ou outras cores que você tenha configurado
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        if item is None:
                            item = QTableWidgetItem()  # Cria um item se estiver vazio
                            self.table.setItem(row, col, item)
                        if data_item and data_item < datetime.today().date(): # Itens atrasados
                            item.setBackground(QBrush(QColor(253, 253, 150))) # Cor Amarela

                        else:
                            item.setBackground(QBrush(QColor(255, 255, 255)))  # Branco (cor padrão)



        def gerar_relatorio_fluxo_diario(self):
            """Gera um relatório de fluxo diário baseado no período filtrado."""

            datas = []
            entradas = []
            saidas = []

            dialog = QDialog(self)
            dialog.setWindowTitle("Relatório de Fluxo Diário")
            dialog.resize(600, 500)  # Tamanho maior para o gráfico

            layout = QVBoxLayout()

            # Iterar pelas linhas visíveis da tabela
            for row in range(self.table.rowCount()):
                if self.table.isRowHidden(row):  # Ignora linhas ocultas
                    continue

                item_data = self.table.item(row, col_vencimento)  # Verifique se a coluna de data está correta
                item_valor = self.table.item(row, col_valor)  # A mesma coisa para a coluna de valor

                # Verificar se os dados existem
                if item_data and item_valor and item_valor.text():
                    try:
                        valor = float(item_valor.text())
                        datas.append(datetime.strptime(item_data.text().strip(), "%Y-%m-%d").strftime("%Y-%m-%d"))


                        if valor >= 0:
                            entradas.append(valor)
                            saidas.append(0.0)
                        else:
                            entradas.append(0.0)
                            saidas.append(valor)
                    except ValueError:
                        continue  # Ignora erros de conversão

            if not datas:
                QMessageBox.information(self, "Sem Dados", "Não há dados para gerar o relatório.", QMessageBox.Ok)
                return

            df = pd.DataFrame({
                'Data': datas,
                'Entradas': entradas,
                'Saídas': saidas
            })

            # Converter a coluna "Vencimento" para datetime
            df['Data'] = pd.to_datetime(df['Data'], errors='coerce')

            # Agrupar por data e calcular o total de entradas e saídas
            df = df.groupby('Data').sum().reset_index()


            # Criar o gráfico de barras usando Matplotlib
            fig, ax = plt.subplots(figsize=(10, 6))

            # Criando barras para as entradas e saídas
            width = 0.35  # Largura das barras
            indices = range(len(df['Data']))  # Índices das datas

            ax.bar(indices, df['Entradas'], width, label='Entradas', color='green')
            ax.bar([i + width for i in indices], df['Saídas'], width, label='Saídas', color='red')

            # Adicionar linha de base em zero
            ax.axhline(0, color='black', linewidth=1)

            # Ajustar a posição das labels no eixo X
            ax.set_xticks([i + width / 2 for i in indices])  # Ajustando para que as barras de entradas e saídas fiquem ao lado
            ax.set_xticklabels(df['Data'], rotation=45, ha="right")

            # Estilizar o gráfico
            ax.set_title('Fluxo Diário de Entradas e Saídas', fontsize=14)
            ax.set_xlabel('Data', fontsize=12)
            ax.set_ylabel('Valor (R$)', fontsize=12)
            ax.legend()

            # Adicionar grade e melhorar a visualização
            ax.grid(True, linestyle='--', alpha=0.6)

            # Ajustar o formato das datas para melhor visualização
            fig.autofmt_xdate()

            # Salvar a imagem do gráfico em um arquivo temporário
            temp_dir = tempfile.gettempdir()
            temp_file = os.path.join(temp_dir, "fluxo_diario.png")
            fig.savefig(temp_file, dpi=150)

            # Fechar o gráfico para liberar recursos
            plt.close(fig)

            # Exibindo a imagem do gráfico no PyQt
            label_img = QLabel()
            pixmap = QPixmap(temp_file)
            label_img.setPixmap(pixmap)
            label_img.setScaledContents(True)  # Faz a imagem se expandir conforme a tela
            label_img.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

            # Adicionar a imagem ao layout do dialog
            layout.addWidget(label_img)
            dialog.setLayout(layout)

            # Ajustar o tamanho da janela
            dialog.adjustSize()
            dialog.show()  # Exibe o QDialog com o gráfico

        def gerar_relatorio_fluxo_mensal(self):
             """Gera um relatório de fluxo mensal baseado no período filtrado."""

             datas = []
             entradas = []
             saidas = []

             dialog = QDialog(self)
             dialog.setWindowTitle("Relatório de Fluxo Mensal")
             dialog.resize(600, 500)  # Tamanho maior para o gráfico

             layout = QVBoxLayout()

             # Iterar pelas linhas visíveis da tabela
             for row in range(self.table.rowCount()):
                 if self.table.isRowHidden(row):  # Ignora linhas ocultas
                     continue
                 
                 item_data = self.table.item(row, col_vencimento)  # Verifique se a coluna de data está correta
                 item_valor = self.table.item(row, col_valor)  # A mesma coisa para a coluna de valor

                 # Verificar se os dados existem
                 if item_data and item_valor and item_valor.text():
                     try:
                         valor = float(item_valor.text())
                         datas.append(datetime.strptime(item_data.text().strip(), "%Y-%m-%d").strftime("%Y-%m"))  # Agrupar por mês

                         if valor >= 0:
                             entradas.append(valor)
                             saidas.append(0.0)
                         else:
                             entradas.append(0.0)
                             saidas.append(valor)
                     except ValueError:
                         continue  # Ignora erros de conversão
                     
             if not datas:
                 QMessageBox.information(self, "Sem Dados", "Não há dados para gerar o relatório.", QMessageBox.Ok)
                 return

             df = pd.DataFrame({
                 'Data': datas,
                 'Entradas': entradas,
                 'Saídas': saidas
             })

             # Converter a coluna "Data" para datetime
             df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m', errors='coerce')

             # Agrupar por mês e calcular o total de entradas e saídas
             df = df.groupby('Data').sum().reset_index()

             # Criar o gráfico de barras usando Matplotlib
             fig, ax = plt.subplots(figsize=(10, 6))

             # Criando barras para as entradas e saídas
             width = 0.35  # Largura das barras
             indices = range(len(df['Data']))  # Índices das datas

             ax.bar(indices, df['Entradas'], width, label='Entradas', color='green')
             ax.bar([i + width for i in indices], df['Saídas'], width, label='Saídas', color='red')

             # Adicionar linha de base em zero
             ax.axhline(0, color='black', linewidth=1)

             # Ajustar a posição das labels no eixo X
             ax.set_xticks([i + width / 2 for i in indices])  # Ajustando para que as barras de entradas e saídas fiquem ao lado
             ax.set_xticklabels(df['Data'].dt.strftime('%b %Y'), rotation=45, ha="right")  # Exibir mês e ano

             # Estilizar o gráfico
             ax.set_title('Fluxo Mensal de Entradas e Saídas', fontsize=14)
             ax.set_xlabel('Mês', fontsize=12)
             ax.set_ylabel('Valor (R$)', fontsize=12)
             ax.legend()

             # Adicionar grade e melhorar a visualização
             ax.grid(True, linestyle='--', alpha=0.6)

             # Ajustar o formato das datas para melhor visualização
             fig.autofmt_xdate()

             # Salvar a imagem do gráfico em um arquivo temporário
             temp_dir = tempfile.gettempdir()
             temp_file = os.path.join(temp_dir, "fluxo_mensal.png")
             fig.savefig(temp_file, dpi=150)

             # Fechar o gráfico para liberar recursos
             plt.close(fig)

             # Exibindo a imagem do gráfico no PyQt
             label_img = QLabel()
             pixmap = QPixmap(temp_file)
             label_img.setPixmap(pixmap)
             label_img.setScaledContents(True)  # Faz a imagem se expandir conforme a tela
             label_img.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

             # Adicionar a imagem ao layout do dialog
             layout.addWidget(label_img)
             dialog.setLayout(layout)

             # Ajustar o tamanho da janela
             dialog.adjustSize()
             dialog.show()  # Exibe o QDialog com o gráfico


        def gerar_relatorio_balanco(self):
           
            total_entradas, total_saidas, saldo_final, icp = self.calcular_balanco()
            self.exibir_relatorio_balanco(total_entradas, total_saidas, saldo_final, icp)

        def calcular_balanco(self):
            """Gera um relatório de balanço baseado no período filtrado."""
            try:
                total_entradas = 0.0
                total_saidas = 0.0

                # Iterar pelas linhas visíveis na tabela
                for row in range(self.table.rowCount()):
                    if self.table.isRowHidden(row):  # Ignora linhas ocultas
                        continue

                    item_valor = self.table.item(row, col_valor)
                    if item_valor and item_valor.text().strip():
                        try:
                            valor = float(item_valor.text().replace(',', '.'))
                            if valor >= 0:
                                total_entradas += valor
                            else:
                                total_saidas += abs(valor)
                        except ValueError:
                            continue  # Se houver erro, pula essa linha

                # Calcular saldo final
                saldo_final = total_entradas - total_saidas
                # Calcular ICP
                icp = total_entradas / total_saidas if total_saidas != 0 else 0.0

                return total_entradas, total_saidas, saldo_final, icp
            except Exception as e:
                return 0.0, 0.0, 0.0, 0.0

            


        def exibir_relatorio_balanco(self, total_entradas, total_saidas, saldo_final, icp):
            """Exibe o relatório de balanço em um QDialog com tabela e gráfico de barras."""

            dialog = QDialog(self)
            dialog.setWindowTitle("Relatório de Balanço")
            dialog.resize(400, 500)  # Tamanho inicial maior para melhor exibição

            layout = QVBoxLayout()

            # Criar a tabela para exibir os dados do balanço
            tabela_balanco = QTableWidget()
            tabela_balanco.setColumnCount(2)
            tabela_balanco.setRowCount(4)
            tabela_balanco.setHorizontalHeaderLabels(["Descrição", "Valor (R$)"])

            # Dados do relatório
            dados_balanco = [
                ("Total de Entradas", total_entradas),
                ("Total de Saídas", total_saidas),
                ("Saldo Final", saldo_final),
                ("ICP", icp)
            ]

            # Preenchendo a tabela com os valores formatados
            for row, (descricao, valor) in enumerate(dados_balanco):
                if (descricao == 'ICP'):
                    tabela_balanco.setItem(row, 0, QTableWidgetItem(descricao))
                    tabela_balanco.setItem(row, 1, QTableWidgetItem(f"{valor:,.2f}"))
                else:

                    tabela_balanco.setItem(row, 0, QTableWidgetItem(descricao))
                    tabela_balanco.setItem(row, 1, QTableWidgetItem(f"R$ {valor:,.2f}"))

            tabela_balanco.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

            layout.addWidget(tabela_balanco)

            # Criando o gráfico de barras
            tipos = ["Entradas", "Saídas"]
            valores = [total_entradas, total_saidas]
            cores = ["green", "red"]

            fig, ax = plt.subplots(figsize=(5, 4))  # Ajuste inicial maior para melhor qualidade
            bars = ax.bar(tipos, valores, color=cores, width=0.6)

            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width() / 2, height + (height * 0.02), 
                        f"R$ {height:,.0f}", ha="center", fontsize=10, fontweight="bold")

            ax.set_title("Balanço Financeiro", fontsize=12, fontweight="bold", pad=5)
            ax.set_ylabel("Valor (R$)", fontsize=10)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.grid(axis="y", linestyle="--", alpha=0.6)

            # Salvando o gráfico temporariamente
            temp_dir = tempfile.gettempdir()
            temp_file = os.path.join(temp_dir, "relatorio_balanco.png")
            fig.savefig(temp_file, dpi=100)
            plt.close(fig)

            # Exibindo a imagem do gráfico
            label_img = QLabel()
            pixmap = QPixmap(temp_file)
            label_img.setPixmap(pixmap)
            label_img.setScaledContents(True)  # Faz a imagem se expandir conforme a tela
            label_img.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

            layout.addWidget(label_img)
            dialog.setLayout(layout)

            # Ajustando tamanho da janela automaticamente
            dialog.adjustSize()
            dialog.exec()

        def gerar_curva_abc(self, tipo):
            """Gera a curva ABC baseada no tipo (entrada ou saída)."""
            # Criar um DataFrame com os dados da tabela
            dados = {'Nome': [], 'Valor': []}

            # Coletar dados da tabela
            for row in range(self.table.rowCount()):
                
                nome_item = self.table.item(row, col_nome)
                valor_item = self.table.item(row, col_valor)


                if nome_item and valor_item and not self.table.isRowHidden(row):  # Considerar apenas linhas visíveis:
                    try:
                        valor = float(valor_item.text())
                        dados['Nome'].append(nome_item.text())
                        dados['Valor'].append(valor)
                    except ValueError:
                        continue

            # Criar DataFrame a partir dos dados
            df = pd.DataFrame(dados)

            # Filtrar dados com base no tipo
            if tipo == 'entrada':
                df = df[df['Valor'] > 0]
            elif tipo == 'saida':
                df = df[df['Valor'] < 0]
            else:
                raise ValueError("Tipo inválido. Use 'entrada' ou 'saida'.")
            
            # Transformar valores negativos em positivos (somente para o relatório)
            df['Valor'] = df['Valor'].abs()

            # Agrupar por nome e somar os valores
            df_agrupado = df.groupby('Nome')['Valor'].sum().sort_values(ascending=False).reset_index()

            # Calcular o total geral e a porcentagem acumulada
            total_geral = df_agrupado['Valor'].sum()
            df_agrupado['%'] = (df_agrupado['Valor'] / total_geral) * 100
            df_agrupado['% Acumulada'] = df_agrupado['%'].cumsum()

            # Classificar em A, B ou C
            df_agrupado['Categoria'] = df_agrupado['% Acumulada'].apply(
                lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C')
            )

            # Formatar os valores monetários
            df_agrupado['Valor'] = df_agrupado['Valor'].map(self.currency)

            # Exibir relatório
            self.exibir_curva_abc(df_agrupado, tipo)

        
        

        def exibir_curva_abc(self, df_agrupado, tipo):
            """Exibe a curva ABC em uma tabela."""
            dialog_c = QDialog(self)
            dialog_c.setWindowTitle(f"Curva ABC - {tipo.capitalize()}")

            # Definir o tamanho da janela
            dialog_c.resize(850, 900)

            layout = QVBoxLayout()

            # Criar a tabela para exibir os dados
            tabela_abc = QTableWidget()
            tabela_abc.setColumnCount(4)
            tabela_abc.setHorizontalHeaderLabels(['Nome', 'Valor', '% Acumulada', 'Categoria'])
            tabela_abc.setColumnWidth(0,350)

            # Definir o número de linhas com base no DataFrame
            tabela_abc.setRowCount(len(df_agrupado))

            # Preencher a tabela com os dados do DataFrame agrupado
            for row, row_data in df_agrupado.iterrows():
                nome_item = QTableWidgetItem(row_data['Nome'])
                valor_item = QTableWidgetItem(row_data['Valor'])
                acumulada_item = QTableWidgetItem(f"{row_data['% Acumulada']:.2f}%")
                categoria_item = QTableWidgetItem(row_data['Categoria'])

                tabela_abc.setItem(row, 0, nome_item)
                tabela_abc.setItem(row, 1, valor_item)
                tabela_abc.setItem(row, 2, acumulada_item)
                tabela_abc.setItem(row, 3, categoria_item)

            layout.addWidget(tabela_abc)
            dialog_c.setLayout(layout)
            dialog_c.exec()

        def salvar_automaticamente_temporario(self):
            """Salva automaticamente os dados da tabela em um arquivo Excel temporário."""
            try:
                # Criar um arquivo temporário
                temp_dir = tempfile.gettempdir()
                data_atual = datetime.now().strftime("%Y-%m-%d")
                nome_arquivo = os.path.join(temp_dir, f"FLUXO_CAIXA/salvamento_temporario_{data_atual}.xlsx")

                # Chamar a função salvar_tabela para salvar os dados no arquivo temporário
                self.salvar_tabela(nome_arquivo)

                print(f"Arquivo temporário salvo em: {nome_arquivo}")
            except Exception as e:
                print(f"Erro ao salvar o arquivo temporário: {e}")
        
                # Método para preencher a tabela com dados dos arquivos Excel
        def preencher_tabela(self, files):
            # Limpar a tabela antes de preencher com novos dados
            self.table.clearContents()
            self.table.setRowCount(0)  # Certificar-se de que não há linhas restantes

            

            # Criar uma lista para armazenar os dados de todos os arquivos
            all_data = []
            self.observacoes = []  # Limpar a lista de observações

            # Iterar sobre cada arquivo selecionado
            for file in files:
                # Verificar se o arquivo existe
                if os.path.exists(file):
                    # Ler os dados do arquivo Excel
                    df = pd.read_excel(file)
                    all_data.append(df)

            # Concatenar todos os DataFrames em um único DataFrame
            all_data_df = pd.concat(all_data, ignore_index=True)

            
            # Transformar valores a pagar em negativo

            all_data_df['VALOR'] = [-row['VALOR'] if row['CARACTERISTICA'] == 'Pagar' else row['VALOR'] for index, row in all_data_df.iterrows()]


            # Transformar COD de Filial em Abreviação

            all_data_df['FILIAL'] = all_data_df['FILIAL'].map({1:'SZM', 2: 'SVA', 3:'SS', 4:'SNF'})

            # Transformar Tipo em Abrevição

            all_data_df['DATA'] = ''

    

            if all_data_df['VENCIMENTO'].isnull().any():

                QMessageBox.warning(self, "Aviso", "Existe Campo de Data Vazio")
                return

            #ordenar dataframe por vencimento e nome
            all_data_df['VENCIMENTO'] = pd.to_datetime(all_data_df['VENCIMENTO'], format='%d/%m/%Y', errors='coerce')
            all_data_df['DOC VENCIMENTO'] = pd.to_datetime(all_data_df['VENCIMENTO'], format='%d/%m/%Y', errors='coerce')
            all_data_df = all_data_df.sort_values(by=['VENCIMENTO', 'NOME'], ascending=[True, True])

            all_data_df['Dias'] =   all_data_df['DOC VENCIMENTO'] - all_data_df['VENCIMENTO']
            all_data_df['Dias'] = all_data_df['Dias'].dt.days

            
            all_data_df = all_data_df[['NOME','DOC VENCIMENTO', 'DOC','VENCIMENTO','Dias', 'FILIAL', 'TIPO', 'VALOR','OBS']]
            # Preencher a tabela com os dados do DataFrame combinado
            num_rows, num_cols = all_data_df.shape
            self.table.setRowCount(num_rows)  # Definir o número de linhas da tabela

            for row_num, row_data in all_data_df.reindex(np.arange(len(all_data_df))).iterrows():
                for col_num, value in enumerate(row_data):
                    item = QTableWidgetItem()  # Criar um novo QTableWidgetItem
                    if col_num == col_vencimento :  # Se for a coluna de vencimento
                        
                        try:
                            value = datetime.strftime(value, '%d/%m/%Y')
                                            # Converter para string no formato 'dd/MM/yyyy'
                        except TypeError:
                            value = value
                        date = QDateTime.fromString(value, 'dd/MM/yyyy')  # Converter para QDateTime
                        item.setData(Qt.EditRole, date.date())  # Definir os dados como data

                    elif col_num == col_vencimento_doc:  # Se for a coluna de vencimento não editavel
                        try:
                            value = datetime.strftime(value, '%d/%m/%Y')
                                            # Converter para string no formato 'dd/MM/yyyy'
                        except TypeError:
                            value = value
                        date = QDateTime.fromString(value, 'dd/MM/yyyy')  # Converter para QDateTime
                        item.setData(Qt.EditRole, date.date())  # Definir os dados como data
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Tornar o item não editável    
                    else:
                        item.setText(str(value))  # Para outras colunas, definir o texto normalmente
                    self.table.setItem(row_num, col_num, item)


            # Permitir redimensionamento interativo das colunas
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)


             # Habilitar o menu de filtro após preencher a tabela
            self.main_window.enable_filter_menu(True)

            ## Ordenar a tabela pela coluna "Vencimento" e  por nome
            self.table.sortItems(col_vencimento)
            
            
            self.calcular_soma_acumulada()
            self.atualizar_lista_filtro() # Atualizar lista de nomes no combo box

            # Atualizar o gráfico
            self.Grafico_chart()

        def atualizar_dias(self):
            """Atualiza a coluna 'Dias' com a diferença entre 'DOC VENCIMENTO' e 'VENCIMENTO'."""
            for row in range(self.table.rowCount()):
                item_doc_vencimento = self.table.item(row, col_vencimento_doc)
                item_vencimento = self.table.item(row, col_vencimento)
                item_dias = self.table.item(row, tempo_vencimento)

                

                if item_doc_vencimento and item_vencimento:
                    try:
                        # Converter as datas para objetos datetime
                        doc_vencimento = datetime.strptime(item_doc_vencimento.text(), "%Y-%m-%d")
                        vencimento = datetime.strptime(item_vencimento.text(), "%Y-%m-%d")

                        # Calcular a diferença em dias
                        diferenca_dias = (doc_vencimento - vencimento).days

                        # Atualizar a célula da coluna 'Dias'
                        if item_dias is None:
                            item_dias = QTableWidgetItem()
                            self.table.setItem(row, tempo_vencimento, item_dias)
                        item_dias.setText(str(diferenca_dias))
                    except ValueError:
                        # Caso as datas sejam inválidas, limpar a célula
                        if item_dias is None:
                            item_dias = QTableWidgetItem()
                            self.table.setItem(row, tempo_vencimento, item_dias)
                        item_dias.setText("")

        def atualizar_grafico(self, row=None, column=None):
            self.table.sortItems(col_vencimento)
            self.colorir_linhas()
            self.atualizar_dias()
            self.calcular_soma_acumulada()  # Recalcula a soma
            self.Grafico_chart()  # Atualiza o gráfico
                            
            self.atualizar_lista_filtro() #atualizar a lista de filtros
            self.salvar_automaticamente_temporario()

        def cadastrar(self):
            dialog = QDialog(self)
            dialog.setWindowTitle("Cadastrar Novo Item")
            dialog.resize(400, 300)  # Tamanho inicial maior para melhor exibição

            # Layout para os campos de entrada de dados
            input_layout = QGridLayout()

            # Campos de entrada para os dados da tabela
            txt_Nome = QLineEdit()
            txt_doc = QLineEdit()
            txt_vencimento = QDateEdit()
            txt_vencimento.setCalendarPopup(True)
            txt_vencimento.setDate(QDate.currentDate())
            txt_filial = QLineEdit()
            txt_tipo = QLineEdit()
            txt_valor = QLineEdit()
            txt_valor.setPlaceholderText('999,99')

            # Adicionando os campos de entrada ao layout
            input_layout.addWidget(QLabel('Nome:'), 0, 0)
            input_layout.addWidget(txt_Nome, 0, 1)
            input_layout.addWidget(QLabel('Doc:'), 1, 0)
            input_layout.addWidget(txt_doc, 1, 1)
            input_layout.addWidget(QLabel('Vencimento:'), 2, 0)
            input_layout.addWidget(txt_vencimento, 2, 1)
            input_layout.addWidget(QLabel('Filial:'), 3, 0)
            input_layout.addWidget(txt_filial, 3, 1)
            input_layout.addWidget(QLabel('Tipo:'), 4, 0)
            input_layout.addWidget(txt_tipo, 4, 1)
            input_layout.addWidget(QLabel('Valor:'), 5, 0)
            input_layout.addWidget(txt_valor, 5, 1)

            # Botão para salvar o cadastro
            salvar_button = QPushButton("Salvar")
            salvar_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
            input_layout.addWidget(salvar_button, 6, 0, 1, 2, alignment=Qt.AlignCenter)

            dialog.setLayout(input_layout)

            def salvar():
                Nome = txt_Nome.text()
                vencimento = txt_vencimento.date().toString("yyyy-MM-dd")
                valor = txt_valor.text()
                filial = txt_filial.text()
                tipo = txt_tipo.text()
                doc = txt_doc.text()

                # Verificar se todos os campos estão preenchidos
                if not Nome or not vencimento or not valor:
                    QMessageBox.warning(dialog, "Aviso", "Todos os campos devem ser preenchidos.")
                    return  # Não fazer nada se algum campo estiver vazio

                valor = valor.replace(',', '.')
                # Verificar se o valor inserido é numérico
                try:
                    valor_float = float(valor)
                except ValueError:
                    QMessageBox.warning(dialog, "Aviso", "O valor deve ser numérico.")
                    return  # Se não for numérico, não faz nada

                NomeItem = QTableWidgetItem(Nome)
                doc = QTableWidgetItem(doc)
                filial = QTableWidgetItem(filial)
                tipo = QTableWidgetItem(tipo)

                vencimentoItem = QTableWidgetItem()
                vencimentoItem.setData(Qt.EditRole, QDate.fromString(vencimento, "yyyy-MM-dd"))

                # Adicionar os itens à tabela
                self.table.insertRow(self.items)
                self.table.setItem(self.items, col_nome, NomeItem)
                self.table.setItem(self.items, col_doc, doc)
                self.table.setItem(self.items, col_vencimento, vencimentoItem)
                self.table.setItem(self.items, col_filial, filial)
                self.table.setItem(self.items, col_tipo, tipo)
                self.table.setItem(self.items, col_valor, QTableWidgetItem(valor))
                self.table.setItem(self.items, col_observacao, QTableWidgetItem(" "))  # Item para a coluna "Observação"
                  # Item para a coluna "Valor"

                # Incrementar o contador de itens
                self.items += 1

                # Limpar os campos de entrada
                txt_Nome.clear()
                txt_vencimento.setDate(QDate.currentDate())
                txt_valor.clear()

                self.table.sortItems(col_vencimento)  # Ordenar a tabela após adicionar a linha


                self.calcular_soma_acumulada()  # Recalcular soma
                self.atualizar_lista_filtro()  # Atualizar lista de nomes no combo box

                dialog.accept()

            salvar_button.clicked.connect(salvar)
            dialog.exec()

        def calcular_drawdown(self):
            """Calcula o drawdown da soma acumulada."""
            soma_acumulada = []

            # Obter os valores da coluna de soma acumulada
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Considerar apenas linhas visíveis
                    item_soma = self.table.item(row, col_soma_acumulada)
                    if item_soma and item_soma.text():
                        try:
                            soma_acumulada.append(float(item_soma.text()))
                        except ValueError:
                            continue  # Ignorar valores inválidos
                        
            if not soma_acumulada:
                QMessageBox.warning(self, "Aviso", "Não há dados suficientes para calcular o drawdown.")
                return None

            # Calcular o drawdown
            max_pico = -float('inf')  # Inicializar o pico máximo como menos infinito
            drawdowns = []

            for valor in soma_acumulada:
                max_pico = max(max_pico, valor)  # Atualizar o pico máximo
                drawdown = (valor - max_pico) 
                drawdowns.append(drawdown)

            # Encontrar o drawdown máximo (mais negativo)
            max_drawdown = min(drawdowns)

            return max_drawdown




        def calcular_soma_acumulada(self):
            try:
                saldo_inicial = float(self.txt_saldo_inicial.text().replace(',', '.'))
            except ValueError:
                saldo_inicial = 0.0  # Se o saldo inicial for inválido, definir como 0
        
            soma_acumulada = saldo_inicial
        
            linha_valida = 0  # Controla a posição real das linhas visíveis para a soma acumulada correta
        
            for row in range(self.table.rowCount()):
                if self.table.isRowHidden(row):  # Se a linha está oculta, pular
                    continue
                
                item_valor = self.table.item(row, col_valor)
        
                if item_valor and item_valor.text().strip():  # Garantir que a célula não é None ou vazia
                    try:
                        valor = float(item_valor.text().replace(',', '.'))
                    except ValueError:
                        continue  # Se o valor não for numérico, pular a linha
                    
                    soma_acumulada += valor
                    self.table.setItem(row, col_soma_acumulada, QTableWidgetItem(str(np.around(soma_acumulada, 2))))
        
                    linha_valida += 1  # Apenas incrementar se a linha não estiver oculta
        
            self.Grafico_chart()  # Atualizar o gráfico após recalcular

        def atualizar_lista_filtro(self):
            # Atualizar combo_filtro
            self.combo_filtro.blockSignals(True)
            self.combo_filtro.clear()
            self.combo_filtro.addItem("Todos")
            nomes = set()
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Considerar apenas linhas visíveis
                    nome_item = self.table.item(row, col_nome)
                    if nome_item and nome_item.text():
                        nomes.add(nome_item.text())
            self.combo_filtro.addItems(sorted(nomes))
            self.combo_filtro.blockSignals(False)

            # Atualizar combo_filtro_filial
            self.combo_filtro_filial.blockSignals(True)
            self.combo_filtro_filial.clear()
            self.combo_filtro_filial.addItem("Todos")
            filiais = set()
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Considerar apenas linhas visíveis
                    filial_item = self.table.item(row, col_filial)
                    if filial_item and filial_item.text():
                        filiais.add(filial_item.text())
            self.combo_filtro_filial.addItems(sorted(filiais))
            self.combo_filtro_filial.blockSignals(False)

            # Atualizar combo_filtro_tipo
            self.combo_filtro_tipo.blockSignals(True)
            self.combo_filtro_tipo.clear()
            self.combo_filtro_tipo.addItem("Todos")
            tipos = set()
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Considerar apenas linhas visíveis
                    tipo_item = self.table.item(row, col_tipo)
                    if tipo_item and tipo_item.text():
                        tipos.add(tipo_item.text())
            self.combo_filtro_tipo.addItems(sorted(tipos))
            self.combo_filtro_tipo.blockSignals(False)

            # Atualizar combo_filtro_doc
            self.combo_filtro_doc.blockSignals(True)
            self.combo_filtro_doc.clear()
            self.combo_filtro_doc.addItem("Todos")
            docs = set()
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Considerar apenas linhas visíveis
                    item = self.table.item(row, col_doc)
                    if item and item.text():
                        # Extrair o documento raiz (antes do '/')
                        doc_raiz = item.text().split('/')[0].strip()
                        docs.add(doc_raiz)  # Adicionar o documento raiz ao conjunto
            self.combo_filtro_doc.addItems(sorted(docs))
            self.combo_filtro_doc.blockSignals(False)


    

        
        def reset_tabela(self):
            self.table.clearContents()
            self.table.setRowCount(0)
            self.items = 0

            self.Grafico_chart()



        def calcular_tempo_caixa_negativo(self):
            """Calcula a quantidade de dias em que o saldo acumulado ficou negativo."""
            soma_acumulada = []
            datas_negativas = set()
            datas_todas = set()

            # Obter os valores da coluna de soma acumulada e as datas correspondentes
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Considerar apenas linhas visíveis
                    item_soma = self.table.item(row, col_soma_acumulada)
                    item_data = self.table.item(row, col_vencimento)
                    if item_soma and item_soma.text() and item_data and item_data.text():
                        try:
                            valor = float(item_soma.text())
                            data = datetime.strptime(item_data.text(), "%Y-%m-%d").date()
                            datas_todas.add(data)  # Adicionar data ao conjunto de todas as datas
                            if valor < 0:  # Verificar se o saldo acumulado é negativo
                                datas_negativas.add(data)
                        except ValueError:
                            continue  # Ignorar valores inválidos
                        
            # Retornar a quantidade de dias únicos em que o saldo ficou negativo
            return len(datas_negativas), len(datas_negativas) / len(datas_todas) * 100 if datas_todas else 0

        def Grafico_chart(self):
            # Criar listas para armazenar vencimentos e valores
            vencimento = []
            valor = []
        
            # Obter o saldo inicial
            try:
                saldo_inicial = float(self.txt_saldo_inicial.text())
            except ValueError:
                saldo_inicial = 0.0
        
            # Iterar pelas linhas visíveis da tabela
            for i in range(self.table.rowCount()):
                if not self.table.isRowHidden(i):  # Considerar apenas linhas visíveis
                    nome_item = self.table.item(i, col_vencimento)
                    valor_item = self.table.item(i, col_valor)
        
                    # Verificar se os itens existem e têm conteúdo
                    if (nome_item and nome_item.text() and valor_item and valor_item.text()):
                        vencimento_text = nome_item.text()
                        valor_text = valor_item.text()
        
                        try:
                            vencimento.append(vencimento_text)
                            valor.append(float(valor_text))
                        except ValueError:
                            continue  # Ignorar valores inválidos sem interromper o processo
                    
            
            # Criar DataFrame com os dados
            if len(vencimento) == len(valor):
                df = pd.DataFrame({'Vencimento': vencimento, 'Valor': valor})
            else:
                QMessageBox.information(self, "Preencher PDF", f"A tabela está vazia. Não há dados para exportar. len vencimento: {len(vencimento)}, len valor: {len(valor)}", QMessageBox.Ok)
                return
        
            # Converter a coluna "Vencimento" para datetime
            df['Vencimento'] = pd.to_datetime(df['Vencimento'], errors='coerce')
        
            # Remover linhas com datas inválidas
            df.dropna(subset=['Vencimento'], inplace=True)
        
            # Adicionar o saldo inicial ao primeiro valor, se aplicável
            if not df.empty:
                df.loc[df.index[0], 'Valor'] += saldo_inicial
        
            # Agrupar por data e calcular a soma acumulada
            df = df.groupby(df['Vencimento'].dt.date)['Valor'].sum().cumsum().reset_index()
            df.columns = ['Vencimento', 'Valor']  # Renomear colunas

                
        
            # Criar o gráfico usando Plotly
            fig = px.line(df, x='Vencimento', y='Valor', title="Fluxo de Caixa")
            fig.add_hline(y=0, line_color='red')
        

            # mostrar banlanço somente se a tabela não estiver vazia
            if not df.empty:
                
                # Adicionar balanço no cabeçalho do gráfico
                
                total_entradas, total_saidas, saldo_final, icp = self.calcular_balanco()

                # Adicionar drawdown
                drawdown = self.calcular_drawdown()

                # Dias de caixa negativo
                dias_caixa_negativo, per_dias_negativo = self.calcular_tempo_caixa_negativo()
                if dias_caixa_negativo > 0:
                    dias_caixa_negativo = f"{dias_caixa_negativo} dias"
                else:
                    dias_caixa_negativo = "Nenhum dia"
                            
                fig.add_annotation(
                     xref="paper", yref="paper",
                     x=1, y=1,
                     text=f"Total Entradas: {self.currency(total_entradas)}<br>Total Saídas: R$ {self.currency(total_saidas)}<br>Saldo Final: R$ {self.currency(saldo_final)}<br>ICP: {icp:.2f}<br>Drawdown: {self.currency(drawdown)}<br>Caixa Negativo: {dias_caixa_negativo}({per_dias_negativo:.2f}%)",
                     showarrow=False,
                     font=dict(size=12, color="black"),
                     align="center",
                     bgcolor="lightgrey",
                     bordercolor="black",
                     borderwidth=1,
                     borderpad=5,
                     opacity=0.8,
                 )
            
            
             
            # Estilizar o gráfico
            fig.update_layout(
                 xaxis_title="Data",
                 yaxis_title="Valor Acumulado (R$)",
                 plot_bgcolor='white',
                 margin=dict(t=50, l=50, r=50, b=50),
             )
    
    
            
            # Exibir o gráfico no componente HTML
            self.grafico.setHtml(fig.to_html(include_plotlyjs='cdn'))


        def currency(self, valor):

            try:
                valor = float(valor)

                valor_moeda = "R$ {:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")
            except ValueError:
                valor_moeda = ' '

            return valor_moeda

        def preencher_pdf(self):

            self.calcular_soma_acumulada()
            
            def categorizar_valor(valores):
                pagar = valores.apply(lambda x: x if x < 0 else ' ')
                receber = valores.apply(lambda x: x if x >= 0 else ' ')
                return pagar, receber

            # Verifica se a tabela está vazia
            if self.table.rowCount() == 0:
                QMessageBox.information(self, "Preencher PDF", "A tabela está vazia. Não há dados para exportar.", QMessageBox.Ok)
                return
            # String para armazenar os dados da tabela formatados em HTML
            datos = ""
            nome = []
            doc = []
            vencimento = []
            #data = []
            filial = []
            valor = []
            soma = []
            # Itera sobre as linhas da tabela e obtém os dados
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Apenas linhas visíveis
                        nome.append(self.table.item(row, col_nome).text())
                        doc.append(self.table.item(row,col_doc).text())
                        vencimento.append(self.table.item(row, col_vencimento).text())
                        filial.append(self.table.item(row, col_filial ).text())
                        valor.append(self.table.item(row, col_valor).text())
                        soma.append(self.table.item(row, col_soma_acumulada).text())

                # Se nenhuma linha visível for encontrada, exibe mensagem
            if not nome:
                QMessageBox.information(self, "Preencher PDF", "Nenhum dado foi encontrado dentro do período selecionado.", QMessageBox.Ok)
                return
                
            df = pd.DataFrame({'nome': nome, 'doc': doc, 'filial': filial, 'vencimento':vencimento,'valor': valor, 'soma':soma })

            try:

                df['valor'] = df['valor'].astype(float)
            except ValueError:
                QMessageBox.warning(self, "Aviso", f'O campo Valor deve ser um número válido.')
                return  # Retorna para interromper a iteração se o campo Valor não puder ser convertido para float
                


            pagar, receber = categorizar_valor(df['valor'])
            df['pagar'] = pagar
            df['receber'] = receber
            


            # Adiciona os dados formatados à string HTML

            datos = ""
            reporteHtml = ""


            ############ TABELA ###############
            
            for data in df['vencimento'].unique():
                dato = ""
                for i in range(len(df[df['vencimento'] == data])):
                    

                    aux = df[df['vencimento'] == data]
                    aux.reset_index(inplace = True)
                    
                    nome = aux['nome'][i]
                    doc = aux['doc'][i]
                    vencimento = aux['vencimento'][i]
                    pagar = self.currency(aux['pagar'][i])
                    receber = self.currency(aux['receber'][i])
                    filial = aux['filial'][i]
                    soma = self.currency(aux['soma'][i])
                    dato += f"<tr><td>{nome[:30]}</td><td>{doc}</td><td>{filial}</td><td>{pagar}</td><td>{receber}</td><td>{soma}</td></tr>"


                try:
                    receber_total = sum(df[(df['vencimento'] == data) & (-df['receber'].isin([' ']))]['receber'])
                    

                except TypeError:
                
                    receber_total = 0.00

                try:
                    pagar_total   = sum(df[(df['vencimento'] == data) & (-df['pagar'].isin([' ']))]['pagar'])
                
                except TypeError:

                    pagar_total = 0.00


                receber_total = self.currency(receber_total)
                pagar_total = self.currency(pagar_total)
            
                

                dato += f"<tr><td>{'_'*31}</td><td>{'_'}</td><td>{'_'}</td><td><strong>{'Total:'}</strong></td><td><strong>{pagar_total}</strong></td><td><strong>{receber_total}</strong></td><td><strong>R$ {soma}</strong></td></tr>"
                
                data_formatada = datetime.strptime(data, '%Y-%m-%d').strftime('%d/%m/%Y')

                datos += f"""
                            <p style="text-align: center;" >{'_'*85}</p>

                            <p><strong>Dt. Vencimento {data_formatada}</strong></p>
                            <table>
                                <thead>
                                    <tr>
                                        <th>Nome</th>
                                        <th>Doc</th>
                                        <th>Filial</th>
                                        <th>Pagar</th>
                                        <th>Receber</th>
                                        <th>Soma</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {dato}
                                </tbody>
                            </table>
                            """



            data_inicio = datetime.strptime(min(df['vencimento']), '%Y-%m-%d').strftime('%d/%m/%Y')
            data_fim = datetime.strptime(max(df['vencimento']), '%Y-%m-%d').strftime('%d/%m/%Y')

            reporteHtml = f"""
                            <!DOCTYPE html>
                            <html>
                            <head>
                                <meta charset="UTF-8">
                                <style>
                                    h1 {{
                                        font-family: Helvetica-Bold;
                                        text-align: center;
                                    }}

                                    table {{
                                        font-family: arial, sans-serif;
                                        border-collapse: collapse;
                                        width: 100%;
                                    }}

                                    td, th {{
                                        text-align: left;
                                        padding: 8px;
                                    }}

                                    /* Definindo o tamanho da coluna "Nome" */
                                    td:nth-child(1) {{
                                        width: 50%; 
                                    }}

                                    th:nth-child(1) {{
                                        width: 50%; 
                                    }}

                                    hr {{
                                        border: none;
                                        border-top: 1px solid black;
                                    }}

                                    tr:nth-child(even) {{
                                        background-color: #dddddd;
                                    }}
                                </style>
                            </head>
                            <body>
                                <h1>Fluxo de Caixa</h1>
                                <h4 style="text-align: left;"  > Saldo Inicial: R$ {self.txt_saldo_inicial.text()} </h4>
                                <h4 style="text-align: left;"  > Saldo Final: R$ {soma} </h4>
                                <p style="text-align: left;"> Periodo: {data_inicio} - {data_fim}</p>
                                <p style="text-align: right;" >Emissão: {datetime.today().strftime("%d-%m-%Y %H:%M")}</p>

                                {datos}
                            """
    



            # HTML do relatório
            


            # Define o documento HTML
            self.documento = QTextDocument()
            datos = QByteArray()
            datos.append(str(reporteHtml))
            codec = QTextCodec.codecForHtml(datos)
            unistr = codec.toUnicode(datos)
            # Verifica se o texto é formato HTML e define o documento
            if Qt.mightBeRichText(unistr):
                self.documento.setHtml(unistr)
            else:
                self.documento.setPlainText(unistr)
            # Exibe mensagem se não houver dados na tabela
            if not datos:
                QMessageBox.information(self, "Preencher PDF", "Não foram encontrados dados na tabela.", QMessageBox.Ok)

        def vistaPrevia(self):
            self.preencher_pdf()
            if not self.documento.isEmpty():
                # Criar uma impressora
                printer = QPrinter(QPrinter.HighResolution)

                # Definir o tamanho da página (por exemplo, A4)
                printer.setPageSize(QPrinter.A4)

                # Definir a orientação da página (retrato ou paisagem)
                printer.setOrientation(QPrinter.Portrait)

                # Definir as margens (em mm)
                printer.setPageMargins(2, 2, 2, 2, QPrinter.Millimeter)

                # Configurar a visualização de impressão
                preview = QPrintPreviewDialog(printer)
                preview.paintRequested.connect(self.documento.print_)
                preview.exec_()
            else:
                QMessageBox.critical(self, "Vista Previa", "Não há dados para visualizar", QMessageBox.Ok)

        def imprimir(self):
            self.preencher_pdf()
            if not self.documento.isEmpty():
                dialogo_imprimir = QPrintDialog()
                if dialogo_imprimir.exec_() == QPrintDialog.Accepted:
                    self.documento.print_(dialogo_imprimir.printer())
            else:
                QMessageBox.critical(self, "Imprimir", "Não há dados para imprimir", QMessageBox.Ok)

        def exportarPDF(self):
            if not self.documento.isEmpty():
                vencimento_archivo, _ = QFileDialog.getSaveFileName(self, "Exportar a PDF", "Lista de Usuarios", "Archivos PDF (*.pdf)")
                if vencimento_archivo:
                    impresion = QPrinter(QPrinter.HighResolution)
                    impresion.setOutputFormat(QPrinter.PdfFormat)
                    impresion.setOutputFileName(vencimento_archivo)
                    self.documento.print_(impresion)
            else:
                QMessageBox.critical(self, "Exportar a PDF", "No hay datos para exportar", QMessageBox.Ok)


        def preencher_pdf_saidas(self):
            self.calcular_soma_acumulada()
        
            def categorizar_valor(valores):
                pagar = valores.apply(lambda x: x if x < 0 else ' ')
                return pagar
        
            # Verifica se a tabela está vazia
            if self.table.rowCount() == 0:
                QMessageBox.information(self, "Preencher PDF", "A tabela está vazia. Não há dados para exportar.", QMessageBox.Ok)
                return
        
            # String para armazenar os dados da tabela formatados em HTML
            datos = ""
            nome = []
            doc = []
            vencimento = []
            filial = []
            valor = []
            soma = []
        
            # Itera sobre as linhas da tabela e obtém os dados
            for row in range(self.table.rowCount()):
                if not self.table.isRowHidden(row):  # Apenas linhas visíveis
                    valor_item = self.table.item(row, col_valor)
                    if valor_item and float(valor_item.text()) < 0:  # Considerar apenas saídas
                        nome.append(self.table.item(row, col_nome).text())
                        doc.append(self.table.item(row, col_doc).text())
                        vencimento.append(self.table.item(row, col_vencimento).text())
                        filial.append(self.table.item(row, col_filial).text())
                        valor.append(self.table.item(row, col_valor).text())
        
            # Se nenhuma linha visível for encontrada, exibe mensagem
            if not nome:
                QMessageBox.information(self, "Preencher PDF", "Nenhum dado foi encontrado dentro do período selecionado.", QMessageBox.Ok)
                return
        
            df = pd.DataFrame({'nome': nome, 'doc': doc, 'filial': filial, 'vencimento': vencimento, 'valor': valor})
        
            try:
                df['valor'] = df['valor'].astype(float)
            except ValueError:
                QMessageBox.warning(self, "Aviso", f'O campo Valor deve ser um número válido.')
                return  # Retorna para interromper a iteração se o campo Valor não puder ser convertido para float
        
            pagar = categorizar_valor(df['valor'])
            df['pagar'] = pagar
        
            # Adiciona os dados formatados à string HTML
            datos = ""
            reporteHtml = ""
        
            ############ TABELA ###############
            for data in df['vencimento'].unique():
                dato = ""
                for i in range(len(df[df['vencimento'] == data])):
                    aux = df[df['vencimento'] == data]
                    aux.reset_index(inplace=True)
        
                    nome = aux['nome'][i]
                    doc = aux['doc'][i]
                    vencimento = aux['vencimento'][i]
                    pagar = self.currency(aux['pagar'][i])
                    filial = aux['filial'][i]
                    dato += f"<tr><td>{nome[:30]}</td><td>{doc}</td><td>{filial}</td><td>{pagar}</td></tr>"
        
                try:
                    pagar_total = sum(df[(df['vencimento'] == data) & (-df['pagar'].isin([' ']))]['pagar'])
                except TypeError:
                    pagar_total = 0.00
                
                    # Calcular o saldo final
                
                # Verificar se o campo de saldo inicial está vazio e definir um valor padrão se necessário
                saldo_inicial_text = self.txt_saldo_inicial.text().replace(',', '.')
                if saldo_inicial_text:
                   saldo_inicial = float(saldo_inicial_text)
                else:
                   saldo_inicial = 0.0

                # Calcular o saldo final
                saldo_final = saldo_inicial + pagar_total
        
                dato += f"<tr><td>{'_'}</td><td>{'_'}</td><td>{'_'}</td><strong>{'Total:'}</strong></td><td><strong>{self.currency(pagar_total)}</strong></td></tr>"
        
                data_formatada = datetime.strptime(data, '%Y-%m-%d').strftime('%d/%m/%Y')
        
                datos += f"""
                            <p style="text-align: center;" >{'_'*85}</p>

                            <p><strong>Dt. Vencimento {data_formatada}</strong></p>
                            <table>
                                <thead>
                                    <tr>
                                        <th style="width: 40%;">Nome</th>
                                        <th style="width: 20%;">Doc</th>
                                        <th style="width: 20%;">Filial</th>
                                        <th style="width: 20%;">Pagar</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {dato}
                                </tbody>
                            </table>
                            """
        
            data_inicio = datetime.strptime(min(df['vencimento']), '%Y-%m-%d').strftime('%d/%m/%Y')
            data_fim = datetime.strptime(max(df['vencimento']), '%Y-%m-%d').strftime('%d/%m/%Y')
        
            reporteHtml = f"""
                            <!DOCTYPE html>
                            <html>
                            <head>
                                <meta charset="UTF-8">
                                <style>
                                  h1 {{
                                      font-family: Helvetica-Bold;
                                      text-align: left;
                                  }}
      
                                  table {{
                                      font-family: arial, sans-serif;
                                      border-collapse: collapse;
                                      width: 100%;
                                  }}
      
                                  td, th {{
                                      text-align: center;
                                      padding: 8px;
                                      border: 1px solid #dddddd;
                                  }}
      
                                  /* Definindo o tamanho da coluna "Nome" */
                                  td:nth-child(1) {{
                                      width: 60%; 
                                  }}
      
                                  th:nth-child(1) {{
                                      width: 60%; 
                                  }}
      
                                  hr {{
                                      border: none;
                                      border-top: 1px solid black;
                                  }}
      
                                  tr:nth-child(even) {{
                                      background-color: #dddddd;
                                  }}
                        </style>
                            </head>
                            <body>
                                <h1>Fluxo de Caixa - Pagamentos</h1>
                                <h4 style="text-align: left;"  > Saldo Inicial: {self.currency(saldo_inicial)} </h4>
                                <h4 style="text-align: left;"  > Total Pago: {self.currency(pagar_total)} </h4>
                                <h4 style="text-align: left;"  > Saldo Final: {self.currency(saldo_final)} </h4>
                                <p style="text-align: left;"> Periodo: {data_inicio} - {data_fim}</p>
                                <p style="text-align: right;" >Emissão: {datetime.today().strftime("%d-%m-%Y %H:%M")}</p>
        
                                {datos}
                            """
        
            # Define o documento HTML
            self.documento = QTextDocument()
            datos = QByteArray()
            datos.append(str(reporteHtml))
            codec = QTextCodec.codecForHtml(datos)
            unistr = codec.toUnicode(datos)
            # Verifica se o texto é formato HTML e define o documento
            if Qt.mightBeRichText(unistr):
                self.documento.setHtml(unistr)
            else:
                self.documento.setPlainText(unistr)
            # Exibe mensagem se não houver dados na tabela
            if not datos:
                QMessageBox.information(self, "Preencher PDF", "Não foram encontrados dados na tabela.", QMessageBox.Ok)
        
        def vistaPreviaSaidas(self):
            self.preencher_pdf_saidas()
            
            if not self.documento.isEmpty():
                # Criar uma impressora
                printer = QPrinter(QPrinter.HighResolution)
        
                # Definir o tamanho da página (por exemplo, A4)
                printer.setPageSize(QPrinter.A4)
        
                # Definir a orientação da página (retrato ou paisagem)
                printer.setOrientation(QPrinter.Portrait)
        
                # Definir as margens (em mm)
                printer.setPageMargins(2, 2, 2, 2, QPrinter.Millimeter)
        
                # Configurar a visualização de impressão
                preview = QPrintPreviewDialog(printer)
                preview.paintRequested.connect(self.documento.print_)
                preview.exec_()
            else:
                QMessageBox.critical(self, "Vista Previa", "Não há dados para visualizar", QMessageBox.Ok)
        
        def imprimirSaidas(self):
            self.preencher_pdf_saidas()
            if not self.documento.isEmpty():
                dialogo_imprimir = QPrintDialog()
                if dialogo_imprimir.exec_() == QPrintDialog.Accepted:
                    self.documento.print_(dialogo_imprimir.printer())
            else:
                QMessageBox.critical(self, "Imprimir", "Não há dados para imprimir", QMessageBox.Ok)
        
        def exportarPDFSaidas(self):
            if not self.documento.isEmpty():
                vencimento_archivo, _ = QFileDialog.getSaveFileName(self, "Exportar a PDF", "Lista de Saídas", "Archivos PDF (*.pdf)")
                if vencimento_archivo:
                    impresion = QPrinter(QPrinter.HighResolution)
                    impresion.setOutputFormat(QPrinter.PdfFormat)
                    impresion.setOutputFileName(vencimento_archivo)
                    self.documento.print_(impresion)
            else:
                QMessageBox.critical(self, "Exportar a PDF", "No hay datos para exportar", QMessageBox.Ok)
        
        
        


    
        def salvar_tabela(self, file_path):
                        
                    try:
                        """
                        Extrai os dados da tabela atual e salva em um arquivo Excel.
                        """
                        rows = self.table.rowCount()
                        columns = self.table.columnCount()
        
                        # Criar um DataFrame vazio
                        data = []
                        for row in range(rows):
                            row_data = []
                            for col in range(columns):
                                item = self.table.item(row, col)
                                row_data.append(item.text() if item else "")
                            data.append(row_data)
        
                        column_headers = [self.table.horizontalHeaderItem(col).text() if self.table.horizontalHeaderItem(col) else f"Coluna {col}"for col in range(columns)]
        
        
                        # Adiciona a coluna "Caracteristica"
                        df = pd.DataFrame(data, columns=column_headers)
        
                        try:
                            if 'VALOR' in df.columns:
                                df['CARACTERISTICA'] = df['VALOR'].apply(lambda x: 'Pagar' if float(x) < 0 else 'Receber')
                            else:
                                QMessageBox.warning(self, "Aviso", "A coluna 'valor' não foi encontrada.")
                        except ValueError:
                            QMessageBox.warning(self, "Aviso", "A coluna 'valor' não foi encontrada.")
                            return
        
        
                        # Transformar COD de Filial em Abreviação
        
                        df['FILIAL'] = df['FILIAL'].map({'SZM':1,'SVA':2, 'SS':3, 'SNF':4})
        
                            # Transformar Tipo em AbreviAção
        
                        # Certifique-se de que é do tipo datetime
        
                        df["VENCIMENTO"] = pd.to_datetime(df["VENCIMENTO"])  
                        # formatar as datas como 'dd-mm-yyyy'
                        df["VENCIMENTO"] = df["VENCIMENTO"].dt.strftime('%d-%m-%Y')
        
                        df["VENCIMENTO"] = pd.to_datetime(df["VENCIMENTO"], format="%d-%m-%Y")  
        
                        df['VALOR'] = df['VALOR'].astype(float)
                        df['VALOR'] = [-row['VALOR'] if row['CARACTERISTICA'] == 'Pagar' else row['VALOR'] for index, row in df.iterrows()]
        
        
                        
                        # Salvar os dados no Excel
        
                        df = df[['NOME', 'DOC', 'VENCIMENTO','OBS','FILIAL', 'TIPO', 'VALOR','CARACTERISTICA']] 
                        try:
                            df.to_excel(file_path, index=False)
                            QMessageBox.information(self, "Sucesso", f"Tabela salva em: {file_path}")
                        except Exception as e:
                            error_message = "".join(traceback.format_exception(None, e, e.__traceback__))
                            print("Ocorreu um erro fatal:\n", error_message)  # Exibe o erro no terminal
                            QMessageBox.critical(None, "Erro Fatal", f"Ocorreu um erro inesperado:\n{str(e)}")
        
                    except Exception as e:
                            error_message = "".join(traceback.format_exception(None, e, e.__traceback__))
                            print("Ocorreu um erro fatal:\n", error_message)  # Exibe o erro no terminal
                            QMessageBox.critical(None, "Erro Fatal", f"Ocorreu um erro inesperado:\n{str(e)}")








    except Exception as e:
        error_message = "".join(traceback.format_exception(None, e, e.__traceback__))
        print("Ocorreu um erro fatal:\n", error_message)  # Exibe o erro no terminal
        QMessageBox.critical(None, "Erro Fatal", f"Ocorreu um erro inesperado:\n{str(e)}")



if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        form = DadosFormulario()
        form.show()
        sys.exit(app.exec_())
    
    except Exception as e:
        error_message = "".join(traceback.format_exception(None, e, e.__traceback__))
        print("Ocorreu um erro fatal:\n", error_message)  # Exibe o erro no terminal
        QMessageBox.critical(None, "Erro Fatal", f"Ocorreu um erro inesperado:\n{error_message}")


