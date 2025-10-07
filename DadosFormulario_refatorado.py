# DadosFormulario_refatorado.py (com os novos botões)

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QGroupBox, QTableWidget, QHBoxLayout, QPushButton, QLabel, QLineEdit,
                             QDateEdit, QTableWidgetItem, QSplitter, QHeaderView, QMessageBox, QDialog,
                             QComboBox) # <<< QComboBox ADICIONADO AQUI
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QTextDocument
from PyQt5.QtWebEngineWidgets import QWebEngineView
import pandas as pd
import plotly.express as px

class DadosFormularioWidget(QWidget):
    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        # Layouts principais
        main_layout = QVBoxLayout(self)
        horizontal_splitter = QSplitter(Qt.Horizontal)
        vertical_splitter = QSplitter(Qt.Vertical)
        
        # --- Lado Esquerdo (Gráfico e Tabela) ---
        self.grafico = QWebEngineView()
        self.tabela = QTableWidget()
        vertical_splitter.addWidget(self.grafico)
        vertical_splitter.addWidget(self.tabela)
        vertical_splitter.setStretchFactor(0, 4)
        vertical_splitter.setStretchFactor(1, 5)

        # --- Lado Direito (Controles e Filtros) ---
        right_widget = QWidget()
        self.layout_direito = QVBoxLayout(right_widget)
        
        # Saldo Inicial
        self.groupBox_saldo = QGroupBox("Saldo Inicial")
        saldo_layout = QHBoxLayout()
        saldo_layout.addWidget(QLabel("Saldo (R$):"))
        self.txt_saldo_inicial = QLineEdit("0,00")
        self.botao_saldo_inicial = QPushButton("Atualizar")
        saldo_layout.addWidget(self.txt_saldo_inicial)
        saldo_layout.addWidget(self.botao_saldo_inicial)
        self.groupBox_saldo.setLayout(saldo_layout)
        
        # Filtros
        self.groupBox_filtros = QGroupBox("Filtros")
        filtro_layout = QVBoxLayout()
        self.combo_filtro_nome = QComboBox()
        self.combo_filtro_filial = QComboBox()
        self.data_inicial = QDateEdit(calendarPopup=True, date=QDate.currentDate())
        self.data_final = QDateEdit(calendarPopup=True, date=QDate.currentDate().addDays(30))
        self.botao_aplicar_filtro = QPushButton("Aplicar Filtro")
        
        filtro_layout.addWidget(QLabel("Filtrar por Nome:"))
        filtro_layout.addWidget(self.combo_filtro_nome)
        filtro_layout.addWidget(QLabel("Filtrar por Filial:"))
        filtro_layout.addWidget(self.combo_filtro_filial)
        filtro_layout.addWidget(QLabel("De:"))
        filtro_layout.addWidget(self.data_inicial)
        filtro_layout.addWidget(QLabel("Até:"))
        filtro_layout.addWidget(self.data_final)
        filtro_layout.addStretch()
        filtro_layout.addWidget(self.botao_aplicar_filtro)
        self.groupBox_filtros.setLayout(filtro_layout)

        # >>> NOVO: Botões de Filtro Rápido <<<
        self.groupBox_filtros_rapidos = QGroupBox("Filtros Rápidos")
        layout_rapido = QHBoxLayout()
        self.botao_atrasados = QPushButton("Atrasados")
        self.botao_hoje = QPushButton("Hoje")
        self.botao_15_dias = QPushButton("15 Dias")
        self.botao_30_dias = QPushButton("30 Dias")
        layout_rapido.addWidget(self.botao_atrasados)
        layout_rapido.addWidget(self.botao_hoje)
        layout_rapido.addWidget(self.botao_15_dias)
        layout_rapido.addWidget(self.botao_30_dias)
        self.groupBox_filtros_rapidos.setLayout(layout_rapido)

        self.layout_direito.addWidget(self.groupBox_saldo)
        self.layout_direito.addWidget(self.groupBox_filtros)
        self.layout_direito.addWidget(self.groupBox_filtros_rapidos) # Adicionado ao layout
        self.layout_direito.addStretch()

        # Montagem final
        horizontal_splitter.addWidget(vertical_splitter)
        horizontal_splitter.addWidget(right_widget)
        horizontal_splitter.setStretchFactor(0, 5)
        horizontal_splitter.setStretchFactor(1, 1)
        main_layout.addWidget(horizontal_splitter)
        
        self._configurar_tabela()

    def _configurar_tabela(self):
        self.colunas = ['NOME', 'DOC', 'VENCIMENTO', 'Dias', 'FILIAL', 'TIPO', 'VALOR', 'SOMA', 'OBS']
        self.tabela.setColumnCount(len(self.colunas))
        self.tabela.setHorizontalHeaderLabels(self.colunas)
        self.tabela.setColumnWidth(0, 350)
        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

    def get_valores_filtros(self):
        nome_filtro = self.combo_filtro_nome.currentText()
        filial_filtro = self.combo_filtro_filial.currentText()
        return {
            "nome": nome_filtro if nome_filtro else "Todos",
            "filial": filial_filtro if filial_filtro else "Todos",
            "data_inicial": pd.to_datetime(self.data_inicial.date().toString("yyyy-MM-dd")),
            "data_final": pd.to_datetime(self.data_final.date().toString("yyyy-MM-dd"))
        }

    def atualizar_tabela(self, df: pd.DataFrame):
        self.tabela.setRowCount(0)
        if df.empty:
            return
        self.tabela.setRowCount(len(df))
        for row_idx, row_data in df.iterrows():
            for col_idx, col_name in enumerate(self.colunas):
                if col_name in df.columns:
                    valor = row_data[col_name]
                    if isinstance(valor, pd.Timestamp):
                        item_texto = valor.strftime('%d/%m/%Y')
                    elif col_name in ['VALOR', 'SOMA']:
                        item_texto = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    else:
                        item_texto = str(valor)
                    item = QTableWidgetItem(item_texto)
                    self.tabela.setItem(row_idx, col_idx, item)

    def atualizar_combos_filtro(self, df: pd.DataFrame):
        for combo, coluna in [(self.combo_filtro_nome, 'NOME'), (self.combo_filtro_filial, 'FILIAL')]:
            combo.blockSignals(True)
            texto_atual = combo.currentText()
            combo.clear()
            combo.addItem("Todos")
            if not df.empty and coluna in df.columns:
                valores_unicos = sorted(df[coluna].dropna().unique())
                combo.addItems(valores_unicos)
            index = combo.findText(texto_atual)
            if index != -1:
                combo.setCurrentIndex(index)
            combo.blockSignals(False)
    
    def atualizar_grafico(self, df: pd.DataFrame, saldo_inicial: float):
        if df.empty:
            self.grafico.setHtml("<h1>Sem dados para exibir no gráfico</h1>")
            return
        df_grafico = df[['VENCIMENTO', 'SOMA']].copy()
        df_grafico.rename(columns={'VENCIMENTO': 'Vencimento', 'SOMA': 'Valor'}, inplace=True)
        fig = px.line(df_grafico, x='Vencimento', y='Valor', title="Fluxo de Caixa Acumulado")
        fig.add_hline(y=0, line_color='red', line_dash="dash")
        fig.update_layout(yaxis_title="Valor Acumulado (R$)", plot_bgcolor='white')
        self.grafico.setHtml(fig.to_html(include_plotlyjs='cdn'))
        
    def exibir_preview_impressao(self, html_content: str):
        from PyQt5.QtPrintSupport import QPrintPreviewDialog
        documento = QTextDocument()
        documento.setHtml(html_content)
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        printer.setOrientation(QPrinter.Portrait)
        preview_dialog = QPrintPreviewDialog(printer, self)
        preview_dialog.paintRequested.connect(documento.print_)
        preview_dialog.exec_()
        




import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import tempfile
import os
from PyQt5.QtWidgets import QTableWidget, QHeaderView, QSizePolicy, QDialogButtonBox
from PyQt5.QtGui import QPixmap
        
class BalancoDialog(QDialog):
    """Um diálogo para exibir o relatório de balanço com uma tabela e um gráfico."""
    def __init__(self, dados_balanco, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Relatório de Balanço")
        self.resize(450, 550)

        layout = QVBoxLayout(self)

        # Tabela com os dados
        tabela_balanco = QTableWidget()
        tabela_balanco.setColumnCount(2)
        tabela_balanco.setRowCount(4)
        tabela_balanco.setHorizontalHeaderLabels(["Descrição", "Valor (R$)"])
        
        # Preenchendo a tabela
        for row, (descricao, valor) in enumerate(dados_balanco.items()):
            item_desc = QTableWidgetItem(descricao)
            if descricao == "ICP":
                item_valor = QTableWidgetItem(f"{valor:,.2f}")
            else:
                item_valor = QTableWidgetItem(f"R$ {valor:,.2f}")
            
            item_desc.setFlags(item_desc.flags() & ~Qt.ItemIsEditable)
            item_valor.setFlags(item_valor.flags() & ~Qt.ItemIsEditable)
            
            tabela_balanco.setItem(row, 0, item_desc)
            tabela_balanco.setItem(row, 1, item_valor)

        tabela_balanco.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tabela_balanco.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(tabela_balanco)

        # Gráfico de Barras
        canvas = self._criar_grafico(dados_balanco)
        layout.addWidget(canvas)

        # Botão OK
        botoes = QDialogButtonBox(QDialogButtonBox.Ok)
        botoes.accepted.connect(self.accept)
        layout.addWidget(botoes)

    def _criar_grafico(self, dados):
        fig, ax = plt.subplots(figsize=(5, 4))
        tipos = ["Entradas", "Saídas"]
        valores = [dados["Total de Entradas"], dados["Total de Saídas"]]
        cores = ["green", "red"]

        bars = ax.bar(tipos, valores, color=cores, width=0.5)

        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, height, f"R$ {height:,.0f}",
                    ha='center', va='bottom', fontsize=10)

        ax.set_title("Balanço Financeiro", fontsize=12, pad=10)
        ax.set_ylabel("Valor (R$)", fontsize=10)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.grid(axis="y", linestyle="--", alpha=0.7)
        fig.tight_layout()

        return FigureCanvas(fig)
    
    
    
class CurvaABCDialog(QDialog):
    """Um diálogo para exibir a tabela da Curva ABC."""
    def __init__(self, df_abc, tipo, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Curva ABC - {tipo.capitalize()}")
        self.resize(700, 500)

        layout = QVBoxLayout(self)

        self.tabela_abc = QTableWidget()
        self.tabela_abc.setColumnCount(4)
        self.tabela_abc.setHorizontalHeaderLabels(['Nome', 'Valor', '% Acumulada', 'Categoria'])
        self.tabela_abc.setColumnWidth(0, 350)
        self.tabela_abc.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)

        self.preencher_tabela(df_abc)
        layout.addWidget(self.tabela_abc)

        botoes = QDialogButtonBox(QDialogButtonBox.Ok)
        botoes.accepted.connect(self.accept)
        layout.addWidget(botoes)

    def preencher_tabela(self, df):
        self.tabela_abc.setRowCount(len(df))
        for row, row_data in df.iterrows():
            nome_item = QTableWidgetItem(row_data['NOME'])
            # Usamos a coluna 'Valor Formatado' que já criamos no core/financas.py
            valor_item = QTableWidgetItem(row_data['Valor Formatado'])
            acumulada_item = QTableWidgetItem(f"{row_data['% Acumulada']:.2f}%")
            categoria_item = QTableWidgetItem(row_data['Categoria'])

            # Tornar células não editáveis
            for item in [nome_item, valor_item, acumulada_item, categoria_item]:
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            self.tabela_abc.setItem(row, 0, nome_item)
            self.tabela_abc.setItem(row, 1, valor_item)
            self.tabela_abc.setItem(row, 2, acumulada_item)
            self.tabela_abc.setItem(row, 3, categoria_item)