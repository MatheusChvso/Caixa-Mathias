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