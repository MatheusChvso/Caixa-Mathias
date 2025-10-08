# DadosFormulario_refatorado.py (versão final completa com Matplotlib)

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QGroupBox, QTableWidget, QHBoxLayout, QPushButton, QLabel, QLineEdit,
                             QDateEdit, QTableWidgetItem, QSplitter, QHeaderView, QMessageBox, QDialog,
                             QComboBox, QTextEdit, QAbstractItemView, QDialogButtonBox, QGridLayout, QCalendarWidget)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor, QTextCharFormat, QTextDocument, QPixmap
from PyQt5.QtPrintSupport import QPrinter
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# Classe de widget de gráfico Matplotlib
class MplCanvas(FigureCanvas):
    """Widget de base para integrar Matplotlib com PyQt5."""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig, self.axes = plt.subplots(figsize=(width, height), dpi=dpi)
        super(MplCanvas, self).__init__(fig)

class DadosFormularioWidget(QWidget):
    def __init__(self):
        super().__init__()
        self._setup_ui()
        
    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        horizontal_splitter = QSplitter(Qt.Horizontal)
        vertical_splitter = QSplitter(Qt.Vertical)
        
        # Substituímos QWebEngineView por MplCanvas
        self.grafico_canvas = MplCanvas(self, width=8, height=6, dpi=100)
        
        self.tabela = QTableWidget()
        vertical_splitter.addWidget(self.grafico_canvas)
        vertical_splitter.addWidget(self.tabela)
        vertical_splitter.setStretchFactor(0, 4)
        vertical_splitter.setStretchFactor(1, 5)

        right_widget = QWidget()
        self.layout_direito = QVBoxLayout(right_widget)
        
        self.groupBox_saldo = QGroupBox("Saldo Inicial")
        saldo_layout = QHBoxLayout()
        saldo_layout.addWidget(QLabel("Saldo (R$):"))
        self.txt_saldo_inicial = QLineEdit("0,00")
        self.botao_saldo_inicial = QPushButton("Atualizar")
        saldo_layout.addWidget(self.txt_saldo_inicial)
        saldo_layout.addWidget(self.botao_saldo_inicial)
        self.groupBox_saldo.setLayout(saldo_layout)
        
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

        self.groupBox_obs = QGroupBox("Observação da Linha Selecionada")
        layout_obs = QVBoxLayout()
        self.texto_observacao = QTextEdit()
        self.texto_observacao.setPlaceholderText("Clique em uma linha na tabela para ver a observação aqui.")
        self.botao_salvar_obs = QPushButton("Salvar Observação")
        layout_obs.addWidget(self.texto_observacao)
        layout_obs.addWidget(self.botao_salvar_obs)
        self.groupBox_obs.setLayout(layout_obs)
        
        self.groupBox_calendario = QGroupBox("Filtro por Dia")
        layout_calendario = QVBoxLayout()
        self.calendario = QCalendarWidget()
        self.calendario.setGridVisible(True)
        layout_calendario.addWidget(self.calendario)
        self.groupBox_calendario.setLayout(layout_calendario)

        self.layout_direito.addWidget(self.groupBox_saldo)
        self.layout_direito.addWidget(self.groupBox_filtros)
        self.layout_direito.addWidget(self.groupBox_filtros_rapidos)
        self.layout_direito.addWidget(self.groupBox_obs)
        self.layout_direito.addWidget(self.groupBox_calendario)
        self.layout_direito.addStretch()

        horizontal_splitter.addWidget(vertical_splitter)
        horizontal_splitter.addWidget(right_widget)
        horizontal_splitter.setStretchFactor(0, 5)
        horizontal_splitter.setStretchFactor(1, 1)
        main_layout.addWidget(horizontal_splitter)
        
        self._configurar_tabela()

    def destacar_periodo_calendario(self, data_inicio, data_fim):
        formato_padrao = QTextCharFormat()
        data_reset = self.calendario.minimumDate()
        while data_reset <= self.calendario.maximumDate():
            self.calendario.setDateTextFormat(data_reset, formato_padrao)
            data_reset = data_reset.addDays(1)
        if not data_inicio or not data_fim: return
        formato_inicio = QTextCharFormat(); formato_inicio.setBackground(QColor("lightgreen"))
        formato_fim = QTextCharFormat(); formato_fim.setBackground(QColor("lightcoral"))
        formato_intervalo = QTextCharFormat(); formato_intervalo.setBackground(QColor("lightyellow"))
        data_corrente = data_inicio
        while data_corrente <= data_fim:
            self.calendario.setDateTextFormat(data_corrente, formato_intervalo)
            data_corrente = data_corrente.addDays(1)
        self.calendario.setDateTextFormat(data_inicio, formato_inicio)
        self.calendario.setDateTextFormat(data_fim, formato_fim)

    def atualizar_grafico(self, df: pd.DataFrame, saldo_inicial: float):
        # Limpa o gráfico anterior
        self.grafico_canvas.axes.cla()

        if df.empty:
            self.grafico_canvas.axes.text(0.5, 0.5, "Sem dados para exibir no gráfico", 
                                          horizontalalignment='center', verticalalignment='center',
                                          fontsize=14, color='gray')
        else:
            # --- LÓGICA FINAL E CORRETA (RECRIA O COMPORTAMENTO ORIGINAL) ---
            # 1. Agrupa por data e pega o último saldo acumulado ('SOMA') de cada dia.
            saldo_final_diario = df.groupby(df['VENCIMENTO'].dt.date)['SOMA'].last()

            # 2. Plota as datas (índice) contra os saldos finais diários (valores).
            self.grafico_canvas.axes.plot(saldo_final_diario.index, saldo_final_diario.values, marker='o', linestyle='-', markersize=4)
            
            # Formatação do gráfico
            self.grafico_canvas.axes.axhline(0, color='red', linestyle='--')
            self.grafico_canvas.axes.set_title("Evolução do Saldo Acumulado (Final do Dia)")
            self.grafico_canvas.axes.set_ylabel("Saldo (R$)")
            self.grafico_canvas.axes.grid(True, linestyle='--', alpha=0.6)
            self.grafico_canvas.figure.autofmt_xdate() # Roda as datas para melhor visualização

        # Redesenha o gráfico na tela
        self.grafico_canvas.draw()
        
    def _configurar_tabela(self):
        self.colunas = ['NOME', 'DOC', 'VENCIMENTO', 'Dias', 'FILIAL', 'TIPO', 'VALOR', 'SOMA', 'OBS']
        self.tabela.setColumnCount(len(self.colunas)); self.tabela.setHorizontalHeaderLabels(self.colunas)
        self.tabela.setColumnWidth(0, 350)
        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.tabela.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabela.setEditTriggers(QAbstractItemView.NoEditTriggers)

    def get_valores_filtros(self):
        nome_filtro = self.combo_filtro_nome.currentText(); filial_filtro = self.combo_filtro_filial.currentText()
        return { "nome": nome_filtro if nome_filtro else "Todos", "filial": filial_filtro if filial_filtro else "Todos", "data_inicial": pd.to_datetime(self.data_inicial.date().toString("yyyy-MM-dd")), "data_final": pd.to_datetime(self.data_final.date().toString("yyyy-MM-dd")) }

    def atualizar_tabela(self, df: pd.DataFrame):
        self.tabela.setRowCount(0)
        if df.empty: return
        self.tabela.setRowCount(len(df))
        for row_idx, row_data in df.iterrows():
            for col_idx, col_name in enumerate(self.colunas):
                if col_name in df.columns:
                    valor = row_data[col_name]
                    if isinstance(valor, pd.Timestamp): item_texto = valor.strftime('%d/%m/%Y')
                    elif col_name in ['VALOR', 'SOMA']: item_texto = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    else: item_texto = str(valor) if pd.notna(valor) else ""
                    item = QTableWidgetItem(item_texto)
                    self.tabela.setItem(row_idx, col_idx, item)

    def atualizar_combos_filtro(self, df: pd.DataFrame):
        for combo, coluna in [(self.combo_filtro_nome, 'NOME'), (self.combo_filtro_filial, 'FILIAL')]:
            combo.blockSignals(True)
            texto_atual = combo.currentText()
            combo.clear(); combo.addItem("Todos")
            if not df.empty and coluna in df.columns:
                valores_unicos = sorted(df[coluna].dropna().unique())
                combo.addItems(valores_unicos)
            index = combo.findText(texto_atual)
            if index != -1: combo.setCurrentIndex(index)
            combo.blockSignals(False)
    
    def exibir_preview_impressao(self, html_content: str):
        from PyQt5.QtPrintSupport import QPrintPreviewDialog
        documento = QTextDocument(); documento.setHtml(html_content)
        printer = QPrinter(QPrinter.HighResolution); printer.setPageSize(QPrinter.A4); printer.setOrientation(QPrinter.Portrait)
        preview_dialog = QPrintPreviewDialog(printer, self)
        preview_dialog.paintRequested.connect(documento.print_)
        preview_dialog.exec_()

# --- CLASSES DE DIÁLOGO ---

class BalancoDialog(QDialog):
    def __init__(self, dados_balanco, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Relatório de Balanço"); self.resize(450, 550)
        layout = QVBoxLayout(self)
        tabela_balanco = QTableWidget(); tabela_balanco.setColumnCount(2); tabela_balanco.setRowCount(4)
        tabela_balanco.setHorizontalHeaderLabels(["Descrição", "Valor (R$)"])
        for row, (descricao, valor) in enumerate(dados_balanco.items()):
            item_desc = QTableWidgetItem(descricao)
            item_valor = QTableWidgetItem(f"{valor:,.2f}" if descricao == "ICP" else f"R$ {valor:,.2f}")
            item_desc.setFlags(item_desc.flags() & ~Qt.ItemIsEditable); item_valor.setFlags(item_valor.flags() & ~Qt.ItemIsEditable)
            tabela_balanco.setItem(row, 0, item_desc); tabela_balanco.setItem(row, 1, item_valor)
        tabela_balanco.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(tabela_balanco)
        canvas = self._criar_grafico(dados_balanco)
        layout.addWidget(canvas)
        botoes = QDialogButtonBox(QDialogButtonBox.Ok); botoes.accepted.connect(self.accept)
        layout.addWidget(botoes)

    def _criar_grafico(self, dados):
        fig, ax = plt.subplots(figsize=(5, 4))
        tipos = ["Entradas", "Saídas"]; valores = [dados["Total de Entradas"], dados["Total de Saídas"]]; cores = ["green", "red"]
        bars = ax.bar(tipos, valores, color=cores, width=0.5)
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, height, f"R$ {height:,.0f}", ha='center', va='bottom', fontsize=10)
        ax.set_title("Balanço Financeiro", fontsize=12, pad=10); ax.set_ylabel("Valor (R$)", fontsize=10)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.grid(axis="y", linestyle="--", alpha=0.7)
        fig.tight_layout()
        return FigureCanvas(fig)

class CurvaABCDialog(QDialog):
    def __init__(self, df_abc, tipo, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Curva ABC - {tipo.capitalize()}"); self.resize(700, 500)
        layout = QVBoxLayout(self)
        self.tabela_abc = QTableWidget(); self.tabela_abc.setColumnCount(4)
        self.tabela_abc.setHorizontalHeaderLabels(['Nome', 'Valor', '% Acumulada', 'Categoria'])
        self.tabela_abc.setColumnWidth(0, 350); self.tabela_abc.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.preencher_tabela(df_abc)
        layout.addWidget(self.tabela_abc)
        botoes = QDialogButtonBox(QDialogButtonBox.Ok); botoes.accepted.connect(self.accept)
        layout.addWidget(botoes)

    def preencher_tabela(self, df):
        self.tabela_abc.setRowCount(len(df))
        for row, row_data in df.iterrows():
            nome_item = QTableWidgetItem(row_data['NOME']); valor_item = QTableWidgetItem(row_data['Valor Formatado'])
            acumulada_item = QTableWidgetItem(f"{row_data['% Acumulada']:.2f}%"); categoria_item = QTableWidgetItem(row_data['Categoria'])
            for item in [nome_item, valor_item, acumulada_item, categoria_item]: item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.tabela_abc.setItem(row, 0, nome_item); self.tabela_abc.setItem(row, 1, valor_item); self.tabela_abc.setItem(row, 2, acumulada_item); self.tabela_abc.setItem(row, 3, categoria_item)

class CadastroDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cadastrar Novo Item"); self.setMinimumWidth(400)
        layout = QGridLayout(self)
        self.txt_nome = QLineEdit(); self.txt_doc = QLineEdit(); self.txt_vencimento = QDateEdit(calendarPopup=True, date=QDate.currentDate())
        self.combo_filial = QComboBox(); self.combo_filial.addItems(['SZM', 'SVA', 'SS', 'SNF'])
        self.txt_tipo = QLineEdit(); self.txt_valor = QLineEdit(); self.txt_valor.setPlaceholderText("Use valor negativo para saídas (ex: -150,50)")
        self.combo_caracteristica = QComboBox(); self.combo_caracteristica.addItems(['Receber', 'Pagar']); self.txt_obs = QTextEdit()
        layout.addWidget(QLabel("Nome:"), 0, 0); layout.addWidget(self.txt_nome, 0, 1)
        layout.addWidget(QLabel("Doc:"), 1, 0); layout.addWidget(self.txt_doc, 1, 1)
        layout.addWidget(QLabel("Vencimento:"), 2, 0); layout.addWidget(self.txt_vencimento, 2, 1)
        layout.addWidget(QLabel("Filial:"), 3, 0); layout.addWidget(self.combo_filial, 3, 1)
        layout.addWidget(QLabel("Tipo:"), 4, 0); layout.addWidget(self.txt_tipo, 4, 1)
        layout.addWidget(QLabel("Característica:"), 5, 0); layout.addWidget(self.combo_caracteristica, 5, 1)
        layout.addWidget(QLabel("Valor (R$):"), 6, 0); layout.addWidget(self.txt_valor, 6, 1)
        layout.addWidget(QLabel("Observação:"), 7, 0); layout.addWidget(self.txt_obs, 7, 1)
        botoes = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        botoes.accepted.connect(self.accept); botoes.rejected.connect(self.reject)
        layout.addWidget(botoes, 8, 0, 1, 2)
        self.dados = None

    def accept(self):
        if not self.txt_nome.text() or not self.txt_valor.text():
            QMessageBox.warning(self, "Campos Obrigatórios", "Os campos 'Nome' e 'Valor' são obrigatórios."); return
        try: valor = float(self.txt_valor.text().replace(',', '.'))
        except ValueError: QMessageBox.warning(self, "Valor Inválido", "O campo 'Valor' deve ser um número válido."); return
        caracteristica = self.combo_caracteristica.currentText()
        if (caracteristica == 'Pagar' and valor > 0) or (caracteristica == 'Receber' and valor < 0): valor *= -1
        self.dados = { "NOME": self.txt_nome.text(), "DOC": self.txt_doc.text(), "VENCIMENTO": pd.to_datetime(self.txt_vencimento.date().toString("yyyy-MM-dd")), "FILIAL": self.combo_filial.currentText(), "TIPO": self.txt_tipo.text(), "VALOR": valor, "OBS": self.txt_obs.toPlainText() }
        super().accept()

    def get_dados(self):
        return self.dados

class FluxoGraficoDialog(QDialog):
    def __init__(self, df_agrupado, titulo, parent=None):
        super().__init__(parent)
        self.setWindowTitle(titulo); self.resize(800, 600)
        layout = QVBoxLayout(self)
        canvas = self._criar_grafico(df_agrupado, titulo)
        layout.addWidget(canvas)
        botoes = QDialogButtonBox(QDialogButtonBox.Ok); botoes.accepted.connect(self.accept)
        layout.addWidget(botoes)

    def _criar_grafico(self, df, titulo):
        fig, ax = plt.subplots(figsize=(10, 6))
        indices = df.index; bar_width = 0.35
        r1 = range(len(indices)); r2 = [x + bar_width for x in r1]
        ax.bar(r1, df['Entradas'], color='green', width=bar_width, edgecolor='grey', label='Entradas')
        ax.bar(r2, df['Saídas'].abs(), color='red', width=bar_width, edgecolor='grey', label='Saídas')
        ax.set_xlabel('Período', fontweight='bold'); ax.set_ylabel('Valor (R$)', fontweight='bold')
        ax.set_title(titulo, fontsize=14)
        ax.set_xticks([r + bar_width/2 for r in range(len(indices))])
        ax.set_xticklabels(df['Label'], rotation=45, ha='right')
        ax.legend()
        fig.tight_layout()
        return FigureCanvas(fig)