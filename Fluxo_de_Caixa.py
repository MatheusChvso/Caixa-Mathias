# Fluxo_de_Caixa.py (VERSÃO DE DIAGNÓSTICO)

import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt

# Importando nossos módulos refatorados!
from DadosFormulario_refatorado import DadosFormularioWidget
from data import excel_handler
from core import financas, relatorios, utils

os.environ["QTWEBENGINE_DISABLE_SANDBOX"] = "1"

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('GRF v1.6 (Diagnóstico)')
        self.resize(1600, 850)
        self.df_dados_completos = pd.DataFrame()
        self.df_dados_filtrados = pd.DataFrame()
        self.view = DadosFormularioWidget()
        self.setCentralWidget(self.view)
        self._criar_menus()
        self._conectar_sinais()

    def _criar_menus(self):
        self.menuBar = self.menuBar()
        self.fileMenu = self.menuBar.addMenu('Arquivo')
        carregar_action = QAction('Carregar Excel', self)
        carregar_action.setShortcut('Ctrl+O')
        carregar_action.triggered.connect(self.carregar_arquivos_excel)
        self.fileMenu.addAction(carregar_action)
        salvar_action = QAction('Salvar como Excel', self)
        salvar_action.setShortcut('Ctrl+S')
        salvar_action.triggered.connect(self.salvar_tabela_excel)
        self.fileMenu.addAction(salvar_action)
        self.relatorioMenu = self.menuBar.addMenu('Relatório')
        imprimir_fluxo_action = QAction('Visualizar Fluxo de Caixa', self)
        imprimir_fluxo_action.setShortcut('Ctrl+P')
        imprimir_fluxo_action.triggered.connect(self.visualizar_relatorio_fluxo_caixa)
        self.relatorioMenu.addAction(imprimir_fluxo_action)

    def _conectar_sinais(self):
        self.view.botao_aplicar_filtro.clicked.connect(self.aplicar_filtros_e_atualizar_view)
        self.view.botao_saldo_inicial.clicked.connect(self.aplicar_filtros_e_atualizar_view)

    def carregar_arquivos_excel(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos Excel", "", "Excel Files (*.xlsx *.xls)")
        if files:
            self.df_dados_completos = excel_handler.carregar_dados_de_excel(files)
            if not self.df_dados_completos.empty:
                print("--- DADOS CARREGADOS COM SUCESSO ---")
                print(f"Total de registros carregados: {len(self.df_dados_completos)}")
                print(f"Tipo da coluna VENCIMENTO: {self.df_dados_completos['VENCIMENTO'].dtype}")
                self.aplicar_filtros_e_atualizar_view(primeira_carga=True)
                QMessageBox.information(self, "Sucesso", f"{len(self.df_dados_completos)} registros carregados.")

    def salvar_tabela_excel(self):
        if self.df_dados_filtrados.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados filtrados para salvar.")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Tabela", "", "Excel Files (*.xlsx)")
        if file_path:
            excel_handler.salvar_dados_para_excel(self.df_dados_filtrados, file_path)

    def aplicar_filtros_e_atualizar_view(self, primeira_carga=False):
        if self.df_dados_completos.empty:
            return

        print("\n--- APLICANDO FILTROS ---")
        df = self.df_dados_completos.copy()

        if primeira_carga and not df.empty:
            data_min = df['VENCIMENTO'].min()
            data_max = df['VENCIMENTO'].max()
            print(f"Detectado período dos dados: de {data_min.strftime('%Y-%m-%d')} a {data_max.strftime('%Y-%m-%d')}")
            print("Ajustando os calendários da interface para este período.")
            self.view.data_inicial.setDate(data_min)
            self.view.data_final.setDate(data_max)

        filtros = self.view.get_valores_filtros()
        print(f"Filtros da interface: {filtros}")
        
        # Filtro de data
        print(f"Registros antes do filtro de data: {len(df)}")
        df = df[(df['VENCIMENTO'] >= filtros['data_inicial']) & (df['VENCIMENTO'] <= filtros['data_final'])]
        print(f"Registros APÓS o filtro de data: {len(df)}")
        
        # Outros filtros
        if filtros['nome'] != "Todos":
            df = df[df['NOME'] == filtros['nome']]
        if filtros['filial'] != "Todos":
            df = df[df['FILIAL'] == filtros['filial']]

        print(f"Registros após TODOS os filtros: {len(df)}")

        if df.empty and not primeira_carga:
            QMessageBox.information(self, "Aviso", "Nenhum registro encontrado para os filtros aplicados.")

        self._atualizar_view_com_dados(df)

    def _atualizar_view_com_dados(self, df_filtrado):
        print(f"--- ATUALIZANDO VIEW com {len(df_filtrado)} registros ---")
        try:
            saldo_inicial = float(self.view.txt_saldo_inicial.text().replace(',', '.'))
        except ValueError:
            saldo_inicial = 0.0

        df_para_processar = df_filtrado.sort_values(by='VENCIMENTO').copy()
        if not df_para_processar.empty:
            df_para_processar['SOMA'] = saldo_inicial + df_para_processar['VALOR'].cumsum()

        self.df_dados_filtrados = df_para_processar
        self.view.atualizar_tabela(self.df_dados_filtrados)
        self.view.atualizar_combos_filtro(self.df_dados_completos)
        self.view.atualizar_grafico(self.df_dados_filtrados, saldo_inicial)

    def visualizar_relatorio_fluxo_caixa(self):
        if self.df_dados_filtrados.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados para gerar o relatório.")
            return
        df_relatorio = self.df_dados_filtrados.rename(columns={'NOME': 'nome', 'DOC': 'doc', 'FILIAL': 'filial', 'VENCIMENTO': 'vencimento', 'VALOR': 'valor', 'SOMA': 'soma'})
        try:
            saldo_inicial = float(self.view.txt_saldo_inicial.text().replace(',', '.'))
        except ValueError:
            saldo_inicial = 0.0
        html = relatorios.gerar_html_fluxo_caixa_completo(df_relatorio, saldo_inicial)
        self.view.exibir_preview_impressao(html)

if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        main_window = MainWindow()
        main_window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Ocorreu um erro fatal: {e}")
        import traceback
        traceback.print_exc()
        QMessageBox.critical(None, "Erro Fatal", f"Ocorreu um erro inesperado:\n{e}")