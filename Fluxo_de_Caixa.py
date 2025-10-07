# Fluxo_de_Caixa.py (com Curva ABC)

import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QDate

# Importando o novo diálogo
from DadosFormulario_refatorado import DadosFormularioWidget, BalancoDialog, CurvaABCDialog
from data import excel_handler
from core import financas, relatorios, utils

os.environ["QTWEBENGINE_DISABLE_SANDBOX"] = "1"

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('GRF v1.9 (Curva ABC)')
        self.resize(1600, 850)
        self.df_dados_completos = pd.DataFrame()
        self.df_dados_filtrados = pd.DataFrame()
        self.view = DadosFormularioWidget()
        self.setCentralWidget(self.view)
        self._criar_menus()
        self._conectar_sinais()

    def _criar_menus(self):
        menu_bar = self.menuBar() 
        self.fileMenu = menu_bar.addMenu('Arquivo')
        
        carregar_action = QAction('Carregar Excel', self)
        carregar_action.setShortcut('Ctrl+O')
        carregar_action.triggered.connect(self.carregar_arquivos_excel)
        self.fileMenu.addAction(carregar_action)

        salvar_action = QAction('Salvar como Excel', self)
        salvar_action.setShortcut('Ctrl+S')
        salvar_action.triggered.connect(self.salvar_tabela_excel)
        self.fileMenu.addAction(salvar_action)

        self.relatorioMenu = menu_bar.addMenu('Relatório')
        imprimir_fluxo_action = QAction('Visualizar Fluxo de Caixa', self)
        imprimir_fluxo_action.setShortcut('Ctrl+P')
        imprimir_fluxo_action.triggered.connect(self.visualizar_relatorio_fluxo_caixa)
        self.relatorioMenu.addAction(imprimir_fluxo_action)
        
        balanco_action = QAction('Balanço', self)
        balanco_action.setShortcut('Ctrl+B')
        balanco_action.triggered.connect(self._exibir_relatorio_balanco)
        self.relatorioMenu.addAction(balanco_action)

        # >>> NOVO: Adicionando as ações de Curva ABC <<<
        self.relatorioMenu.addSeparator()
        abc_entrada_action = QAction('Curva ABC - Entradas', self)
        abc_entrada_action.triggered.connect(lambda: self._exibir_curva_abc('entrada'))
        self.relatorioMenu.addAction(abc_entrada_action)

        abc_saida_action = QAction('Curva ABC - Saídas', self)
        abc_saida_action.triggered.connect(lambda: self._exibir_curva_abc('saida'))
        self.relatorioMenu.addAction(abc_saida_action)


    def _conectar_sinais(self):
        self.view.botao_aplicar_filtro.clicked.connect(self.aplicar_filtros_e_atualizar_view)
        self.view.botao_saldo_inicial.clicked.connect(self.aplicar_filtros_e_atualizar_view)
        self.view.botao_hoje.clicked.connect(self._filtrar_por_hoje)
        self.view.botao_atrasados.clicked.connect(self._filtrar_por_atrasados)
        self.view.botao_15_dias.clicked.connect(lambda: self._filtrar_proximos_dias(15))
        self.view.botao_30_dias.clicked.connect(lambda: self._filtrar_proximos_dias(30))

    # --- NOVO MÉTODO PARA O RELATÓRIO DE CURVA ABC ---
    def _exibir_curva_abc(self, tipo):
        if self.df_dados_filtrados.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados filtrados para gerar a Curva ABC.")
            return

        # 1. Chamar a função de cálculo do nosso módulo de finanças
        df_resultado_abc = financas.gerar_curva_abc(self.df_dados_filtrados, tipo)

        if df_resultado_abc.empty:
            QMessageBox.information(self, "Aviso", f"Não foram encontradas '{tipo}s' no período selecionado.")
            return

        # 2. Criar e exibir o diálogo
        dialog = CurvaABCDialog(df_resultado_abc, tipo, self)
        dialog.exec_()

    # (O resto do arquivo continua o mesmo)
    def _exibir_relatorio_balanco(self):
        if self.df_dados_filtrados.empty:
            QMessageBox.warning(self, "Aviso", "Não há dados filtrados para gerar o balanço.")
            return
        entradas, saidas, saldo, icp = financas.calcular_balanco(self.df_dados_filtrados)
        dados_formatados = {"Total de Entradas": entradas, "Total de Saídas": saidas, "Saldo Final": saldo, "ICP": icp}
        dialog = BalancoDialog(dados_formatados, self)
        dialog.exec_()

    def _filtrar_por_hoje(self):
        hoje = QDate.currentDate()
        self.view.data_inicial.setDate(hoje)
        self.view.data_final.setDate(hoje)
        self.aplicar_filtros_e_atualizar_view()

    def _filtrar_por_atrasados(self):
        hoje = QDate.currentDate()
        data_inicio_atrasados = hoje.addYears(-10)
        data_fim_atrasados = hoje.addDays(-1)
        self.view.data_inicial.setDate(data_inicio_atrasados)
        self.view.data_final.setDate(data_fim_atrasados)
        self.aplicar_filtros_e_atualizar_view()

    def _filtrar_proximos_dias(self, dias):
        hoje = QDate.currentDate()
        data_futura = hoje.addDays(dias)
        self.view.data_inicial.setDate(hoje)
        self.view.data_final.setDate(data_futura)
        self.aplicar_filtros_e_atualizar_view()

    def carregar_arquivos_excel(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos Excel", "", "Excel Files (*.xlsx *.xls)")
        if files:
            self.df_dados_completos = excel_handler.carregar_dados_de_excel(files)
            if not self.df_dados_completos.empty:
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
        df = self.df_dados_completos.copy()
        if primeira_carga and not df.empty:
            data_min = df['VENCIMENTO'].min()
            data_max = df['VENCIMENTO'].max()
            self.view.data_inicial.setDate(data_min)
            self.view.data_final.setDate(data_max)
        filtros = self.view.get_valores_filtros()
        if filtros['nome'] != "Todos":
            df = df[df['NOME'] == filtros['nome']]
        if filtros['filial'] != "Todos":
            df = df[df['FILIAL'] == filtros['filial']]
        df = df[(df['VENCIMENTO'] >= filtros['data_inicial']) & (df['VENCIMENTO'] <= filtros['data_final'])]
        if df.empty and not primeira_carga:
            QMessageBox.information(self, "Aviso", "Nenhum registro encontrado para os filtros aplicados.")
        self._atualizar_view_com_dados(df)

    def _atualizar_view_com_dados(self, df_filtrado):
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