import sys
import os
from PyQt5.QtWidgets import *
from PyQt5.Qt import Qt
from DadosFormulario import DadosFormulario
from PyQt5.QtCore import QTimer
import traceback

## TESTE DE BRANCH

os.environ["QTWEBENGINE_DISABLE_SANDBOX"] = "1"

class MainWindow(QMainWindow):
    def __init__(self, widget):
        super().__init__()

        self.setWindowTitle('GRF v1.4')  # Titulo da Janela
        self.resize(1600, 850)

        self.setWindowFlag

        self.menuBar = self.menuBar()
        self.fileMenu = self.menuBar.addMenu('Arquivo')
        self.RelatorioMenu = self.menuBar.addMenu('Relatório')
        self.CadastroMenu = self.menuBar.addMenu('Cadastro')
        self.FiltroMenu = self.menuBar.addMenu('Filtro')
        self.calendarMenu = self.menuBar.addMenu('Calendário')

         # Inicialmente desativar o menu de filtro
        self.FiltroMenu.setEnabled(False)

        
        # Adicionar calendário

        calendario = QAction('Calendário', self)
        calendario.setShortcut('Ctrl+D')
        calendario.triggered.connect(lambda: widget.gerar_calendario())
        self.calendarMenu.addAction(calendario)

        
        # Adicionar botão para filtrar
        filtrar = QAction('Filtrar', self)
        filtrar.setShortcut('Ctrl+F')
        filtrar.triggered.connect(lambda: DadosFormulario.filtrar_tabela(widget))
        self.FiltroMenu.addAction(filtrar)

        # Adicionar botão para cadastrar
        cadastrar = QAction('Cadastrar', self)
        cadastrar.setShortcut('Ctrl+N')
        cadastrar.triggered.connect(lambda: DadosFormulario.cadastrar(widget))
        self.CadastroMenu.addAction(cadastrar)

        # Adicionar botão para carregar arquivos Excel
        carregar_excel = QAction('Carregar Excel', self)
        carregar_excel.setShortcut('Ctrl+O')
        carregar_excel.triggered.connect(lambda: self.carregar_arquivos_excel(widget))  # Conectando ao método para carregar arquivos Excel
        self.fileMenu.addAction(carregar_excel)


        # Salvar tabela em Excel
        salvar_excel = QAction('Salvar como Excel', self)
        salvar_excel.setShortcut('Ctrl+S')
        salvar_excel.triggered.connect(lambda: self.salvar_tabela_excel(widget))
        self.fileMenu.addAction(salvar_excel)

        #Imprimir pagamentos
        pagamentos = QAction('Imprimir Pagamentos', self)
        pagamentos.setShortcut('Ctrl+shift+P')
        pagamentos.triggered.connect(lambda: DadosFormulario.vistaPreviaSaidas(widget))
        self.RelatorioMenu.addAction(pagamentos)

        # Adicionar botão para sair
        sair = QAction('Sair', self)
        sair.setShortcut('Ctrl+Q')
        sair.triggered.connect(lambda: app.quit())

        self.fileMenu.addAction(sair)

        self.setCentralWidget(widget)

        # Imprimir Fluxo de Caixa
        imprime_fc = QAction('Imprimir Fluxo', self)
        imprime_fc.setShortcut('Ctrl+P')
        imprime_fc.triggered.connect(lambda: DadosFormulario.vistaPrevia(widget))
        self.RelatorioMenu.addAction(imprime_fc)

         # Menu para alternar entre entrada e saída

        curva_abc_entrada = QAction('Curva ABC - Entradas', self)
        curva_abc_entrada.setShortcut('Ctrl+i')
        curva_abc_entrada.triggered.connect(lambda: widget.gerar_curva_abc('entrada'))
        self.RelatorioMenu.addAction(curva_abc_entrada)

        curva_abc_saida = QAction('Curva ABC - Saídas', self)
        curva_abc_saida.setShortcut('Ctrl+shift+s')
        curva_abc_saida.triggered.connect(lambda: widget.gerar_curva_abc('saida'))
        self.RelatorioMenu.addAction(curva_abc_saida)

        # Relatório de balanço

        relatorio_balanco = QAction('Balanço', self)
        relatorio_balanco.setShortcut('Ctrl+B')
        relatorio_balanco.triggered.connect(lambda: widget.gerar_relatorio_balanco())
        self.RelatorioMenu.addAction(relatorio_balanco)
        
        
        

        # Relatório balanço Diário
        self.botao_fluxo_diario = QAction("Fluxo Diário")
        self.botao_fluxo_diario.triggered.connect(lambda: widget.gerar_relatorio_fluxo_diario())
        self.RelatorioMenu.addAction(self.botao_fluxo_diario)

        relatorio_fluxo_mensal = QAction('Fluxo Mensal', self)
        relatorio_fluxo_mensal.setShortcut('Ctrl+M')
        relatorio_fluxo_mensal.triggered.connect(lambda: widget.gerar_relatorio_fluxo_mensal())
        self.RelatorioMenu.addAction(relatorio_fluxo_mensal)
        

    def enable_filter_menu(self, enable):
            self.FiltroMenu.setEnabled(enable)

    # Método para carregar arquivos Excel
    def carregar_arquivos_excel(self, widget):
        # Abrir uma janela de diálogo para selecionar vários arquivos Excel
        files, _ = QFileDialog.getOpenFileNames(self, "Selecione Arquivos", "", "Arquivos Excel (*.xlsx *.xls)")
        if files:
            widget.preencher_tabela(files)  # Chamar o método para preencher a tabela
        else:
            print("Nenhum arquivo selecionado.")

    def salvar_tabela_excel(self, widget):
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Tabela", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            widget.salvar_tabela(file_path)
        else:
            print("Salvamento cancelado.")


if __name__ == '__main__':

    try:
        app = QApplication(sys.argv)
        form = DadosFormulario()
        main_window = MainWindow(form)
        form.main_window = main_window  # Passar a referência da janela principal para o formulário
        main_window.show()
        sys.exit(app.exec_())

    except Exception as e:
        error_message = "".join(traceback.format_exception(None, e, e.__traceback__))
        print("Ocorreu um erro fatal:\n", error_message)  # Exibe o erro no terminal
        QMessageBox.critical(None, "Erro Fatal", f"Ocorreu um erro inesperado:\n{error_message}")