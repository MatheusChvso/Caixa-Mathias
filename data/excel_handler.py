# data/excel_handler.py

import pandas as pd
import os
from PyQt5.QtWidgets import QMessageBox

def carregar_dados_de_excel(lista_arquivos: list):
    """
    VERSÃO DEFINITIVA: Lê o Excel sem 'DIAS' ou 'DOC VENCIMENTO' e cria essas
    colunas programaticamente.
    """
    if not lista_arquivos:
        return pd.DataFrame()

    lista_dfs = []
    for arquivo in lista_arquivos:
        if os.path.exists(arquivo):
            try:
                df = pd.read_excel(arquivo)
                df.columns = df.columns.str.strip() # Remove espaços dos nomes das colunas
                lista_dfs.append(df)
            except Exception as e:
                QMessageBox.warning(None, "Erro de Leitura", f"Não foi possível ler o arquivo:\n{arquivo}\n\nErro: {e}")
                continue

    if not lista_dfs:
        QMessageBox.information(None, "Aviso", "Nenhum dado foi carregado dos arquivos selecionados.")
        return pd.DataFrame()

    df_final = pd.concat(lista_dfs, ignore_index=True)

    # 1. Verificar se colunas essenciais existem (AGORA SEM 'DIAS' OU 'DOC VENCIMENTO')
    colunas_necessarias = {'VALOR', 'CARACTERISTICA', 'FILIAL', 'VENCIMENTO', 'NOME', 'OBS', 'DOC', 'TIPO'}
    colunas_no_excel = set(df_final.columns)
    colunas_faltando = colunas_necessarias - colunas_no_excel

    if colunas_faltando:
        mensagem_erro = f"As seguintes colunas obrigatórias não foram encontradas no seu arquivo Excel:\n\n{', '.join(colunas_faltando)}"
        QMessageBox.critical(None, "Erro de Coluna", mensagem_erro)
        return pd.DataFrame()

    # --- Início da Lógica de Transformação de Dados ---
    
    # 2. Converter 'VENCIMENTO' para o tipo data
    df_final['VENCIMENTO'] = pd.to_datetime(df_final['VENCIMENTO'], format='%d/%m/%Y', errors='coerce')

    if df_final['VENCIMENTO'].isnull().any():
        QMessageBox.warning(None, "Aviso", "Algumas datas de vencimento no Excel são inválidas e serão ignoradas.")
        df_final.dropna(subset=['VENCIMENTO'], inplace=True)

    # 3. >>> A LÓGICA CORRETA AGORA <<<
    # Cria 'DOC VENCIMENTO' como uma cópia de 'VENCIMENTO'
    df_final['DOC VENCIMENTO'] = df_final['VENCIMENTO']
    # Cria 'Dias' com base na diferença (que será 0 no início)
    df_final['Dias'] = (df_final['DOC VENCIMENTO'] - df_final['VENCIMENTO']).dt.days

    # 4. Transformar valores a pagar em negativo
    df_final['VALOR'] = df_final.apply(
        lambda row: -row['VALOR'] if row['CARACTERISTICA'] == 'Pagar' else row['VALOR'],
        axis=1
    )

    # 5. Mapear código da Filial para abreviação
    mapa_filial = {1: 'SZM', 2: 'SVA', 3: 'SS', 4: 'SNF'}
    df_final['FILIAL'] = df_final['FILIAL'].map(mapa_filial)

    # 6. Ordenar valores
    df_final = df_final.sort_values(by=['VENCIMENTO', 'NOME'], ascending=[True, True])

    # 7. Selecionar e reordenar as colunas que a aplicação usará
    colunas_finais = ['NOME', 'DOC VENCIMENTO', 'DOC', 'VENCIMENTO', 'Dias', 'FILIAL', 'TIPO', 'VALOR', 'OBS']
    df_final = df_final[colunas_finais]
    
    return df_final


def salvar_dados_para_excel(df: pd.DataFrame, caminho_arquivo: str):
    """
    Recebe um DataFrame e o salva em um arquivo Excel no caminho especificado.
    Toda a lógica de 'salvar_tabela' foi movida para cá.
    """
    if df.empty:
        QMessageBox.warning(None, "Aviso", "A tabela está vazia. Nada para salvar.")
        return False

    try:
        # Criamos uma cópia para não alterar o DataFrame original que está em uso no programa
        df_para_salvar = df.copy()

        # --- Início da Lógica de Reversão de Dados ---
        # (Faz o processo inverso da leitura para salvar no formato original)

        # 1. Adicionar a coluna "CARACTERISTICA"
        df_para_salvar['CARACTERISTICA'] = df_para_salvar['VALOR'].apply(lambda x: 'Pagar' if float(x) < 0 else 'Receber')

        # 2. Mapear abreviação da Filial de volta para o código
        mapa_filial_reverso = {'SZM': 1, 'SVA': 2, 'SS': 3, 'SNF': 4}
        df_para_salvar['FILIAL'] = df_para_salvar['FILIAL'].map(mapa_filial_reverso)

        # 3. Garantir que o valor seja sempre positivo no Excel
        df_para_salvar['VALOR'] = df_para_salvar['VALOR'].abs()
        
        # 4. Formatar a data para o padrão dd-mm-yyyy
        df_para_salvar["VENCIMENTO"] = pd.to_datetime(df_para_salvar["VENCIMENTO"]).dt.strftime('%d-%m-%Y')

        # 5. Selecionar e reordenar as colunas para o formato de salvamento
        colunas_salvamento = ['NOME', 'DOC', 'VENCIMENTO', 'OBS', 'FILIAL', 'TIPO', 'VALOR', 'CARACTERISTICA']
        df_para_salvar = df_para_salvar[colunas_salvamento]

        # 6. Salvar o arquivo
        df_para_salvar.to_excel(caminho_arquivo, index=False)
        QMessageBox.information(None, "Sucesso", f"Tabela salva com sucesso em:\n{caminho_arquivo}")
        return True

    except Exception as e:
        QMessageBox.critical(None, "Erro Fatal ao Salvar", f"Ocorreu um erro inesperado ao salvar o arquivo:\n{str(e)}")
        return False