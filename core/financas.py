# core/financas.py

import pandas as pd

# Importamos nossa nova função utilitária
from core.utils import formatar_moeda

def calcular_balanco(df: pd.DataFrame):
    """
    Calcula o balanço financeiro a partir de um DataFrame.
    O DataFrame deve conter uma coluna 'VALOR'.
    """
    if df.empty:
        return 0.0, 0.0, 0.0, 0.0

    total_entradas = df[df['VALOR'] >= 0]['VALOR'].sum()
    total_saidas = abs(df[df['VALOR'] < 0]['VALOR'].sum())
    saldo_final = total_entradas - total_saidas
    icp = total_entradas / total_saidas if total_saidas > 0 else 0.0

    return total_entradas, total_saidas, saldo_final, icp


def calcular_drawdown(df_fluxo_caixa: pd.DataFrame):
    """
    Calcula o drawdown da soma acumulada em um DataFrame de fluxo de caixa.
    O DataFrame deve conter uma coluna 'Valor' com a soma acumulada.
    """
    if df_fluxo_caixa.empty:
        return 0.0

    soma_acumulada = df_fluxo_caixa['Valor']
    max_pico = soma_acumulada.expanding().max()
    drawdowns = soma_acumulada - max_pico
    max_drawdown = drawdowns.min()

    return max_drawdown if pd.notna(max_drawdown) else 0.0


def calcular_tempo_caixa_negativo(df_fluxo_caixa: pd.DataFrame):
    """
    Calcula a quantidade de dias e o percentual em que o saldo acumulado ficou negativo.
    O DataFrame deve ter as colunas 'Vencimento' e 'Valor' (soma acumulada).
    """
    if df_fluxo_caixa.empty:
        return 0, 0.0

    dias_negativos = df_fluxo_caixa[df_fluxo_caixa['Valor'] < 0]['Vencimento'].nunique()
    total_dias = df_fluxo_caixa['Vencimento'].nunique()

    percentual = (dias_negativos / total_dias) * 100 if total_dias > 0 else 0
    return dias_negativos, percentual


def gerar_curva_abc(df: pd.DataFrame, tipo: str):
    """
    Gera a curva ABC baseada no tipo (entrada ou saida) a partir de um DataFrame.
    O DataFrame deve conter as colunas 'NOME' e 'VALOR'.
    """
    if tipo == 'entrada':
        df_filtrado = df[df['VALOR'] > 0].copy()
    elif tipo == 'saida':
        df_filtrado = df[df['VALOR'] < 0].copy()
    else:
        return pd.DataFrame()

    df_filtrado['VALOR'] = df_filtrado['VALOR'].abs()

    if df_filtrado.empty:
        return pd.DataFrame()

    df_agrupado = df_filtrado.groupby('NOME')['VALOR'].sum().sort_values(ascending=False).reset_index()

    total_geral = df_agrupado['VALOR'].sum()
    df_agrupado['%'] = (df_agrupado['VALOR'] / total_geral) * 100
    df_agrupado['% Acumulada'] = df_agrupado['%'].cumsum()

    def classificar(percentual):
        if percentual <= 80:
            return 'A'
        elif percentual <= 95:
            return 'B'
        else:
            return 'C'

    df_agrupado['Categoria'] = df_agrupado['% Acumulada'].apply(classificar)
    df_agrupado['Valor Formatado'] = df_agrupado['VALOR'].apply(formatar_moeda)

    return df_agrupado