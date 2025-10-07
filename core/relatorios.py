# core/relatorios.py

from datetime import datetime
import pandas as pd
from .utils import formatar_moeda # Usando nosso formatador de moeda!

def gerar_html_fluxo_caixa_completo(df: pd.DataFrame, saldo_inicial: float):
    """
    Gera uma string HTML para o relatório de fluxo de caixa completo.
    """
    if df.empty:
        return "<h1>Nenhum dado para exibir no período selecionado.</h1>"

    # Preparação dos dados para o relatório
    df_relatorio = df.copy()
    df_relatorio['valor'] = pd.to_numeric(df_relatorio['valor'], errors='coerce')
    df_relatorio['pagar'] = df_relatorio['valor'].apply(lambda x: x if x < 0 else None)
    df_relatorio['receber'] = df_relatorio['valor'].apply(lambda x: x if x >= 0 else None)
    df_relatorio['vencimento'] = pd.to_datetime(df_relatorio['vencimento']).dt.strftime('%d/%m/%Y')
    
    saldo_final_str = formatar_moeda(df.iloc[-1]['soma']) if 'soma' in df.columns and not df.empty else 'N/D'

    # Construção do HTML por dia
    html_dados_por_dia = ""
    for data_venc, grupo in df_relatorio.groupby('vencimento'):
        
        html_linhas_tabela = ""
        for _, linha in grupo.iterrows():
            nome = linha['nome'][:30] # Limita o nome a 30 caracteres
            doc = linha['doc']
            filial = linha['filial']
            pagar_str = formatar_moeda(linha['pagar'])
            receber_str = formatar_moeda(linha['receber'])
            soma_str = formatar_moeda(linha['soma'])
            html_linhas_tabela += f"""
                <tr>
                    <td>{nome}</td>
                    <td>{doc}</td>
                    <td>{filial}</td>
                    <td style="color: red;">{pagar_str}</td>
                    <td style="color: green;">{receber_str}</td>
                    <td>{soma_str}</td>
                </tr>
            """
        
        # Totais do dia
        total_pagar_dia = formatar_moeda(grupo['pagar'].sum())
        total_receber_dia = formatar_moeda(grupo['receber'].sum())
        ultimo_saldo_dia = formatar_moeda(grupo.iloc[-1]['soma'])

        html_dados_por_dia += f"""
            <p style="text-align: center;" >{'_'*100}</p>
            <p><strong>Data de Vencimento: {data_venc}</strong></p>
            <table class="dados-tabela">
                <thead>
                    <tr>
                        <th class="nome">Nome</th>
                        <th>Doc</th>
                        <th>Filial</th>
                        <th>Pagar</th>
                        <th>Receber</th>
                        <th>Saldo Acum.</th>
                    </tr>
                </thead>
                <tbody>
                    {html_linhas_tabela}
                    <tr class="total-row">
                        <td colspan="3" style="text-align: right;"><strong>Totais do Dia:</strong></td>
                        <td style="color: red;"><strong>{total_pagar_dia}</strong></td>
                        <td style="color: green;"><strong>{total_receber_dia}</strong></td>
                        <td><strong>{ultimo_saldo_dia}</strong></td>
                    </tr>
                </tbody>
            </table>
        """
        
    data_inicio = df_relatorio['vencimento'].min()
    data_fim = df_relatorio['vencimento'].max()

    # Template HTML final
    reporte_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h1, h4, p {{ text-align: center; }}
            .header-info p {{ text-align: left; margin: 2px 0; }}
            .header-info {{ border: 1px solid #ccc; padding: 10px; margin-bottom: 20px; }}
            .dados-tabela {{ width: 100%; border-collapse: collapse; font-size: 10px; }}
            .dados-tabela th, .dados-tabela td {{ border: 1px solid #dddddd; text-align: left; padding: 4px; }}
            .dados-tabela th {{ background-color: #f2f2f2; }}
            .dados-tabela th.nome, .dados-tabela td:nth-child(1) {{ width: 40%; }}
            .dados-tabela tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .total-row td {{ background-color: #e0e0e0; font-weight: bold; }}
        </style>
    </head>
    <body>
        <h1>Relatório de Fluxo de Caixa</h1>
        <div class="header-info">
            <p><strong>Período de Análise:</strong> {data_inicio} a {data_fim}</p>
            <p><strong>Saldo Inicial do Período:</strong> {formatar_moeda(saldo_inicial)}</p>
            <p><strong>Saldo Final do Período:</strong> {saldo_final_str}</p>
            <p style="text-align: right;"><strong>Data de Emissão:</strong> {datetime.now().strftime("%d/%m/%Y %H:%M")}</p>
        </div>
        {html_dados_por_dia}
    </body>
    </html>
    """
    return reporte_html

# Você pode adicionar a função para gerar o HTML de pagamentos aqui também,
# seguindo o mesmo padrão: receber um DataFrame e retornar uma string HTML.