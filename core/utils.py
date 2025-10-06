# core/utils.py

def formatar_moeda(valor):
    """
    Formata um número para o padrão de moeda brasileiro (R$ 1.234,56).
    Retorna uma string vazia se o valor não for um número válido.
    """
    try:
        valor_float = float(valor)
        # O formato :_ serve para usar o underscore como separador de milhar
        # Trocamos ele por ponto, e o ponto decimal por vírgula.
        return f"R$ {valor_float:_.2f}".replace('.', '#').replace('_', '.').replace('#', ',')
    except (ValueError, TypeError):
        # Se o valor for inválido (ex: texto ou None), retorna um espaço.
        return ' '