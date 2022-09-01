"""Programa que cria planilhas dentro de arquivo do Excel.

Projeto 2 do curso de Introdução ao Python UFRGS/TCE-RS

Autor: Everton da Rosa
"""

import app

def main():
    """Ponto de entrada do programa.

    Por se tratar de uma tarefa simples, fiz o seguinte:

    A entrada será feita via arquivo INI lido com configparser

    A saída será feita com logger

    Todas bibliotecas padrão do Python

    Usei openpyxl para manipular as planilhas XLS como orientado no exercício.
    """

    welcome = '''
    ==========================================================================
    Bem-vindo(a) ao programa que insere planilhas em pasta de trabalho.
    Projeto 2 do Curso de Introdução ao Python - UFRGS/TCE-RS
    by Everton da Rosa
    ==========================================================================
    '''
    print(welcome)

    # Carrega o logger
    logger = app.log()
    logger.info('Logger carregado.')

    # Carrega as configurações
    config = app.config(r'proj2.ini')
    logger.info('Configurações carregadas')

    # Carrega a pasta de trabalho
    logger.debug(f"Abrindo {config['DESTINO']['path']}...")
    wb = app.get_workbook(config['DESTINO']['path'], logger)

    # Cria as planilhas
    sheets = app.get_sheets(config['PLANILHAS']['names'])
    for sheet_name in sheets:
        app.inject_sheet(wb, sheet_name, logger)

    # Salva a pasta de trabalho
    app.save_workbook(wb, config['DESTINO']['path'], logger)

    eop = '''
    ==========================================================================
    Fim do programa!
    ==========================================================================
    '''
    print(eop)



if __name__ == '__main__':
    main()