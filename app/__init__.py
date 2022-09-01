"""Utilitários e entrada, saída e processamento
"""
import configparser
import logging
import openpyxl

def config(file: str):
    """Cria uma instância do configparser e retorna ela.

    :param file: str Caminho para o arquivo INI com as configurações.

    :return configparser
    """
    parser = configparser.ConfigParser()
    parser.read(file)
    return parser

def log():
    """Configura o logger e retorna.

    :return logging
    """
    # logging.NOSET mostra tudo.
    logging.basicConfig(level=logging.NOTSET)
    return logging

def get_workbook(file: str, logger: logging):
    """Retorna uma pasta de trabalho do Excel.

    :param file: str Caminho para a pasta de trabalho.
    :param logger: logging
    :return openpyxl.Workbook
    """

    try:
        wb = openpyxl.load_workbook(filename=file)
        logger.info(f'Pasta de trabalho {file} carregada.')
        return wb
    except BaseException as err:
        logger.error(f'Não foi possível carregar {file}.')
        logger.error(err)

def inject_sheet(wb: openpyxl.Workbook, sheet_name: str, logger: logging):
    try:
        wb.create_sheet(title=sheet_name)
        logger.info(f'Planilha {sheet_name} inserida.')
    except BaseException as err:
        logger.error(f'Não foi possílve inserir a planilha {sheet_name}.')
        logger.error(err)

def get_sheets(sheet_list_as_string: str):
    """Quebra uma string com nomes de planilhas separadas por vírgula e remove o espaço.

    :param sheet_list_as_string: strin String com nomes de planilhas separadas por vírcula.
    :return list
    """
    sheets = [sheet.strip() for sheet in sheet_list_as_string.split(',')]
    return sheets

def save_workbook(wb: openpyxl.Workbook, file: str, logger: logging):
    """Salva a pasta de trabalho.

    :param wb: openpyxl.Workbook
    :param file: str Caminho para salvar a pasta de trabalho.
    :param logger: logging
    """
    try:
        wb.save(file)
        logger.info(f'Pasta de trabalho salva em {file}.')
    except BaseException as err:
        logger.error(f'Falha ao salvar a pasta de trabalho em {file}')
        logger.error(err)