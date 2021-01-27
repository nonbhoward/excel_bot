from excel_bot import ExcelBot
from minimalog.minimal_log import MinimalLog
from os import getcwd
from pathlib2 import Path
ml = MinimalLog()
MIN_COLUMN, MAX_COLUMN = 1, 100
MIN_ROW, MAX_ROW = 1, 100
workbook_keywords_of_interest = ['example']
worksheet_keywords_of_interest = ['words', 'vehicles']
columns_to_search = [1, 5, 27]
search_terms_to_find = ['bacon', 'cheese', 'superbutt']


def mainloop():
    xb = ExcelBot(workbook_keywords_of_interest, worksheet_keywords_of_interest)
    while True:
        xb.set_search_area(MIN_COLUMN, MAX_COLUMN, MIN_ROW, MAX_ROW)
        xb.search_worksheets_of_interest(search_terms_to_find)
        output_file_path = get_output_file_path()
        xb.perform_write_operations(output_file_path)
        exit()


def add_xlsx_extension_to(file_name: str) -> str:
    return file_name + '.xlsx'


def get_data_directory_name() -> str:
    return 'data_src'


def get_data_path() -> Path:
    try:
        data_path = Path(get_project_path(), get_data_directory_name())
        return data_path
    except OSError as o_err:
        ml.log_event(o_err)


def get_output_file_name() -> str:
    return 'output_file'


def get_output_file_path() -> Path:
    try:
        output_file_path = Path(get_data_path(), add_xlsx_extension_to(get_output_file_name()))
        return output_file_path
    except OSError as o_err:
        ml.log_event(o_err)


def get_project_path() -> Path:
    try:
        return Path(getcwd())
    except OSError as o_err:
        ml.log_event(o_err)


mainloop()
