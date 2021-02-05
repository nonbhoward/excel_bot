from excel_bot import ExcelBot
from minimalog.minimal_log import MinimalLog
from os import getcwd
from pathlib2 import Path
ml = MinimalLog()


def mainloop():
    """
    the abstracted main operations of the program that are performed during program execution
    :return: None
    """
    search_terms_to_find = get_search_terms_to_find()
    workbook_keywords_of_interest = get_workbook_keywords_of_interest()
    worksheet_keywords_of_interest = get_worksheet_keywords_of_interest()
    xb = ExcelBot(workbook_keywords_of_interest, worksheet_keywords_of_interest)
    while True:
        search_area_coordinates = get_search_area()
        xb.set_search_area(search_area_coordinates)
        xb.search_worksheets_of_interest_and_record_cells_containing_(search_terms_to_find)
        output_file_path = get_output_file_path()
        xb.write_file_to_disk(output_file_path)
        exit()


def add_xlsx_extension_to(filename: str) -> str:
    """
    :param filename: adds excel extension onto empty filename
    :return:
    """
    return filename + '.xlsx'


def get_data_directory_name() -> str:
    """
    :return: the name of the data directory chosen by the developer
    """
    return 'data_src'


def get_data_path() -> Path:
    """
    :return: the path to the data directory
    """
    try:
        data_path = Path(get_project_path(), get_data_directory_name())
        return data_path
    except OSError as o_err:
        ml.log_event(o_err)


def get_output_file_name() -> str:
    """
    :return: the name of the output excel file chosen by the developer
    """
    return 'output_file'


def get_output_file_path() -> Path:
    """
    :return: the path directly to the output file
    """
    try:
        output_file_path = Path(get_data_path(), add_xlsx_extension_to(get_output_file_name()))
        return output_file_path
    except OSError as o_err:
        ml.log_event(o_err)


def get_project_path() -> Path:
    """
    :return: the project's working directory, from which all other project paths branch
    """
    try:
        return Path(getcwd())
    except OSError as o_err:
        ml.log_event(o_err)


def get_search_area() -> tuple:
    """
    # FIXME customize search area
    :return:
    """
    min_column, max_column = 1, 100
    min_row, max_row = 1, 100
    search_area_coordinates = min_column, max_column, min_row, max_row
    return search_area_coordinates


def get_search_terms_to_find() -> list:
    """
    # FIXME customize search terms
    add your custom search terms here
    :return:
    """
    search_terms = [
        'first search term',
        'second search term',
        'third search term',
        'bacon',
        'cheese',
        'super'
    ]
    return search_terms


def get_workbook_keywords_of_interest() -> list:
    """
    # FIXME customize workbook koi
    add your custom workbook keywords here
    :return:
    """
    workbook_keywords_of_interest = [
        'first workbook koi',
        'second workbook koi',
        'example'
    ]
    return workbook_keywords_of_interest


def get_worksheet_keywords_of_interest() -> list:
    """
    # FIXME customize worksheet koi
    add your custom worksheet keywords here
    :return:
    """
    worksheet_keywords_of_interest = [
        'first worksheet koi',
        'second worksheet koi',
        'words',
        'vehicles'
    ]
    return worksheet_keywords_of_interest


mainloop()
