from minimalog.minimal_log import MinimalLog
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from os import getcwd
from os import walk
from os.path import exists
from pathlib2 import Path
EXCEL_FORMATS = ['xlsx']  # fyi openpyxl supports four file types but only one is used here
ml = MinimalLog(__name__)


class ExcelBot:
    def __init__(self, workbook_koi, worksheet_koi):
        ml.log_event('initializing {}'.format(self.__class__.__name__), event_completed=False, announce=True)
        # initialize variables
        self.search_results, self.workbooks, self.worksheets, self.worksheet_data, self.worksheet_data_of_interest = \
            dict(), dict(), dict(), dict(), dict()
        self.workbook_koi, self.worksheet_koi = workbook_koi, worksheet_koi
        self.min_col, self.max_col, self.min_row, self.max_row = self._get_default_range()
        # setup data paths and fetch file data
        self.data_path = self._get_data_path()
        self.excel_files = self._get_excel_files_from_data_path()
        self._read_worksheets_into_class()
        self._extract_worksheet_data_into_class()
        self._extract_worksheet_data_of_interest()
        ml.log_event('initializing {}'.format(self.__class__.__name__), event_completed=True, announce=True)

    @staticmethod
    def perform_write_operations(output_file_name):
        ml.log_event('performing write operations on file {}'.format(output_file_name), event_completed=False)
        try:
            if not exists(output_file_name):
                new_workbook = Workbook()
                new_workbook.save(output_file_name)
                ml.log_event('performing write operations on file {}'.format(output_file_name), event_completed=True)
                return
            ml.log_event('warning, file already exists')
            ml.log_event('performing write operations on file {}'.format(output_file_name), event_completed=True)
        except OSError as o_err:
            ml.log_event(o_err)

    def search_worksheets_of_interest(self, search_terms):
        try:
            ml.log_event('searching worksheets of interest', event_completed=False)
            self._record_cells_which_contain(search_terms)
            ml.log_event('searching worksheets of interest', event_completed=True)
        except KeyError as k_err:
            ml.log_event(k_err)

    def set_keywords_of_interest(self, workbook_keywords, worksheet_keywords):
        try:
            ml.log_event('set workbook_keywords: {} and worksheet keywords: {}'.format(
                workbook_keywords, worksheet_keywords))
            self.workbook_koi, self.worksheet_koi = workbook_keywords, worksheet_keywords
        except ValueError as v_err:
            ml.log_event(v_err)

    def set_search_area(self, min_col, max_col, min_row, max_row):
        ml.log_event('search area set..\n min_col: {}\n max_col: {}\n min_row: {}\n max_row: {}'.format(
            min_col, max_col, min_row, max_row))
        self.min_col = min_col
        self.max_col = max_col
        self.min_row = min_row
        self.max_row = max_row

    def _extract_data_from_worksheet(self, workbook_path, a_worksheet):
        ml.log_event('extracting data from worksheet: {} at workbook_path: {}'.format(
            a_worksheet, workbook_path), event_completed=False)
        try:
            self.worksheet_data[a_worksheet.title] = dict()
            self.worksheet_data[a_worksheet.title]['parent_workbook'] = str(workbook_path)
            for column in range(1, self.max_col + 1):
                for row in range(1, self.max_row + 1):
                    cell = get_column_letter(column) + str(row)
                    value = a_worksheet[cell].value
                    if value is not None:
                        ml.log_event('value: {} found at cell: {}'.format(value, cell))
                        self.worksheet_data[a_worksheet.title][cell] = value
            ml.log_event('extracting data from worksheet: {} at workbook_path: {}'.format(
                a_worksheet, workbook_path), event_completed=True)
        except KeyError as k_err:
            ml.log_event(k_err)

    def _extract_worksheet_data_into_class(self):
        ml.log_event('extract worksheet data into class', event_completed=False)
        try:
            read_data_success = list()
            for workbook_path, worksheets_dict in self.worksheets.items():
                for worksheet_title in worksheets_dict.keys():
                    a_worksheet = worksheets_dict[worksheet_title]
                    if self._extract_data_from_worksheet(workbook_path, a_worksheet):
                        pass
            if all(read_data_success):
                ml.log_event('extract worksheet data into class', event_completed=True)
                return True
            ml.log_event('extract worksheet data into class, failure encountered', event_completed=True)
            return False
        except OSError as o_err:
            ml.log_event(o_err)

    @staticmethod
    def _get_data_directory_name() -> str:
        ml.log_event('get data directory name')
        return 'data_src'

    def _get_data_path(self) -> Path:
        ml.log_event('get data path', event_completed=False)
        try:
            data_path = Path(self._get_project_directory(), self._get_data_directory_name())
            ml.log_event('get data path', event_completed=True)
            return data_path
        except RuntimeError as r_err:
            ml.log_event(r_err)

    @staticmethod
    def _get_default_range() -> tuple:
        ml.log_event('get default range')
        return 1, 100, 1, 100

    def _get_data_path_files(self) -> list:
        ml.log_event('get files from data path', event_completed=False)
        try:
            all_files = list()
            for root, dirs, files in walk(self.data_path):
                for file in files:
                    all_files.append(Path(root, file))
            ml.log_event('get files from data path', event_completed=True)
            return all_files
        except FileNotFoundError as f_err:
            ml.log_event(f_err)

    def _get_excel_files_from_data_path(self) -> list:
        ml.log_event('get excel files from data path', event_completed=False)
        try:
            excel_files = list()
            data_path_files = self._get_data_path_files()
            for file in data_path_files:
                if _is_excel_file(file):
                    excel_files.append(file)
            ml.log_event('get excel files from data path', event_completed=True)
            return excel_files
        except FileNotFoundError as f_err:
            ml.log_event(f_err)

    @staticmethod
    def _get_project_directory() -> Path:
        ml.log_event('get project directory')
        try:
            return Path(getcwd())
        except RuntimeError as r_err:
            ml.log_event(r_err)

    def _extract_worksheet_data_of_interest(self):
        ml.log_event('extract worksheet data of interest', event_completed=False)
        try:
            worksheet_data_of_interest = dict()
            for worksheet_title, worksheet_data in self.worksheet_data.items():
                for workbook_keywords in self.workbook_koi:
                    if workbook_keywords in worksheet_data['parent_workbook']:
                        for worksheet_keyword in self.worksheet_koi:
                            if worksheet_keyword in worksheet_title:
                                ml.log_event('workbook of interest found: {} worksheet of interest found: {}'.format(
                                    worksheet_data['parent_workbook'], worksheet_title))
                                worksheet_data_of_interest[worksheet_title] = worksheet_data
            self.worksheet_data_of_interest = worksheet_data_of_interest
            ml.log_event('extract worksheet data of interest', event_completed=True)
        except RuntimeError as r_err:
            ml.log_event(r_err)

    def _read_worksheets_into_class(self):
        ml.log_event('read worksheets into class', event_completed=False)
        try:
            excel_file_read_success = list()
            for excel_file in self.excel_files:
                if self._worksheets_saved_to_class(excel_file):
                    ml.log_event('worksheets from {} saved to class'.format(excel_file))
                    excel_file_read_success.append(True)
            if all(excel_file_read_success):
                ml.log_event('read worksheets into class', event_completed=True)
                return True
            ml.log_event('read worksheets into class, a failure was encountered', event_completed=True)
            return False
            pass
        except RuntimeError as r_err:
            ml.log_event(r_err)

    def _record_cells_which_contain(self, search_terms):
        ml.log_event('record cells that contain search_terms: {}'.format(search_terms), event_completed=False)
        try:
            for worksheet_title, worksheet_data in self.worksheet_data.items():
                self.search_results[worksheet_title] = dict()
                self.search_results[worksheet_title]['parent_workbook'] = worksheet_data['parent_workbook']
                for search_term in search_terms:
                    self.search_results[worksheet_title][search_term] = list()
                    for cell_address, cell_value in worksheet_data.items():
                        if search_term in str(cell_value):
                            self.search_results[worksheet_title][search_term].append(cell_address)
                            ml.log_event('{} found in {} at {}'.format(cell_value, worksheet_title, cell_address))
            ml.log_event('record cells that contain search_terms: {}'.format(search_term), event_completed=True)
        except RuntimeError as r_err:
            ml.log_event(r_err)

    def _worksheets_saved_to_class(self, excel_file):
        ml.log_event('save worksheets from {} to class'.format(excel_file), event_completed=False)
        try:
            self.workbooks[excel_file] = load_workbook(excel_file)
            self.worksheets[excel_file] = dict()
            for misc_worksheet in self.workbooks[excel_file]:
                worksheet_title = misc_worksheet.title
                self.worksheets[excel_file][worksheet_title] = misc_worksheet
            worksheet_count = len(self.worksheets[excel_file])
            if worksheet_count > 0:
                ml.log_event('save worksheets from {} to class'.format(excel_file), event_completed=True)
                return True
            ml.log_event('failure to save worksheets from {} to class'.format(excel_file))
            return False
        except KeyError as k_err:
            ml.log_event(k_err)


def _is_excel_file(file_name: Path) -> bool:
    ml.log_event('check if {} is excel file'.format(file_name), event_completed=False)
    try:
        name_parts = str(file_name).split('.')
        if name_parts[-1] in EXCEL_FORMATS:
            ml.log_event('check if {} is excel file'.format(file_name), event_completed=True)
            return True
        ml.log_event('{} failed check'.format(file_name))
        return False
    except IndexError as i_err:
        ml.log_event(i_err)
