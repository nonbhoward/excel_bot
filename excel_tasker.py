from data_src.CONSTANTS import EXCEL_EXTS
from directory_utils.directory_utils import *
from minimalog.minimal_log import MinimalLog
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl import worksheet
from time import sleep
ml = MinimalLog()
SUCCESSFUL = True


class ExcelTasker:
    def __init__(self, read=False, write=False, fetch_downloads=False, debug=False):
        """
        :param read: bool, the attribute 'files_to_read' will be populated
        :param write: bool, the attribute 'files_to_write' will be populated
        :param fetch_downloads: bool, an attempt to copy files from downloads will occur
        :param debug: bool, debug function will run
        """
        ml.log_event(event='init ExcelTasker with read = {} and write = {}'.format(read, write),
                     event_completed=False)
        self.open_workbooks, self.workbooks_to_create = dict(), dict()
        self.path, self.data_dir, self.downloads_dir = get_path_at_cwd(), get_data_path(), get_os_downloads_path()
        if write:
            ml.log_event(event='write mode init', event_completed=False)
            self.workbook_status, self.created_workbooks = dict(), dict()
            self.workbooks_to_create = self.append_xlsx_extension_to_filenames(self.get_filenames_to_create())
            self.create_workbooks_in_queue()
            ml.log_event(event='write mode init', event_completed=True)
        self.files_in_data_path = get_all_files_in(self.data_dir)
        if fetch_downloads:
            pass  # TODO fetch files from downloads & move to data dir
        self.workbook_names = filter_files_by_ext(files=self.files_in_data_path, valid_extensions=EXCEL_EXTS)
        if read:
            ml.log_event('read mode init', event_completed=False)
            self.open_workbooks = self.open_excel_files_in_data_dir()
            ml.log_event('read mode init', event_completed=True)
        ml.log_event(event='init ExcelTasker', event_completed=True)
        if debug:
            ml.log_event('debug routine', event_completed=False)
            self.__debug()
            ml.log_event('debug routine', event_completed=True)

    @staticmethod
    def append_xlsx_extension_to_filenames(filenames: list) -> dict:
        """
        splits filename string on '.' delimiter, ideally finds zero or two, appends '.xlsx' to first
        :return: filename.xlsx
        """
        xlsx_ext = '.xlsx'
        filenames_with_ext = dict()
        try:
            for filename in filenames:
                if '.' not in filename:
                    filenames_with_ext[filename] = filename + xlsx_ext
                else:
                    string_parts = filename.split('.')
                    if len(string_parts) > 2:
                        ml.log_event('potential error with filename {}'.format(filename), level=ml.WARN)
                    filenames_with_ext[filename] = string_parts[0] + xlsx_ext
            return filenames_with_ext
        except OSError as o_err:
            ml.log_exception(o_err)

    def create_workbooks_in_queue(self):
        try:
            for k, file_val in self.workbooks_to_create.items():
                full_file_path = build_full_path_to_filename(self.data_dir, file_val)
                self.workbook_status[full_file_path] = dict()
                if exists(full_file_path):
                    ml.log_event('warning, file {} already exists'.format(full_file_path))
                    self.workbook_status[full_file_path]['newly_created'] = False
                    self.workbook_status[full_file_path]['exists'] = True
                    continue
                self.instantiate_and_create_workbook_at(full_file_path)
                self.wait_one_tenth_of_a_second()
                if exists(full_file_path):
                    self.workbook_status[full_file_path]['newly_created'] = True
                    self.workbook_status[full_file_path]['exists'] = True
                else:
                    self.workbook_status[full_file_path]['newly_created'] = False
                    self.workbook_status[full_file_path]['exists'] = False
        except KeyError as k_err:
            ml.log_exception(k_err)

    def create_worksheet_name_in_workbook(self, ws_name: str, workbook: Workbook) -> bool:
        pass

    @staticmethod
    def get_filenames_to_create() -> list:
        # TODO debug function, will be replaced with something more substantial after testing
        return ['z1', 'z2', 'z3', 'z4', 'z5', 'z6', 'z7', 'z8', 'z9']

    def instantiate_and_create_workbook_at(self, path: Path) -> bool:
        """
        TODO : note that each file takes > 1s to create, asyncio opportunity here?
        :param path: path where we will save the instantiated file
        :param filename: filename that will be used to refer to the file
        :return: boolean, success or failure
        """
        try:
            new_wb = Workbook()
            new_wb.save(path)
            full_path = str(path.resolve())
            self.created_workbooks[full_path] = dict()
            self.created_workbooks[full_path]['workbook'] = new_wb
            self.wait_one_tenth_of_a_second()  # give the disk time to create the file, is this delay necessary or not?
        except OSError as o_err:
            ml.log_exception(o_err)
        if exists(path):
            return True
        return False

    def get_all_worksheets_from_workbook(self) -> dict():
        for workbook in self.workbook_names:
            pass

    def get_value_at_cell(self) -> str:
        pass

    def get_worksheet_from_workbook(self, worksheet_name: str) -> worksheet:
        pass

    def open_excel_files_in_data_dir(self) -> dict:
        ml.log_event('open excel files {}'.format(self.workbook_names), event_completed=False)
        workbook_metadata, open_workbooks = list(), dict()
        for excel_file in self.workbook_names:
            workbook_metadata.append(self.open_excel_workbook(excel_file))
            for metadata in workbook_metadata:
                for data in metadata:
                    open_workbooks[data[0]] = dict()
                    open_workbooks[data[0]]['workbook'] = data[1]
        ml.log_event('open excel files {}'.format(self.workbook_names), event_completed=True)
        return open_workbooks

    def open_excel_workbook(self, excel_file) -> Workbook:
        try:
            ml.log_event('open excel file {}'.format(excel_file), event_completed=False)
            excel_file = str(Path(str(self.data_dir), excel_file))
            wb = load_workbook(excel_file)
            workbook_data = [excel_file, wb]
            ml.log_event('open excel file {}'.format(excel_file), event_completed=True)
            yield workbook_data
        except OSError as o_err:
            ml.log_exception(o_err)

    @staticmethod
    def wait_one_second():
        sleep(1)

    @staticmethod
    def wait_one_tenth_of_a_second():
        sleep(0.1)

    def write_value_at_cell(self) -> bool:
        pass

    def __debug(self):
        pass


if __name__ == '__main__':
    ml.log_event('execute ExcelTask', event_completed=False, announce=True)
    et_read = ExcelTasker(debug=True, read=True, write=True)
    ml.log_event('close ExcelTask', event_completed=True, announce=True)
    pass
else:
    print('importing {}'.format(__name__))
    ml.log_event('importing {}'.format(__name__))
