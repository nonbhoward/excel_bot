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
        self.open_excel_files_to_read, self.excel_files_to_create = dict(), dict()
        self.path, self.data_dir, self.downloads_dir = get_path_at_cwd(), get_data_path(), get_os_downloads_path()
        self.files_in_data_path = get_all_files_in(self.data_dir)
        if fetch_downloads:
            pass  # TODO fetch files from downloads & move to data dir
        self.all_excel_files = filter_files_by_ext(files=self.files_in_data_path,
                                                   valid_extensions=EXCEL_EXTS)
        if read:
            ml.log_event('read mode init', event_completed=False)
            self.open_excel_files_to_read = self.open_excel_files()
            ml.log_event('read mode init', event_completed=True)
        if write:
            ml.log_event(event='write mode init', event_completed=False)
            self.excel_files_to_create = self.get_excel_filenames_to_create()
            self.file_status, self.created_excel_files = dict(), dict()
            self.excel_files_to_create = self.append_xlsx_extension_to_filenames(self.get_excel_filenames_to_create())
            self.create_excel_files_in_queue()
            ml.log_event(event='write mode init', event_completed=True)
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
        for filename in filenames:
            if '.' not in filename:
                filenames_with_ext[filename] = filename + xlsx_ext
            else:
                string_parts = filename.split('.')
                if len(string_parts) > 2:
                    ml.log_event('potential error with filename {}'.format(filename), level=ml.WARN)
                filenames_with_ext[filename] = string_parts[0] + xlsx_ext
        return filenames_with_ext

    def create_excel_files_in_queue(self):
        for k, file_val in self.excel_files_to_create.items():
            full_file_path = build_full_path_to_filename(self.data_dir, file_val)
            self.file_status[full_file_path] = dict()
            if exists(full_file_path):
                ml.log_event('warning, file {} already exists'.format(full_file_path))
                self.file_status[full_file_path]['newly_created'] = False
                self.file_status[full_file_path]['exists'] = True
                continue
            self.instantiate_handle_and_create_excel_file_at(full_file_path)
            self.wait_one_tenth_of_a_second()
            if exists(full_file_path):
                self.file_status[full_file_path]['newly_created'] = True
                self.file_status[full_file_path]['exists'] = True
            else:
                self.file_status[full_file_path]['newly_created'] = False
                self.file_status[full_file_path]['exists'] = False

    def create_worksheet_name_in_workbook(self, ws_name: str, workbook: Workbook) -> bool:
        pass

    @staticmethod
    def get_excel_filenames_to_create() -> list:
        # TODO debug function, will be replaced with something more substantial after testing
        return ['z1', 'z2', 'z3', 'z4', 'z5', 'z6', 'z7', 'z8', 'z9']

    def instantiate_handle_and_create_excel_file_at(self, path: Path) -> bool:
        """
        TODO : note that each file takes > 1s to create, asyncio opportunity here?
        :param path: path where we will save the instantiated file
        :param filename: filename that will be used to refer to the file
        :return: boolean, success or failure
        """
        new_wb = Workbook()
        new_wb.save(path)
        full_path = str(path.resolve())
        self.created_excel_files[full_path] = dict()
        self.created_excel_files[full_path]['workbook'] = new_wb
        self.wait_one_tenth_of_a_second()  # give the disk time to create the file, is this delay necessary or not?
        if exists(path):
            return True
        return False

    def get_all_worksheets_from_workbook(self) -> dict():
        pass

    def get_value_at_cell(self) -> str:
        pass

    def get_worksheet_from_workbook(self, worksheet_name: str) -> worksheet:
        pass

    def move_excel_files_from_downloads_to_data_dir(self) -> bool:
        # TODO untested
        pass
        files_in_downloads = get_all_files_in(self.downloads_dir)
        excel_files_in_downloads = filter_files_by_ext(files=files_in_downloads,
                                                       valid_extensions=EXCEL_EXTS)
        if not len(excel_files_in_downloads) < 1:
            if move_files(files=excel_files_in_downloads,
                          src_path=self.downloads_dir,
                          dest_path=self.data_dir) is SUCCESSFUL:
                return True
            return False
        return True

    def open_excel_file(self, excel_file) -> Workbook:
        ml.log_event('open excel file {}'.format(excel_file), event_completed=False)
        excel_file = str(Path(str(self.data_dir), excel_file))
        wb = load_workbook(excel_file)
        workbook_data = [excel_file, wb]
        ml.log_event('open excel file {}'.format(excel_file), event_completed=True)
        yield workbook_data

    def open_excel_files(self) -> dict:
        ml.log_event('open excel files {}'.format(self.all_excel_files), event_completed=False)
        excel_metadata = list()
        open_excel_files = dict()
        for excel_file in self.all_excel_files:
            excel_metadata.append(self.open_excel_file(excel_file))
            for metadata in excel_metadata:
                for data in metadata:
                    open_excel_files[data[0]] = dict()
                    open_excel_files[data[0]]['workbook'] = data[1]
        ml.log_event('open excel files {}'.format(self.all_excel_files), event_completed=True)
        return open_excel_files

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
    ml.log_event('execute ExcelTasker read test', event_completed=False, announce=True)
    et_read = ExcelTasker(debug=True, read=True)
    ml.log_event('close ExcelTasker read test', event_completed=True, announce=True)
    # ml.log_event('execute ExcelTasker write test', event_completed=False, announce=True)
    # et_write = ExcelTasker(debug=True, write=True)
    # ml.log_event('execute ExcelTasker write test', event_completed=False, announce=False)
    pass
else:
    print('importing {}'.format(__name__))
    ml.log_event('importing {}'.format(__name__))
