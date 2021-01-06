from data_src.CONSTANTS import EXCEL_EXTS
from directory_utils.directory_utils import *
from minimalog.minimal_log import MinimalLog
from openpyxl import load_workbook
from openpyxl import Workbook
ml = MinimalLog()


class ExcelTasker:
    def __init__(self, open_file=False, create_file=False, debug=False):
        if open_file:
            ml.log_event(event='init ExcelTasker', event_completed=False)
            self.path = get_path_object_at_cwd()
            self.data_dir = get_data_dir(get_project_home(self.path))
            self.files_in_data_dir = get_all_files_in(self.data_dir)
            self.all_excel_files = get_all_files_with_valid_extensions(files=self.files_in_data_dir,
                                                                       valid_extensions=EXCEL_EXTS)
            self.open_excel_files = self.open_excel_files()
            ml.log_event(event='init ExcelTasker', event_completed=True)
        if debug:
            ml.log_event('debug routine', event_completed=False)
            self.__debug()
            ml.log_event('debug routine', event_completed=True)

    def open_excel_file(self, excel_file) -> Workbook:
        ml.log_event('open excel file {}'.format(excel_file), event_completed=False)
        excel_file = str(Path(str(self.data_dir), excel_file))
        wb = load_workbook(excel_file)
        excel_data = [excel_file, wb]
        ml.log_event('open excel file {}'.format(excel_file), event_completed=True)
        yield excel_data

    def open_excel_files(self) -> dict:
        ml.log_event('open excel files {}'.format(self.all_excel_files), event_completed=False)
        excel_metadata = list()
        open_excel_files = dict()
        for excel_file in self.all_excel_files:
            excel_metadata.append(self.open_excel_file(excel_file))
            for metadata in excel_metadata:
                for data in metadata:
                    open_excel_files[data[0]] = data[1]
        ml.log_event('open excel files {}'.format(self.all_excel_files), event_completed=True)
        return open_excel_files

    def __debug(self):
        pass


if __name__ == '__main__':
    ml.log_event('execute ExcelTasker()', event_completed=False, announce=True)
    et = ExcelTasker(debug=True, open_file=True)
    ml.log_event('close ExcelTasker()', event_completed=True, announce=True)
else:
    print('importing {}'.format(__name__))
    ml.log_event('importing {}'.format(__name__))
