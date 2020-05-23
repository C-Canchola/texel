import xlwings as xw
from texel.api import SheetTracker
from texel.api.types import sheet_types
from texel.naming import book_name_strings_with_sheet_name_filter as nm_sht_filter


class TexlBook:

    def __init__(self, bk: xw.Book):

        self.bk = bk
        self._sheet_tracker = SheetTracker(self.bk)

        self.add_sheet_to_track = self._sheet_tracker.add_sheet
        self.rename_sht = self._sheet_tracker.rename_sheet
        self.get_sheet_and_type_dict = self._sheet_tracker.get_sheet_name_and_type_dict

    def _get_sht_name_and_nr_nm_dict(self) -> dict:

        return {sht_nm: nm_sht_filter(self.bk, sht_nm) for sht_nm in self.get_sheet_and_type_dict()}
