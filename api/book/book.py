import xlwings as xw
import texel.naming as txl_nm

from texel.api.sheet_tracker import SheetTracker
from texel.api.name_manager import NameManager
from texel.api.types import sheet_types
from texel.api.artist import ColorArtist
from more_itertools import flatten

nm_sht_filter = txl_nm.book_name_strings_with_sheet_name_filter
all_bk_nm_strs = txl_nm.book_name_strings
add_nm_by_addr = txl_nm.add_named_range_from_addr


class TexlBook:

    SHEET_TYPE = sheet_types

    def __init__(self, bk: xw.Book):

        self.bk = bk
        self._sheet_tracker = SheetTracker(self.bk)
        self._name_manager = NameManager(self.bk)
        self.add_sheet_to_track = self._sheet_tracker.add_sheet
        self.rename_sht = self._sheet_tracker.rename_sheet
        self.get_sheet_and_type_dict = self._sheet_tracker.get_sheet_name_and_type_dict

    def _get_sht_name_and_nr_nm_dict(self) -> dict:

        return {sht_nm: nm_sht_filter(self.bk, sht_nm) for sht_nm in self.get_sheet_and_type_dict()}

    def _get_all_tracked_nr_nms(self) -> list:
        return list(flatten(self._get_sht_name_and_nr_nm_dict().values()))

    def _get_track_ref_error_nms(self):
        return self._name_manager.get_ref_err_nms_to_delete(
            self._get_all_tracked_nr_nms())

    def remove_sht_from_tracking(self, sht_nm):
        """Removes a sheet from the tracker sheet.

        Arguments:
            sht_nm {str} -- the name of the sheet to be removed.
        """

        self._sheet_tracker.remove_sheet(sht_nm)

    def get_sht_potential_nms(self) -> dict:
        return {sht_nm: self._name_manager.get_list_of_potential_names(sht_nm, int(sht_type))
                for sht_nm, sht_type in self.get_sheet_and_type_dict().items()}

    def get_all_potential_nms(self):
        return list(flatten(self.get_sht_potential_nms().values()))

    def full_update(self):
        """Full update for all the currently tracked sheets.
        Will update the named ranges as well as apply proper coloring.

        """
        self.update_all_names()
        self.color_all_sheets()

    def color_all_sheets(self):

        for sht_nm, sht_type in self.get_sheet_and_type_dict().items():
            ColorArtist.color_typed_sheet(self.bk.sheets[sht_nm], sht_type)

    def update_all_names(self):

        self._name_manager.delete_all_ref_error_nm_rngs()

        potential_nm_list = self.get_all_potential_nms()
        all_nms = all_bk_nm_strs(self.bk)
        all_tracked_nms = self._get_all_tracked_nr_nms()

        self._name_manager.handle_all_naming_cases(
            potential_nm_list, all_nms, all_tracked_nms)
