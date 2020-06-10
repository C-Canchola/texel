import xlwings as xw

from more_itertools import flatten

from ..sheet_tracker import SheetTracker
from ..name_manager import NameManager
from ..types import sheet_types
from ..artists import ColorArtist
from ..tasks import task_manager
from ..xl_fmt import formula_formatter

from ... import naming as txl_nm

nm_sht_filter = txl_nm.book_name_strings_with_sheet_name_filter
all_bk_nm_strs = txl_nm.book_name_strings
add_nm_by_addr = txl_nm.add_named_range_from_addr

_all_bk_nm_strs = all_bk_nm_strs
_task_manager = task_manager


class TexlBook:
    SHEET_TYPE = sheet_types

    def __init__(self, bk: xw.Book):

        self.bk = bk
        self._sheet_tracker = SheetTracker(self.bk)
        self._name_manager = NameManager(self.bk)
        self._get_sheet_and_type_dict = self._sheet_tracker.get_sheet_name_and_type_dict
        self._tracker_active_func = task_manager.activate_sheet_factory(
            self.bk, self._sheet_tracker.SHEET_NAME)

        self._initial_sheet_visible_func = task_manager.memo_sht_visible(
            self.bk)

    def add_two_dimensional_look_up_sheet_to_track(self, sht_nm, sht_descr: str, row_label: str,
                                                   col_label: str, data_label: str):
        """
        Convenience method to add two dimensional look up as its creation requires more initial
        information to correctly handle the naming of ranges.
        Args:
            sht_nm (str): name of sheet to add. Must Exist in the workbook.
            sht_descr (str): description of the sheet and its purpose.
            row_label (str): label of row in two dim sheet.
            col_label (str): label of col in two dim sheet.
            data_label (str): label for the data in two dim sheet.

        Returns:

        """
        self.add_sheet_to_track(sht_nm, sht_descr, sheet_types.TWO_DIMENSIONAL_LOOK_UP, row_label=row_label,
                                col_label=col_label, data_label=data_label)

    def add_sheet_to_track(self, sht_nm: str, sht_descr: str, sht_type: SHEET_TYPE, **kwargs):
        """Adds a sheet to track my name.

        This sheet will be tracked by what is essentially an excel pointer.
        A formula will directly reference the sheet so that it will always stay in sync
        with any user updates or deletions.


        Arguments:
            sht_nm {str} -- name of sheet to add.
            sht_descr {str} -- description of sheet to display on the sheet tracker.
            sht_type {TexlBook.SHEET_TYPE} -- Sheet type. Will determine formatting and naming rules.
        """
        self._sheet_tracker.add_sheet(sht_nm, sht_descr, sht_type, **kwargs)

    def _get_sht_name_and_nr_nm_dict(self) -> dict:

        return {sht_nm: nm_sht_filter(self.bk, sht_nm) for sht_nm in self._get_sheet_and_type_dict()}

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

    def _get_sht_potential_nms(self) -> dict:
        return {sht_nm: self._name_manager.get_list_of_potential_names(sht_nm, int(sht_type))
                for sht_nm, sht_type in self._get_sheet_and_type_dict().items()}

    def _get_all_potential_nms(self):
        return list(flatten(self._get_sht_potential_nms().values()))

    def full_update(self):
        """Full update for all the currently tracked sheets.
        Will update the named ranges as well as apply proper coloring.

        """
        self.update_all_names()
        self.color_all_sheets()

    def format_sheet_formulas(self):
        """
        Formats the formulas of the registered sheets based upon their sheet type.

        Returns:None

        """
        for sht_nm, sht_type in self._get_sheet_and_type_dict().items():
            formula_formatter.format_typed_sheet_formulas(self.bk.sheets[sht_nm], sht_type)

    def color_all_sheets(self):
        """Applies coloring to tracked sheets.
        Coloring will be applied to sheets based off of their sheet type as well
        as the types of data they contain.
        See color_artist module for more details.
        """
        for sht_nm, sht_type in self._get_sheet_and_type_dict().items():
            ColorArtist.color_typed_sheet(self.bk.sheets[sht_nm], sht_type)

    def update_all_names(self):
        """Attempts to rename all the tracked sheet ranges
        based off the rules governed by their sheet type.

        As of right now there is no validation or error handling as well
        as no formal documentation of what is valid or considered good practice.

        This is the most important aspect that will need to be worked on.
        """
        self._name_manager.delete_all_ref_error_nm_rngs()

        potential_nm_list = self._get_all_potential_nms()
        all_nms = all_bk_nm_strs(self.bk)
        all_tracked_nms = self._get_all_tracked_nr_nms()

        self._name_manager.handle_all_naming_cases(
            potential_nm_list, all_nms, all_tracked_nms)

    def activate_tracking_sheet(self):
        """Command to access the sheet tracker tab.
        """
        self._tracker_active_func()

    def set_initial_sheet_visibility(self):
        """Attempts to return to the initial state of visibility at
        registration time.
        """
        self._initial_sheet_visible_func()

    def _get_list_of_tracked_sheet_names(self):
        return list(self._get_sheet_and_type_dict().keys())

    def show_only_tracked_sheets(self):
        """Alters the visibility to only tracked sheets as well as the
        sheet tracker tab FOR NOW.
        """
        _task_manager.hide_sheets_not_in_list(
            self.bk, self._get_list_of_tracked_sheet_names())
        self.bk.sheets[self._sheet_tracker.SHEET_NAME].api.visible = -1

    def unhide_all_sheets(self):
        """Unhides every sheet in the workbook.
        """
        _task_manager.unhide_all_sheets(self.bk)

    def active_sheet_by_name(self, sht_nm):
        """Activates to sheet by name.
        If the sheet does not exist, an error will occur.

        If the sheet exists but is not visible, it will be unhidden and activated.

        Arguments:
            sht_nm {[type]} -- name of sheet to activate/navigate to.
        """
        possible_nms = [sht.name for sht in self.bk.sheets]
        if sht_nm not in possible_nms:
            raise KeyError("{} does not exist.".format(sht_nm))

        self.bk.sheets[sht_nm].api.visible = -1
        self.bk.sheets[sht_nm].activate()

    def log_sheet_info(self):
        self._sheet_tracker.log_sheet_status()
