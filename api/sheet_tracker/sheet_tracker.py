import xlwings as xw
import pandas as pd

from functools import wraps
from texel.book import get_sheet, get_names_of_sheets
from texel.constants import Color
from texel.api.types import sheet_types


def format_after_call(func):

    @wraps(func)
    def inner_func(self, *args, **kwargs):

        ret_val = func(self, *args, **kwargs)
        SheetTracker._reformat(self)

        return ret_val

    return inner_func


def get_key_sheet_index(sht: xw.Sheet) -> str:
    """Function to return the =SHEET(cell) formula
    that will provide a reference to a sheet.
    This will provide the ability to determine if a sheet
    has been deleted or renamed when its name listed on the SheetTracker
    tab is not found anywhere in the workbook.


    Arguments:
        sht {xw.Sheet} -- Sheet added to SheetTracker

    Returns:
        str -- Formula to be embedded into sheet_index column
    """
    return '=SHEET({})'.format(sht.cells[0, 0].get_address(include_sheetname=True))


def sht_nm_must_exist(func):
    """Decorator to ensure that a sheet name is present when attempting to add.

    Raises:
        KeyError: Sheet name is not in current workbook.


    """
    @wraps(func)
    def inner_func(self, sht_nm, *args, **kwargs):
        sht_nms = [sht.name for sht in self._bk.sheets]
        if sht_nm not in sht_nms:
            raise KeyError(
                '{} is not found in this workbook. Please add it manually before attempting to track it.')

        return func(self, sht_nm, *args, **kwargs)

    return inner_func


class SheetTracker:

    SHEET_NAME = 'TXL_SHEET_TRACKER'

    KEY_SHEET_NAME = 'sheet_name'
    KEY_DESCR = 'descr'
    KEY_TYPE = 'sheet_type'
    KEY_SHEET_INDEX = 'sheet_index'

    KEYS = [KEY_SHEET_NAME, KEY_DESCR, KEY_TYPE, KEY_SHEET_INDEX]

    def __init__(self, bk: xw.Book):
        self._bk = bk
        self._sht = self._set_sheet()
        self._init_sheet()

    def _set_sheet(self):
        return get_sheet(
            self._bk, SheetTracker.SHEET_NAME, before=self._bk.sheets[0])

    @format_after_call
    def _init_sheet(self):
        if self._sht.cells[0, 0].value is None:
            self._sht.cells[0, 0].value = SheetTracker.KEYS

    def _reformat(self):
        self._sht.autofit('columns')
        self._sht.range('a1').current_region[0, ::].color = Color.INDEX

    @format_after_call
    def remove_sheet(self, sht_nm, delete=False):
        """Removes a sheet name from tracking sheet.
        Optionally, the sheet can be removed from the workbook as well.
        Very strong possibility that I remove this feature.
        Arguments:
            sht_nm {[type]} -- name of sheet to delete.

        Keyword Arguments:
            delete {bool} -- optional flag to remove the worksheet from the workbook.(default: {False})
        """

        df = self._get_info_df()
        df = df[df[SheetTracker.KEY_SHEET_NAME] != sht_nm]

        if delete:
            try:
                del_sht: xw.Sheet = self._bk.sheets[sht_nm]
                del_sht.delete()
            except:
                pass

        self._update_info_df(df)

    @sht_nm_must_exist
    def add_sheet(self, sht_nm: str, descr: str, sht_type: sheet_types.SheetType):
        """Adds a sheet to the tracker sheeet.
        If the sheet name does not exist, a new sheet will be added to the entire workbook.
        TODO
        Only existing sheets should be allowed to be tracked.
        Handling the adding of sheets to a workbook does not add anything and adds much undesired complexity.
        Arguments:
            sht_nm {str} -- name of sheet to ad.
            descr {str} -- description of sheet.
        """

        df = self._get_info_df()

        if len(df[df[SheetTracker.KEY_SHEET_NAME] == sht_nm]):
            return self._bk.sheets[sht_nm]

        sht = get_sheet(self._bk, sht_nm)

        add_df = pd.DataFrame.from_records(
            [(sht_nm, descr, sht_type.index, get_key_sheet_index(sht))], columns=SheetTracker.KEYS)
        df = pd.concat([df, add_df])

        self._update_info_df(df)

        return sht

    def _sht_nm_bool_series(self, df, sht_nm):
        return df[SheetTracker.KEY_SHEET_NAME] == sht_nm

    def _sht_nm_in_df(self, df, sht_nm):
        bool_srs = self._sht_nm_bool_series(df, sht_nm)
        return bool_srs.any()

    def rename_sheet(self, original_nm, new_nm):
        df = self._get_info_df()

        if not self._sht_nm_in_df(df, original_nm):
            raise ValueError(
                '{} is not a sheet name currently being tracked.'.format(original_nm))

        if self._sht_nm_in_df(df, new_nm):
            raise ValueError(
                '{} is already a tracked name, please choose something else.'.format(new_nm))

        bool_srs = self._sht_nm_bool_series(df, original_nm)
        df.loc[bool_srs, SheetTracker.KEY_SHEET_NAME] = new_nm
        self._update_info_df(df)

        self._bk.sheets[original_nm].name = new_nm

    @format_after_call
    def _update_info_df(self, df):
        self._sht.range('a1').current_region.value = None
        self._sht.range('a1').options(pd.DataFrame, index=False).value = df

    def _get_info_df(self) -> pd.DataFrame:
        return self._sht.range('a1').current_region.options(pd.DataFrame, index=False).value

    def activate(self):
        self._sht.activate()

    def print_sheet_info(self):
        df: pd.DataFrame = self._get_info_df()

        for i in range(len(df)):
            print(df.iloc[[i]])

    def get_sheet_name_and_type_dict(self):
        df: pd.DataFrame = self._get_info_df()
        return dict(zip(df[SheetTracker.KEY_SHEET_NAME], df[SheetTracker.KEY_TYPE]))
