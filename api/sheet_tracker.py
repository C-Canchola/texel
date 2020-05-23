import xlwings as xw

from texel.book import get_sheet


class SheetTracker:

    SHEET_NAME = 'TEXEL_SHEET_TRACKER'

    def __init__(self, bk: xw.Book):
        self._bk = bk
        self._set_sheet()

    def _set_sheet(self):
        self._sht = get_sheet(
            self._bk, SheetTracker.SHEET_NAME, before=self._bk.sheets[0])
