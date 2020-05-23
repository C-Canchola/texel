import xlwings as xw
import texel.naming as txl_nm


class NameManager:
    """A class that is responsible for updating any named ranges
    on tracked sheets.

    Its primary responsibility is determing how to determine which ranges
    are to be named.
    """

    def __init__(self, bk: xw.Book):
        self._bk: xw.Book = bk


    