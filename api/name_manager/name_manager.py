import xlwings as xw
import texel.naming as txl_nm

from functools import wraps

REF_ERROR_SUBSTRS = ['=#REF', '!#REF']


def nm_str_is_ref_err(nm_str):

    for substr in REF_ERROR_SUBSTRS:
        if substr in nm_str:
            return True

    return False


def repopulate_if_needed(func):

    @wraps(func)
    def inner_func(self, *args, **kwargs):

        if self.names_to_repopulate:
            NameManager.repopulate_nms(self)

        return func(self, *args, **kwargs)

    return inner_func


class NameManager:
    """A class that is responsible for updating any named ranges
    on tracked sheets.

    """

    def __init__(self, bk: xw.Book):
        self.bk: xw.Book = bk
        self.nm_addr_strs = []
        self.nm_nm_strs = []
        self.names_to_repopulate = True

    @repopulate_if_needed
    def get_ref_error_nms(self):
        return [nm_str for nm_str, nm_addr in zip(self.nm_nm_strs, self.nm_addr_strs)
                if nm_str_is_ref_err(nm_addr)]

    def repopulate_nms(self):
        self.nm_addr_strs = [nm.refers_to for nm in self.bk.names]
        self.nm_nm_strs = [nm.name for nm in self.bk.names]
