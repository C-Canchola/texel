import xlwings as xw
import texel.naming as txl_nm
import texel.api.name_manager.range_getters as rng_getters

REF_ERROR_SUBSTRS = ['=#REF', '!#REF']


def refers_to_is_ref_err(refers_to_str):

    for substr in REF_ERROR_SUBSTRS:
        if substr in refers_to_str:
            return True

    return False


class NameManager:
    """A class that is responsible for updating any named ranges
    on tracked sheets.

    """

    def __init__(self, bk: xw.Book):
        self.bk: xw.Book = bk
        self.nm_addr_strs = []
        self.nm_nm_strs = []
        self.names_to_repopulate = True

    def convert_nr_names_to_refers_to(self, nms):
        return [self.bk.names(nm).refers_to for nm in nms]

    def get_ref_err_nms_to_delete(self, nms):
        refers_to_list = self.convert_nr_names_to_refers_to(nms)
        return [nm for nm, refers_to in zip(nms, refers_to_list) if refers_to_is_ref_err(refers_to)]

    def get_list_of_potential_names(self, sht_nm, sht_type):
        return rng_getters.get_names_and_addresses(self.bk.sheets[sht_nm], sht_type)
