import xlwings as xw


def book_name_strings(nms):
    """Returns a list of the named range names in
    as workbook.

    Arguments:
        bk {xw.Book} -- Workbook to retrieve list from.

    Returns:
        list -- named range names.
    """
    return [nm.name for nm in nms]


def book_name_addrs(nms):
    """Returns a list of the named range addresses in a workbook.

    Arguments:
        nms -- Named range list

    """

    return [nm.refers_to_range.get_address(include_sheetname=True) for nm in nms]


def add_named_range(bk: xw.Book, rng: xw.Range, nm: str):
    bk.names.add(nm, '=' + rng.get_address(include_sheetname=True))


def delete_named_range(bk: xw.Book, nm: str):

    nm_rng: xw.Name = bk.names[nm]
    nm_rng.delete()


def rename_named_range(bk: xw.Book, prev_nm: str, new_nm: str):
    nm_rng: xw.Name = bk.names[prev_nm]
    nm_rng.name = new_nm
