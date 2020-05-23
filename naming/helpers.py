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
    """Adds a range to a book with a given name.
    Prefer to use add_named_range_from_addr

    Arguments:
        bk {xw.Book} -- Book to add named range
        rng {xw.Range} -- Range to add
        nm {str} -- name of range
    """
    add_named_range_from_addr(bk, rng.get_address(include_sheetname=True), nm)


def add_named_range_from_addr(bk: xw.Book, addr: str, nm: str):
    """Adds a range to a book with a given name by the range address
    PREFFERRED METHOD OF ADDING A RANGE.

    Arguments:
        bk {xw.Book} -- book to add named range
        addr {str} -- addr of range
        nm {str} -- name of range
    """
    bk.names.add(nm, '=' + addr)


def delete_named_range(bk: xw.Book, nm: str):

    nm_rng: xw.Name = bk.names[nm]
    nm_rng.delete()


def rename_named_range(bk: xw.Book, prev_nm: str, new_nm: str):
    nm_rng: xw.Name = bk.names[prev_nm]
    nm_rng.name = new_nm
