import xlwings as xw


def get_names_of_sheets(bk: xw.Book):
    return [sht.name for sht in bk.sheets]


def get_sheet(bk: xw.Book, sht_nm: str, **kwargs) -> xw.Sheet:
    """Helper function that returns a sheet and adds it if it
    does not exist.

    Arguments:
        bk {xw.Book} -- book to add/get sheet from
        sht_nm {str} -- name of sheet to add/get

    Returns:
        xw.Sheet -- Sheet added.
    """

    if sht_nm in get_names_of_sheets(bk):
        return bk.sheets[sht_nm]

    return bk.sheets.add(sht_nm, **kwargs)
