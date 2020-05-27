import xlwings as xw


def activate_sheet_factory(bk: xw.Book, sht_nm: str):

    sht = bk.sheets[sht_nm]
    return sht.activate


def activate_sheet_by_nm(bk: xw.Book, sht_nm: str):
    bk.sheets[sht_nm].activate()


def hide_sheets_not_in_list(bk: xw.Book, nm_list):

    cur_nm_list = [valid_sht.name for valid_sht in bk.sheets]
    visible_list = [sht_nm for sht_nm in nm_list if sht_nm in cur_nm_list]

    if not len(visible_list):
        raise KeyError('Name list to hide has no sheets in current workbook.')

    for cur_nm in cur_nm_list:
        if cur_nm in visible_list:
            bk.sheets[cur_nm].api.visible = -1
        else:
            bk.sheets[cur_nm].api.visible = 0


def unhide_all_sheets(bk: xw.Book):

    for sht in bk.sheets:
        bk.api.visible = -1


def memo_sht_visible(bk: xw.Book):
    """A function that rembers a the list of visble sheets
    and returns a function that attempts to return to the state when
    it was called.

    Arguments:
        bk {xw.Book} -- Book to remember state of.
    """

    visible_sht = [sht.name for sht in bk.sheets if sht.api.visible == -1]

    def inner_func():
        hide_sheets_not_in_list(bk, visible_sht)

    return inner_func
