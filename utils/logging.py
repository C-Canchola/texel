import xlwings as xw


def log_sheets_with_index(bk: xw.Book):
    """Logs the indicies and name of every sheet in a workbook.

    Arguments:
        bk {xw.Book} -- The book to log.
    """

    for index, sht in enumerate(bk.sheets):
        print(index, sht.name)
