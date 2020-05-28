import xlwings as xw


def log_sheets_with_index(bk: xw.Book):
    """Logs the indicies and name of every sheet in a workbook.

    Arguments:
        bk {xw.Book} -- The book to log.
    """

    for index, sht in enumerate(bk.sheets):
        print(index, sht.name)


def log_active_books():
    """Looks the open books in the active workbook along with their index.
    Meant to aid in quickly setting the correct book so the name does not 
    have to be typed out.
    """

    for index, bk in enumerate(xw.apps.active.books):
        print(index, bk)
