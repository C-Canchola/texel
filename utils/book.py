import os
import xlwings as xw
from ..utils.decorators import create_active_app
# %%


@create_active_app
def get_bk(bk_str, **kwargs):
    """Util function to open an excel file while also handling the application
    instance of excel.

    Will open an application if one is currently not active.


    Arguments:
        bk_str {str or path like} -- name or path to excel file. name is used when the file is already open in an excel app.

    Returns:
        xw.Book -- xlwings book instance.
    """
    app: xw.App = xw.apps.active
    bks = app.books

    if bk_str in [bk.name for bk in bks]:
        return bks[bk_str]

    if bk_str in [bk.fullname for bk in bks]:
        return bks[os.path.split(bk_str)[-1]]

    return app.books.open(bk_str, **kwargs)
