import os
import xlwings as xw
from texel.utils.decorators import create_active_app
# %%


@create_active_app
def get_bk(bk_str, **kwargs):

    app: xw.App = xw.apps.active
    bks = app.books

    if bk_str in [bk.name for bk in bks]:
        return bks[bk_str]

    return app.books.open(bk_str, **kwargs)
