import xlwings as xw
from functools import wraps


def create_active_app(func):
    """Decorator for any function that uses the active
    excel application.

    Opens an application in the case that there is none open.

    Arguments:
        func {function} -- Function requiring active app.
    """
    @wraps(func)
    def inner_func(*args, **kwargs):

        close_bk_one = False

        if xw.apps.active is None:
            app = xw.App()
            app.activate()
            close_bk_one = True

        val = func(*args, **kwargs)

        if close_bk_one:
            close_bk: xw.Book = app.books[0]
            close_bk.close()

        return val

    return inner_func
