from functools import wraps


def stop_on_empty_first_cell(func):

    @wraps(func)
    def inner_func(sht, *args, **kwargs):
        if sht.range('a1').value is None:
            return

        return func(sht, *args, **kwargs)
    return inner_func