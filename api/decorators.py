from functools import wraps


def stop_on_empty_first_cell(func):
    """
    Decorator that is meant to be applied to a majority of formatting
    procedures.

    The function that performs the formatting must have sht be its first argument.

    If a sheet's first cell is empty, no formatting will be attempted.
    Args:
        func (func):formatting function to be decorated.

    Returns:A decorated formatting function.

    """
    @wraps(func)
    def inner_func(sht, *args, **kwargs):
        if sht.range('a1').value is None:
            return

        return func(sht, *args, **kwargs)

    return inner_func
