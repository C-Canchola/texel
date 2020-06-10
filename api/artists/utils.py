import xlwings as xw

from texel.constants import Color


def get_first_sht_row(sht: xw.Sheet):
    return sht.range('a1').current_region.options(ndim=1)[0, ::]


def create_sht_type_register_decorator(func_dict: dict):
    """
    Decorator factory that uses the func dict as a container for
    set of style functions for the different sheet types.
    Args:
        func_dict (): The dictionary to hold the functions.

    Returns:
        A decorator that registers the function in the original dictionary.
    """

    def inner_func(sht_type):
        return typed_sht_style_register_decorator(func_dict, sht_type)

    return inner_func


def typed_sht_style_register_decorator(func_dict: dict, sht_type):
    def deco(func):
        func_dict[sht_type] = func
        return func

    return deco


def get_column_header_color(header_cell: xw.Range) -> int:
    header_str: str = str(header_cell.value).lower()
    value_cell = header_cell.offset(1, 0)

    if header_str.startswith('used_'):
        return Color.USED_IDENTIFIER
    elif header_is_index(header_str):
        return Color.INDEX
    elif value_cell.value is None:
        return Color.INPUT
    elif str(value_cell.formula)[0] != '=':
        return Color.INPUT
    else:
        return Color.CALCULATION


def header_is_index(header_str: str) -> bool:
    index_options = ['index', 'idx', 'ix']
    for option in index_options:
        if header_str.startswith(option):
            return True

    return False
