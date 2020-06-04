import xlwings as xw

from functools import wraps

from ..decorators import stop_on_empty_first_cell as _stop_on_empty_first_cell

from ..types import sheet_types as txl_sht_types

from ...constants import Color
from ...formatting import sheet_fmt


def _get_first_sht_row(sht: xw.Sheet):
    return sht.range('a1').current_region.options(ndim=1)[0, ::]


_SHEET_COLOR_FUNCS = {}


def register_sheet_color_func(sht_type, tab_color):
    def decorator(func):
        @wraps(func)
        def inner_func(sht, *args, **kwargs):
            sheet_fmt.color_sht_tab(sht, tab_color)

            return func(sht, *args, **kwargs)

        _SHEET_COLOR_FUNCS[sht_type] = inner_func
        return inner_func

    return decorator


@register_sheet_color_func(txl_sht_types.SCALAR_INPUT, Color.INPUT)
@_stop_on_empty_first_cell
def color_input_sht(sht: xw.Sheet):
    input_row = _get_first_sht_row(sht)
    input_row.color = Color.INPUT

    input_row[0, 0].color = Color.INDEX


@register_sheet_color_func(txl_sht_types.STANDARD_ROW_OPERATION, Color.CALCULATION)
@_stop_on_empty_first_cell
def color_column_calc_sht(sht: xw.Sheet):
    """Coloring rules for column calc tab.
    Will assume that any header with the word index in it
    will be colored as an index, any cell with a formula below will
    be colored as a calculation and any cell below without a formula or is None
    is an input.

    Arguments:
        sht {xw.Sheet} -- sheet to color
    """

    header_row = _get_first_sht_row(sht)
    first_value_row = header_row.offset(1, 0)

    for header_cell, value_cell in zip(header_row, first_value_row):

        if str(header_cell.value).lower().startswith('used_'):
            header_cell.color = Color.USED_IDENTIFIER
        elif 'index' in str(header_cell.value).lower():
            header_cell.color = Color.INDEX
        elif value_cell.value is None:
            header_cell.color = Color.INPUT
        elif str(value_cell.formula)[0] != '=':
            header_cell.color = Color.INPUT
        else:
            header_cell.color = Color.CALCULATION


def color_typed_sheet(sht: xw.Sheet, sht_type):
    """Colors a worksheet given that is type has a valid registered color function.

    Arguments:
        sht {xw.Sheet} -- sheet to color
        sht_type {[type]} -- sheet type.
    """

    if sht_type in _SHEET_COLOR_FUNCS:
        _SHEET_COLOR_FUNCS[sht_type](sht)
