import xlwings as xw

from functools import wraps

from texel.api.artists import utils

from ..decorators import stop_on_empty_first_cell as _stop_on_empty_first_cell
from ..types import sheet_types as txl_sht_types

from ...constants import Color
from ...formatting import sheet_fmt

_get_first_sht_row = utils.get_first_sht_row
_SHEET_COLOR_FUNCS = {}
test_register_deco = utils.create_sht_type_register_decorator(_SHEET_COLOR_FUNCS)


def color_tab_decorator(color):
    def deco(func):
        @wraps(func)
        def inner_func(sht, *args, **kwargs):
            ret_val = func(sht, *args, **kwargs)
            sheet_fmt.color_sht_tab(sht, color)
            return ret_val

        return inner_func

    return deco


@test_register_deco(txl_sht_types.TWO_DIMENSIONAL_LOOK_UP)
@color_tab_decorator(Color.INDEX)
@_stop_on_empty_first_cell
def color_two_dim_look_up_sheet(sht: xw.Sheet):
    input_row = _get_first_sht_row(sht)
    input_row.color = Color.INDEX


@test_register_deco(txl_sht_types.SCALAR_INPUT)
@color_tab_decorator(Color.INPUT)
@_stop_on_empty_first_cell
def color_input_sht(sht: xw.Sheet):
    input_row = _get_first_sht_row(sht)
    input_row.color = Color.INPUT

    input_row[0, 0].color = Color.INDEX


@test_register_deco(txl_sht_types.STANDARD_ROW_OPERATION)
@color_tab_decorator(Color.CALCULATION)
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
    for header_cell in header_row:
        header_cell.color = utils.get_column_header_color(header_cell)


def color_typed_sheet(sht: xw.Sheet, sht_type):
    """Colors a worksheet given that is type has a valid registered color function.

    Arguments:
        sht {xw.Sheet} -- sheet to color
        sht_type {[type]} -- sheet type.
    """
    if sht_type in _SHEET_COLOR_FUNCS:
        _SHEET_COLOR_FUNCS[sht_type](sht)
