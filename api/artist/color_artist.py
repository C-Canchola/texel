import xlwings as xw
import texel.api.types.sheet_types as txl_sht_types
from functools import wraps
from texel.constants.color import Color


def _get_first_sht_row(sht: xw.Sheet):
    return sht.range('a1').current_region.options(ndim=1)[0, ::]


def _stop_on_empty_first_cell(func):

    @wraps(func)
    def inner_func(sht, *args, **kwargs):
        if sht.range('a1').value is None:
            return

        return func(sht, *args, **kwargs)
    return inner_func


@_stop_on_empty_first_cell
def color_input_sht(sht: xw.Sheet):

    input_row = _get_first_sht_row(sht)
    input_row.color = Color.INPUT

    input_row[0, 0].color = Color.INDEX


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

        if 'index' in str(header_cell.value).lower():
            header_cell.color = Color.INDEX
        elif value_cell.value is None:
            header_cell.color = Color.INPUT
        elif str(value_cell.formula)[0] != '=':
            header_cell.color = Color.INPUT
        else:
            header_cell.color = Color.CALCULATION
