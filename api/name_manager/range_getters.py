import xlwings as xw
from ... import naming as txl_nm
from ..types import sheet_types as txl_sht_types

from functools import wraps

SCALAR_INPUT = txl_sht_types.SCALAR_INPUT
STANDARD_ROW_OPERATION = txl_sht_types.STANDARD_ROW_OPERATION
TWO_DIMENSIONAL_LOOK_UP = txl_sht_types.TWO_DIMENSIONAL_LOOK_UP

VALID_TYPES = [SCALAR_INPUT, STANDARD_ROW_OPERATION, TWO_DIMENSIONAL_LOOK_UP]

NAMED_RANGE_NAME_FUNCS = {}
NAMED_RANGE_ADDRESS_FUNCS = {}


def register_type_func(sht_type, func_dict):
    def decorator(func):
        func_dict[sht_type] = func
        return func

    return decorator


def stop_if_row_empty(func):
    @wraps(func)
    def inner_func(sht: xw.Sheet):
        if len(sht.range('a1').current_region.rows) < 2:
            return []

        return func(sht)

    return inner_func


def stop_if_col_empty(func):
    @wraps(func)
    def inner_func(sht: xw.Sheet):
        if sht.range('A1').value is None:
            return []

        return func(sht)

    return inner_func


@register_type_func(SCALAR_INPUT.index, NAMED_RANGE_ADDRESS_FUNCS)
@stop_if_row_empty
def get_scalar_ranges_to_name(sht: xw.Sheet):
    """Returns a list containing the address of each filled row in the value (2nd) column
    of a scalar sheet.

    Arguments:
        sht {xw.Sheet} -- scalar input sheet.
    """

    stop = len(sht.range('a1').current_region.rows) + 1
    sht_nm = sht.name
    return ['{}!$B${}'.format(sht_nm, i) for i in range(2, stop)]


@register_type_func(SCALAR_INPUT.index, NAMED_RANGE_NAME_FUNCS)
@stop_if_row_empty
def get_scalar_names(sht: xw.Sheet):
    """Returns a list of 

    Arguments:
        sht {xw.Sheet} -- [description]

    Returns:
        [type] -- [description]
    """
    values = sht.range('a1').current_region[1:, 0].options(ndim=1).value
    sht_nm = sht.name

    return ['{}__{}'.format(sht_nm, value) for value in values]


@register_type_func(TWO_DIMENSIONAL_LOOK_UP.index, NAMED_RANGE_ADDRESS_FUNCS)
@stop_if_row_empty
@stop_if_col_empty
def get_two_dimensional_look_up_ranges_to_name(sht: xw.Sheet):
    entire_rng: xw.Range = sht.range('A1').current_region
    row_rng = entire_rng[1:, 0]
    header_rng = entire_rng[0, 1:]
    data_rng = entire_rng[1:, 1:]
    return row_rng.get_address(include_sheetname=True), \
           header_rng.get_address(include_sheetname=True), data_rng.get_address(include_sheetname=True)


@register_type_func(STANDARD_ROW_OPERATION.index, NAMED_RANGE_ADDRESS_FUNCS)
@stop_if_col_empty
def get_standard_row_ranges_to_name(sht: xw.Sheet):
    stop = len(sht.range('A1').current_region.columns) + 1
    sht_nm = sht.name

    return ['{sht_nm}!${col_letter}:${col_letter}'.format(
        sht_nm=sht_nm, col_letter=xw.utils.col_name(i))
        for i in range(1, stop)]


@register_type_func(STANDARD_ROW_OPERATION.index, NAMED_RANGE_NAME_FUNCS)
@stop_if_col_empty
def get_standard_row_names(sht: xw.Sheet):
    values = sht.range('A1').current_region[0, ::].options(ndim=1).value
    sht_nm = sht.name

    return ['{}__{}'.format(sht_nm, value) for value in values]


def _get_label_name(sht_nm: str, label_type: str, label: str):
    return '{}__{}__{}'.format(sht_nm, label_type, label)


@register_type_func(TWO_DIMENSIONAL_LOOK_UP, NAMED_RANGE_NAME_FUNCS)
@stop_if_col_empty
@stop_if_row_empty
def get_two_dimensional_row_names(sht: xw.Sheet, row_label: str, col_label: str, data_label: str):
    sht_nm: str = sht.name
    return (_get_label_name(sht_nm, 'ROW', row_label),
            _get_label_name(sht_nm, 'COL', col_label),
            _get_label_name(sht_nm, 'DATA', data_label))


def get_names_and_addresses(sht: xw.Sheet, sht_type: txl_sht_types.SheetType, **kwargs):
    return NAMED_RANGE_NAME_FUNCS[sht_type](sht, **kwargs), NAMED_RANGE_ADDRESS_FUNCS[sht_type](sht)
