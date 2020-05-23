import xlwings as xw
import texel.naming as txl_nm
import texel.api.types.sheet_types as txl_sht_types

from functools import wraps

SCALAR_INPUT = txl_sht_types.SCALAR_INPUT
VALID_TYPES = [SCALAR_INPUT]



def stop_if_empty(func):

    @wraps(func)
    def inner_func(sht: xw.Sheet):

        if len(sht.range('a1').current_region.rows) < 2:
            return []

        return func(sht)

    return inner_func


@stop_if_empty
def get_scalar_ranges_to_name(sht: xw.Sheet):
    """Returns a list containing the address of each filled row in the value (2nd) column
    of a scalar sheet.

    Arguments:
        sht {xw.Sheet} -- scalar input sheet.
    """

    stop = len(sht.range('a1').current_region.rows) + 1
    sht_nm = sht.name
    return ['{}!$B${}'.format(sht_nm, i) for i in range(2, stop)]


@stop_if_empty
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


