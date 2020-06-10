import xlwings as xw
from ..xl_fmt import get_spaced_formula
from ..types import sheet_types as txl_sht_types
from ..decorators import stop_on_empty_first_cell

_TYPED_FORMULA_FORMATTING_FUNCTIONS = {}


def format_typed_sheet_formulas(sht: xw.Sheet, sht_type):
    """Formats the formulas of a typed sheet based of the formatting function that is registered under
    its type.

    Arguments:
        sht {xw.Sheet} -- sheet to format formulas.
        sht_type {[type]} -- sheet type.
    """

    if sht_type in _TYPED_FORMULA_FORMATTING_FUNCTIONS:
        _TYPED_FORMULA_FORMATTING_FUNCTIONS[sht_type](sht)


def register_sheet_formula_formatting_func(sht_type):
    def decorator(func):
        _TYPED_FORMULA_FORMATTING_FUNCTIONS[sht_type] = func
        return func

    return decorator


@register_sheet_formula_formatting_func(txl_sht_types.SCALAR_INPUT)
def _do_nothing(*args, **kwargs):
    return


@register_sheet_formula_formatting_func(txl_sht_types.STANDARD_ROW_OPERATION)
@stop_on_empty_first_cell
def space_first_row_of_sheet(sht: xw.Sheet):
    total_rng: xw.Range = sht.range('a1').current_region
    format_row: xw.Range = total_rng[1, ::]

    cell: xw.Range
    for cell in format_row:
        apply_spaced_format_to_cell(cell)


def apply_spaced_format_to_cell(cell: xw.Range):
    spaced_formula: str = get_spaced_formula(cell.formula)
    if rng_is_array(cell):
        cell.formula_array = spaced_formula
    else:
        cell.formula = spaced_formula


def apply_spacing_to_rng(rng: xw.Range):

    for cell in rng:
        apply_spaced_format_to_cell(cell)


def rng_is_array(rng: xw.Range):
    """
    Returns whether the given range contains an array formula.
    This will be useful for deciding on how to set the formulas of cells.
    Args:
        rng (xw.Range): range to check.

    Returns: bool

    """
    return rng.api.HasArray
