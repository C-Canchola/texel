import xlwings as xw

from functools import wraps

from ..decorators import stop_on_empty_first_cell as _stop_on_empty_first_cell

from ..types import sheet_types as txl_sht_types

from ...constants import Color
from ...formatting import sheet_fmt

