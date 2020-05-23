""" Helper functions for working with excel ranges.
"""
# %%
import xlwings as xw


# %%
def get_entire_sheet_column(sht: xw.Sheet, col_index: int) -> xw.Range:
    """Returns an entire column range from a sheet and column (0-based)
    number.

    Arguments:
        sht {xw.Sheet} -- worksheet
        col_index {int} -- col index - 0-based

    Returns:
        xw.Range -- Entire column.
    """

    return sht.cells[::, col_index]


# %%
