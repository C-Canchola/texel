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
def create_hyperlink_to_other_cell(src_cell: xw.Range, dst_cell: xw.Range, friendly_nm: str):
    """Creates a hyperlink formula in scr_cell to dst_cell with the given
    friendly_nm being the text displayed.

    Arguments:
        src_cell {xw.Range} -- cell which will contain the hyperlink.
        dst_cell {xw.Range} -- cell which the hyperlink will lead to.
        friendly_nm {str} -- text of hyperlink
    """
    link_location = '"#{}"'.format(
        dst_cell.get_address(include_sheetname=True))
    friendly_nm = '"{}"'.format(friendly_nm.replace('"', ''))

    src_cell.formula = '=HYPERLINK({}, {})'.format(link_location, friendly_nm)


# %%
