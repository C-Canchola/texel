"""
Module to aid in initializing data for sheet tracking.
"""
from ..types import sheet_types


def get_data_labels_for_sheet_addition(sht_type: sheet_types.SheetType, **kwargs):
    """
    Function to initialize row, column, and data labels for two dimensional worksheets.

    Args:
        sht_type ():
        **kwargs ():

    Returns:

    """
    if sht_type != sheet_types.TWO_DIMENSIONAL_LOOK_UP:
        return "na", "na", "na"

    labels = ['row_label', 'col_label', 'data_label']
    row_label, col_label, data_label = [kwargs.get(label, 'untitled_{}'.format(label)) for label in labels]

    return row_label, col_label, data_label
