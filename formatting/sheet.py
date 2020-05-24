import xlwings as _xw


def color_sht_tab(sht: _xw.Sheet, color_val):
    """Colors the tab of an xw.Sheet

    Arguments:
        sht {xw.Sheet} -- Sheet to color
        color_val {[type]} -- Color to set sheet.
    """
    sht.api.Tab.Color = color_val


def get_sht_tab_color(sht: _xw.Sheet):
    """Returns the color value of a sheet's tab.

    Arguments:
        sht {xw.Sheet} -- Sheet to get color of

    Returns:
        [type] -- Color value
    """
    return sht.api.Tab.Color
