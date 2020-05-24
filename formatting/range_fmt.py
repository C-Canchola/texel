import xlwings as xw

CONSTANTS_HALIGN = xw.constants.HAlign


def set_rng_font_color(rng: xw.Range, color_val):
    rng.api.font.color = color_val


def set_rng_halign(rng: xw.Range, halign):
    """Sets horizontal alignment of a range.


    Arguments:
        rng {xw.Range} -- range to set horizontal alignment.
        halign {[type]} -- CONSTANTS_HALIGN constant.
    """
    rng.api.HorizontalAlignment = halign


def set_bold_font(rng: xw.Range, bold_val: bool):

    rng.api.font.bold = bold_val
