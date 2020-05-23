import pytest
import os
import xlwings as xw
import texel.utils.book as book
import texel.utils.decorators as dec
import texel.utils.constants as constants


def test_open_book_from_path():
    bk = book.get_bk(constants.TEST_FILE_PATH)

    assert isinstance(bk, xw.Book)

    


def test_open_by_from_name():

    path_bk = book.get_bk(constants.TEST_FILE_PATH)

    nm = os.path.split(constants.TEST_FILE_PATH)[-1]

    nm_bk = book.get_bk(nm)

    assert isinstance(nm_bk, xw.Book)

    


def test_get_by_path_twice():

    path_bk = book.get_bk(constants.TEST_FILE_PATH)
    sec_bk = book.get_bk(constants.TEST_FILE_PATH)

    assert path_bk.name == sec_bk.name
