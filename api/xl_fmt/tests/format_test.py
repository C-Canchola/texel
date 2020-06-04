import pytest
from ...xl_fmt import spacing as helpers


def test_remove_whitespace():

    test_str = '=SUM( A:A, A2)'
    assert helpers.remove_white_space(test_str) == '=SUM(A:A,A2)'


def test_add_spacing():
    test_str = '=SUM(A:A, B2)'
    assert helpers.add_arg_and_formula_spacing(test_str) == '=SUM( A:A, B2)'


def test_equal_rng_add_spacing():
    test_str = '=SOME_NAMED_RANGE'
    assert helpers.add_arg_and_formula_spacing(test_str) == test_str


def test_literal():
    test_str = 'LOL'
    assert helpers.add_arg_and_formula_spacing(test_str) == test_str


def test_add_spacing_with_existing_spacing():
    test_str = '=SUM(A:A   , B2        )'
    assert helpers.add_arg_and_formula_spacing(test_str) == '=SUM( A:A, B2)'


def test_add_if_spacing():

    test_str = '=IF(a2,b2,c2)'
    assert helpers.space_if_formulas(test_str) == '=IF(a2,\n   b2,\n   c2)'
