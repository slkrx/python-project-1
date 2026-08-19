import sys
sys.path.insert(1, "../")
from census_faker_io import s_to_i, dropdown_validate, s_to_b
import pytest

def test_s_to_i_with_string():
    with pytest.raises(ValueError, match="Not an integer"):
        s_to_i("not an int")

def test_s_to_i_with_negative():
    with pytest.raises(ValueError, match="Value must be a positive integer"):
        s_to_i("-1")

def test_s_to_i_valid():
    assert s_to_i("1") == 1

def test_dropdown_validate_with_invalid_input():
    with pytest.raises(ValueError, match="Must be one of the following values: a, b, c"):
        dropdown_validate("d", ["a", "b", "c"])

def test_dropdown_validate_with_valid_input():
    assert dropdown_validate("a", ["a", "b", "c"]) == "a"

def test_s_to_b_with_invalid_input():
    with pytest.raises(ValueError, match="Must enter \"y\" or \"n\""):
        s_to_b("something else")

def test_s_to_b_with_valid_input():
    assert s_to_b("y") == True