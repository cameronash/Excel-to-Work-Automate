from run_value_into_word import format_number_as_words, format_number


def test_format_number_as_words():
    assert format_number_as_words(6800000) == "(Six Million Eight Hundred Thousand Dollars)"


def test_format_number_default():
    assert format_number(1500, None) == "1,500"


def test_format_number_custom():
    assert format_number(123.456, "{:.2f}") == "123.46"

