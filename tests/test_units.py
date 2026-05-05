from pella_order_automation.units import parse_dimension_pair, parse_number


def test_parse_number_handles_mixed_fraction_with_dash():
    assert parse_number("3-9/16") == 3.5625


def test_parse_number_handles_mixed_fraction_with_space():
    assert parse_number("95 1/2") == 95.5


def test_parse_dimension_pair():
    assert parse_dimension_pair("75 X 95 1/2") == (75, 95.5)
