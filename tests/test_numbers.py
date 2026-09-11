from decimal import Decimal

import pytest

from app.services.pdf_parser import parse_ptbr_decimal


@pytest.mark.parametrize(("raw", "expected"), [("75.000,000", "75000.000"), ("1.155.000,000", "1155000.000"), ("781,20", "781.20"), ("0,0104160", "0.0104160")])
def test_ptbr_decimal_parser(raw, expected):
    assert parse_ptbr_decimal(raw) == Decimal(expected)
