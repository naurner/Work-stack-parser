import pytest

from channel_parser import parse_views_text


@pytest.mark.parametrize("text, expected", [
    ("987", 987),
    ("1.2K", 1200),
    ("15K", 15000),
    ("3,4M", 3400000),
    ("", 0),
    ("n/a", 0),
])
def test_parse_views_text(text, expected):
    assert parse_views_text(text) == expected
