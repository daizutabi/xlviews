import pytest


@pytest.mark.parametrize(
    ("label", "key", "expected"),
    [
        ("a", {}, "a"),
        ("a{b}", {"b": "B"}, "aB"),
        ("{a}{b}", {"a": "A", "b": "B"}, "AB"),
        (lambda x: f"_{x['a']}_", {"a": "A"}, "_A_"),
    ],
)
def test_format_label(label, key, expected):
    from xlviews.figure.plot import get_label

    assert get_label(label, key) == expected
