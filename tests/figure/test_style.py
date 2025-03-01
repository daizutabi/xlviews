import pytest


@pytest.mark.parametrize(
    ("style", "name", "expected"),
    [
        (None, None, None),
        ("a", None, "a"),
        ("a", "name", "a"),
        ("a{}", "name", "aname"),
        ("{}a{}", ("A", "B"), "AaB"),
        ("{1}a{0}", ("A", "B"), "BaA"),
        (lambda x: f"_{x}_", "A", "_A_"),
        (lambda x: f"_{x}_", ("A", "B"), "_('A', 'B')_"),
        (lambda x: f"_{x}_", None, "_None_"),
    ],
)
def test_format_style(style, name, expected):
    from xlviews.figure.style import format_label

    assert format_label(style, name) == expected
