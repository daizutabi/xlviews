import pytest


def test_wrap_wrap():
    from xlviews.stats import get_wrap

    assert get_wrap(wrap="wrap") == "wrap"
    assert get_wrap(wrap={"a": "wrap"}) == {"a": "wrap"}


@pytest.mark.parametrize(
    ("kwarg", "value"),
    [("na", "IFERROR({},NA())"), ("null", 'IFERROR({},"")')],
)
def test_wrap_true(kwarg, value):
    from xlviews.stats import get_wrap

    assert get_wrap(**{kwarg: True}) == value  # type: ignore


def test_wrap_list():
    from xlviews.stats import get_wrap

    x = get_wrap(na=["a", "b"], null="c")
    assert isinstance(x, dict)
    assert x["a"] == "IFERROR({},NA())"
    assert x["b"] == "IFERROR({},NA())"
    assert x["c"] == 'IFERROR({},"")'


def test_wrap_none():
    from xlviews.stats import get_wrap

    assert get_wrap() == {}
