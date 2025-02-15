import pytest

from xlviews.core.index import Index, WideIndex


@pytest.fixture(scope="module")
def wide_index():
    return WideIndex({"A": [1, 2, 3], "B": [4, 5, 6, 7]})


def test_wide_getitem(wide_index: WideIndex):
    assert wide_index["A"] == [1, 2, 3]
    assert wide_index["B"] == [4, 5, 6, 7]


def test_wide_len(wide_index: WideIndex):
    assert len(wide_index) == 7


def test_wide_names(wide_index: WideIndex):
    assert wide_index.names == ["A", "B"]


def test_wide_to_list(wide_index: WideIndex):
    x = [("A", 1), ("A", 2), ("A", 3), ("B", 4), ("B", 5), ("B", 6), ("B", 7)]
    assert wide_index.to_list() == x


@pytest.mark.parametrize(("key", "loc"), [("A", (0, 3)), ("B", (3, 7))])
def test_wide_loc(wide_index: WideIndex, key, loc):
    assert wide_index.get_loc(key) == loc


def test_wide_append():
    index = WideIndex()
    index.append("A", [1, 2, 3])
    assert "A" in index


def test_wide_append_error():
    index = WideIndex({"A": [1, 2, 3]})
    with pytest.raises(ValueError, match="key 'A' already exists"):
        index.append("A", [1, 2, 3])


@pytest.fixture(scope="module")
def index(wide_index: WideIndex):
    return Index(["x", "y", "z"], wide_index)


def test_index_len(index: Index):
    assert len(index) == 10


def test_index_names(index: Index):
    assert index.names == [None]


def test_index_nlevels(index: Index):
    assert index.nlevels == 1


def test_index_to_list(index: Index):
    x = ["x", "y", "z"]
    y = [("A", 1), ("A", 2), ("A", 3), ("B", 4), ("B", 5), ("B", 6), ("B", 7)]
    assert index.to_list() == x + y


def test_index_append():
    index = Index(["a"])
    index.append("b")
    assert index.to_list() == ["a", "b"]


def test_index_append_wide():
    index = Index(["a"])
    index.append("b", [1, 2, 3])
    assert index.to_list() == ["a", ("b", 1), ("b", 2), ("b", 3)]
