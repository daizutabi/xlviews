import numpy as np
import pytest
from pandas import DataFrame, Series

from xlviews.grouper import _to_dict


@pytest.mark.parametrize(
    ("keys", "values", "expected"),
    [
        ([1, 2, 1], ["a", "b", "c"], {1: ["a", "c"], 2: ["b"]}),
        (["x", "y", "x"], [10, 20, 30], {"x": [10, 30], "y": [20]}),
        ([True, False, True], [1.1, 2.2, 3.3], {True: [1.1, 3.3], False: [2.2]}),
        ([], [], {}),
        ([None, None], ["a", "b"], {None: ["a", "b"]}),
    ],
)
def test_to_dict(keys, values, expected):
    assert _to_dict(keys, values) == expected


@pytest.mark.parametrize("func", [lambda x: x, np.array, Series])
def test_create_group_index_series(func):
    from xlviews.grouper import create_group_index

    values = [1, 1, 2, 2, 2, 3, 3, 1, 1, 2, 2, 3, 3]
    index = create_group_index(func(values))
    assert index[(1,)] == [(0, 1), (7, 8)]
    assert index[(2,)] == [(2, 4), (9, 10)]
    assert index[(3,)] == [(5, 6), (11, 12)]


@pytest.mark.parametrize("func", [lambda x: x, DataFrame])
def test_create_group_index_dataframe(func):
    from xlviews.grouper import create_group_index

    values = [[1, 2], [1, 2], [3, 4], [3, 4], [1, 2], [3, 4], [3, 4]]
    index = create_group_index(func(values))
    assert index[(1, 2)] == [(0, 1), (4, 4)]
    assert index[(3, 4)] == [(2, 3), (5, 6)]
