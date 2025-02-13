def test_wide_index():
    from xlviews.core.index import WideIndex

    index = WideIndex({"A": [1, 2, 3], "B": [4, 5, 6]})

    assert index["A"] == [1, 2, 3]
    assert index["B"] == [4, 5, 6]
