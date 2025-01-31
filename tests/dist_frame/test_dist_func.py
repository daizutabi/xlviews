def test_dist_func_str():
    from xlviews.dataframes.dist import get_dist_func

    df = get_dist_func("norm", ["a", "b"])
    assert df == {"a": "norm", "b": "norm"}


def test_dist_func_dict():
    from xlviews.dataframes.dist import get_dist_func

    df = get_dist_func({"a": "none"}, ["a", "b"])
    assert df == {"a": "none", "b": "norm"}
