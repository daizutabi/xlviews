from importlib.metadata import version


def test_version():
    assert version("xlviews").count(".") == 2
