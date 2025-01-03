def test_config_file():
    from xlviews.config import CONFIG_FILE

    assert CONFIG_FILE.exists()


def test_rcParams():  # noqa: N802
    from xlviews.config import rcParams

    assert rcParams["chart.width"] == 200
    assert rcParams["chart.title.font.bold"] is True

    rcParams["chart.width"] = 100
    assert rcParams["chart.width"] == 100

    rcParams["chart.width"] = 200
