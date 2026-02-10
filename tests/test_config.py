from __future__ import annotations

import pytest

from xlviews.config import CONFIG_FILE, rcParams


def test_config_file():
    assert CONFIG_FILE.exists()


def test_rcparams():
    assert rcParams["chart.width"] == 200
    assert rcParams["chart.title.font.bold"] is True

    rcParams["chart.width"] = 100
    assert rcParams["chart.width"] == 100

    rcParams["chart.width"] = 200


def test_rcparams_error():
    with pytest.raises(KeyError):
        rcParams["invalid"]


def test_rcparams_get():
    assert rcParams.get("chart.width", 100) == 200


def test_rcparams_get_default():
    assert rcParams.get("invalid", "default") == "default"
