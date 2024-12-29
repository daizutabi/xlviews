"""Read the configuration file.

The configuration parameters are described in `xlviews.toml`
located in the same directory.

Custom settings can be made by directly modifying `xlviews.config.rcParams`.
"""

from __future__ import annotations

import tomllib
from pathlib import Path
from typing import Any

CONFIG_FILE = Path(__file__).parent / "xlviews.toml"


def load_config() -> dict[str, Any]:
    with CONFIG_FILE.open("rb") as f:
        return tomllib.load(f)


class Config:
    params: dict[str, Any]

    def __init__(self) -> None:
        self.params = load_config()

    def __getitem__(self, key: str) -> Any:  # noqa: ANN401
        keys = key.split(".")
        value = self.params

        for k in keys:
            value = value[k]

        return value


rcParams = Config()  # noqa: N816
