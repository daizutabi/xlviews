import pandas as pd
import pytest
from pandas import DataFrame


@pytest.fixture(scope="module")
def df_unit():
    df1 = DataFrame(
        {
            "x": [1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3],
            "y": [1, 2, 3, 4] * 3,
            "c": range(1, 13),
        },
    )
    df2 = df1.copy()
    df2["c"] = df2["c"] * 2
    df3 = df1.copy()
    df3["c"] = df3["c"] * 4
    df_unit = pd.concat([df1, df2, df3])
    return df_unit.sort_values(by=["x", "y"])


@pytest.fixture(scope="module")
def df(df_unit: DataFrame):
    dfs = []
    k = 0
    for a in ["s", "t"]:
        for b in [10, 20]:
            df = df_unit.copy()
            df["a"] = a
            df["b"] = b
            df["c"] += k
            k += 1
            dfs.append(df)

    df = pd.concat(dfs, ignore_index=True)
    df = df[["a", "b", "x", "y", "c"]]
    df["d"] = df["c"] * 2

    return df.sort_values(by=["a", "b", "x", "y"])
