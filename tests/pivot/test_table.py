from pandas import DataFrame


def test_df(df: DataFrame):
    print(df)
    pt = df.pivot_table(index="x", columns="y", values="c")
    print(pt)
    # assert 0
