import pytest
from pandas import DataFrame
from xlwings import Range, Sheet

from xlviews.core.formula import NONCONST_VALUE
from xlviews.dataframes.table import Table
from xlviews.testing import is_app_available

pytestmark = pytest.mark.skipif(not is_app_available(), reason="Excel not installed")


@pytest.fixture(scope="module")
def df():
    return DataFrame({"a": [1, 1, 1, 0], "b": [2, 2, 3, 3], "c": [4, 4, 5, 6]})


@pytest.fixture(scope="module")
def rng(df: DataFrame, sheet_module: Sheet):
    rng = sheet_module.range("B3")
    rng.options(DataFrame, header=1, index=False).value = df
    return rng.expand()


def test_range_address(rng: Range):
    assert rng.get_address() == "$B$3:$D$7"


@pytest.fixture(scope="module")
def table(rng: Range):
    return Table(rng)


def test_cell(table: Table):
    assert table.cell.get_address() == "$B$3"


def test_column(table: Table):
    assert table.column.get_address() == "$B$4:$B$7"


def test_header(table: Table):
    assert table.header.get_address() == "$B$3:$D$3"


def test_const_header(table: Table):
    assert table.const_header.get_address() == "$B$2:$D$2"


def test_columns(table: Table):
    assert table.columns == ["a", "b", "c"]


def test_add_const_header(table: Table):
    table.add_const_header()
    assert table.const_header.value == [NONCONST_VALUE, NONCONST_VALUE, NONCONST_VALUE]


def test_add_const_header_clear(table: Table):
    table.add_const_header(clear=True)
    assert table.const_header.value == [None, None, None]


def test_from_api(table: Table):
    api = list(table.cell.sheet.api.ListObjects)[0]
    table2 = Table(api=api, sheet=table.sheet)
    assert table2.cell.row == table.cell.row
    assert table2.cell.column == table.cell.column


def test_unlist(df: DataFrame, sheet: Sheet):
    rng = sheet.range("B100")
    rng.options(DataFrame, header=1, index=False).value = df
    table = Table(rng.expand())
    assert len(list(sheet.api.ListObjects)) == 1
    table.unlist()
    assert not list(sheet.api.ListObjects)


@pytest.mark.parametrize(
    ("name", "value", "const"),
    [
        ("a", 1, [1, NONCONST_VALUE, NONCONST_VALUE]),
        ("a", 0, [0, 3, 6]),
        ("b", 2, [1, 2, 4]),
        ("b", 3, [NONCONST_VALUE, 3, NONCONST_VALUE]),
        ("c", 4, [1, 2, 4]),
        ("c", 5, [1, 3, 5]),
        ("c", 6, [0, 3, 6]),
        ("c", (4, 5), [1, NONCONST_VALUE, NONCONST_VALUE]),
        ("c", [5, 4], [1, NONCONST_VALUE, NONCONST_VALUE]),
        ("c", [4, 6], [NONCONST_VALUE, NONCONST_VALUE, NONCONST_VALUE]),
    ],
)
@pytest.mark.parametrize("to_dict", [True, False])
def test_auto_filter(table: Table, name, value, const, to_dict):
    table.add_const_header()

    if to_dict:
        table.auto_filter({name: value})
    else:
        table.auto_filter(name, value)

    assert table.const_header.value == const
    table.auto_filter(clear=True)
    table.add_const_header(clear=True)


def test_table_error():
    msg = "Either range or sheet and api must be provided"
    with pytest.raises(ValueError, match=msg):
        Table()
